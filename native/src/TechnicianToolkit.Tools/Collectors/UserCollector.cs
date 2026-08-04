using System.Collections;
using System.DirectoryServices;

namespace TechnicianToolkit.Tools.Collectors;

/// <summary>A local user account, mirroring the fields WARD/SCRYER read.</summary>
public sealed class UserAccount
{
    public string Name { get; init; } = "";
    public string FullName { get; init; } = "";
    public bool Enabled { get; init; }
    public bool IsAdmin { get; init; }
    public DateTime? LastLogon { get; init; }
    public DateTime? PasswordLastSet { get; init; }
    public DateTime? PasswordExpires { get; init; }
    public bool PasswordRequired { get; init; }
    public string Description { get; init; } = "";
}

/// <summary>
/// Enumerates local user accounts and Administrators-group membership via the
/// WinNT ADSI provider — the native stand-in for <c>Get-LocalUser</c> plus
/// <c>Get-LocalGroupMember -Group Administrators</c>.
/// </summary>
/// <remarks>
/// One deliberate approximation vs. the PowerShell <c>LocalUser</c> object:
/// <c>PasswordExpires</c> is computed as <c>PasswordLastSet + MaxPasswordAge</c>
/// (machine policy) unless the account has UF_DONT_EXPIRE_PASSWD, rather than
/// read from a dedicated field. When policy has no maximum age, expiry is null
/// ("Never / No Expiry"), matching how WARD renders a null.
/// </remarks>
public static class UserCollector
{
    private const int UfAccountDisable = 0x0002;
    private const int UfPasswdNotReqd = 0x0020;
    private const int UfDontExpirePasswd = 0x10000;

    /// <summary>
    /// Returns the bare account names (domain/machine prefix stripped) that are
    /// members of the local Administrators group. Never throws — an empty list
    /// is returned on any failure, matching WARD's <c>Get-AdminMembers</c>.
    /// </summary>
    public static IReadOnlyList<string> GetAdminMembers()
    {
        var names = new List<string>();
        try
        {
            using var group = new DirectoryEntry($"WinNT://{Environment.MachineName}/Administrators,group");
            if (group.Invoke("Members") is IEnumerable members)
            {
                foreach (var member in members)
                {
                    using var entry = new DirectoryEntry(member);
                    var name = entry.Name;
                    if (!string.IsNullOrEmpty(name))
                    {
                        // Keep only the bare account name after any backslash.
                        var idx = name.LastIndexOf('\\');
                        names.Add(idx >= 0 ? name[(idx + 1)..] : name);
                    }
                }
            }
        }
        catch
        {
            return Array.Empty<string>();
        }

        return names;
    }

    /// <summary>Enumerates all local user accounts with derived password/logon fields.</summary>
    public static IReadOnlyList<UserAccount> Collect(IReadOnlyList<string>? adminNames = null)
    {
        adminNames ??= GetAdminMembers();
        var adminSet = new HashSet<string>(adminNames, StringComparer.OrdinalIgnoreCase);

        var result = new List<UserAccount>();

        using var computer = new DirectoryEntry($"WinNT://{Environment.MachineName},computer");

        // Machine max-password-age policy (seconds) for expiry computation.
        long maxPwdAgeSeconds = 0;
        try
        {
            var raw = computer.Properties["MaxPasswordAge"]?.Value;
            if (raw is not null)
            {
                maxPwdAgeSeconds = Convert.ToInt64(raw);
            }
        }
        catch
        {
            maxPwdAgeSeconds = 0;
        }

        foreach (DirectoryEntry child in computer.Children)
        {
            using (child)
            {
                if (!string.Equals(child.SchemaClassName, "User", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                var name = child.Name;
                var userFlags = ReadInt(child, "UserFlags") ?? 0;
                var enabled = (userFlags & UfAccountDisable) == 0;
                var passwordRequired = (userFlags & UfPasswdNotReqd) == 0;
                var dontExpire = (userFlags & UfDontExpirePasswd) != 0;

                DateTime? lastLogon = ReadDate(child, "LastLogin");

                DateTime? passwordLastSet = null;
                var pwdAge = ReadInt(child, "PasswordAge");
                if (pwdAge is > 0)
                {
                    passwordLastSet = DateTime.Now.AddSeconds(-pwdAge.Value);
                }

                DateTime? passwordExpires = null;
                if (!dontExpire && maxPwdAgeSeconds > 0 && passwordLastSet is not null)
                {
                    passwordExpires = passwordLastSet.Value.AddSeconds(maxPwdAgeSeconds);
                }

                result.Add(new UserAccount
                {
                    Name = name,
                    FullName = ReadString(child, "FullName"),
                    Description = ReadString(child, "Description"),
                    Enabled = enabled,
                    PasswordRequired = passwordRequired,
                    LastLogon = lastLogon,
                    PasswordLastSet = passwordLastSet,
                    PasswordExpires = passwordExpires,
                    IsAdmin = adminSet.Contains(name),
                });
            }
        }

        return result;
    }

    private static string ReadString(DirectoryEntry entry, string prop)
    {
        try
        {
            return entry.Properties[prop]?.Value?.ToString() ?? "";
        }
        catch
        {
            return "";
        }
    }

    private static int? ReadInt(DirectoryEntry entry, string prop)
    {
        try
        {
            var v = entry.Properties[prop]?.Value;
            return v is null ? null : Convert.ToInt32(v);
        }
        catch
        {
            return null;
        }
    }

    private static DateTime? ReadDate(DirectoryEntry entry, string prop)
    {
        try
        {
            var v = entry.Properties[prop]?.Value;
            if (v is null)
            {
                return null;
            }

            var dt = Convert.ToDateTime(v);
            // WinNT reports a sentinel (year 1600) for accounts that never logged
            // on; treat anything implausibly old as "never".
            return dt.Year < 1980 ? null : dt;
        }
        catch
        {
            return null;
        }
    }
}
