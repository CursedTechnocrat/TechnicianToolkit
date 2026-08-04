using TechnicianToolkit.Core.Html;
using Xunit;

namespace TechnicianToolkit.Core.Tests;

public class TkHtmlTests
{
    [Fact]
    public void Head_EmitsDoctypeTitleAndOpensMain()
    {
        var head = TkHtml.Head(title: "My Report", scriptName: "S.C.R.Y.E.R.");

        Assert.StartsWith("<!DOCTYPE html>", head);
        Assert.Contains("<title>My Report</title>", head);
        Assert.Contains("S.C.R.Y.E.R. REPORT", head);
        Assert.Contains("tk-main", head);
        Assert.Contains(TkHtmlTheme.Css, head);
    }

    [Fact]
    public void Head_RendersNavItemsWithTwoDigitNumbers()
    {
        var head = TkHtml.Head(navItems: new[] { "Alpha", "Beta" });

        Assert.Contains("href='#s01'", head);
        Assert.Contains("href='#s02'", head);
        Assert.Contains("Alpha", head);
        Assert.Contains("Beta", head);
    }

    [Fact]
    public void Head_RendersMetaItems()
    {
        var head = TkHtml.Head(metaItems: new[]
        {
            new KeyValuePair<string, string>("Generated", "2026-08-04 15:30"),
            new KeyValuePair<string, string>("OS", "Windows 11"),
        });

        Assert.Contains("Generated", head);
        Assert.Contains("2026-08-04 15:30", head);
        Assert.Contains("Windows 11", head);
    }

    [Fact]
    public void Head_EscapesMetaValues()
    {
        var head = TkHtml.Head(metaItems: new[]
        {
            new KeyValuePair<string, string>("Owner", "<script>"),
        });

        Assert.Contains("&lt;script&gt;", head);
        Assert.DoesNotContain("<script>", head);
    }

    [Fact]
    public void Foot_ClosesDocument()
    {
        var foot = TkHtml.Foot("A.U.G.U.R. v3.6");

        Assert.Contains("A.U.G.U.R. v3.6", foot);
        Assert.Contains("</body>", foot);
        Assert.Contains("</html>", foot);
    }
}
