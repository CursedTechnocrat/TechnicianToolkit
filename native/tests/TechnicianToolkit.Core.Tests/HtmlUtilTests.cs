using TechnicianToolkit.Core.Html;
using Xunit;

namespace TechnicianToolkit.Core.Tests;

public class HtmlUtilTests
{
    [Fact]
    public void Esc_EscapesTheFourReservedChars()
    {
        Assert.Equal("a &amp; b &lt; c &gt; d &quot; e", HtmlUtil.Esc("a & b < c > d \" e"));
    }

    [Fact]
    public void Esc_LeavesSingleQuoteUntouched_MatchingModule()
    {
        // The PowerShell EscHtml deliberately does not escape single quotes.
        Assert.Equal("it's fine", HtmlUtil.Esc("it's fine"));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public void Esc_NullOrEmpty_ReturnsEmpty(string? input)
    {
        Assert.Equal(string.Empty, HtmlUtil.Esc(input));
    }

    [Fact]
    public void Esc_AmpersandFirst_DoesNotDoubleEscape()
    {
        // Ampersand must be replaced before the entities it would otherwise mangle.
        Assert.Equal("&lt;&amp;&gt;", HtmlUtil.Esc("<&>"));
    }

    [Theory]
    [InlineData(0, "0 B")]
    [InlineData(512, "512 B")]
    [InlineData(1024, "1.00 KB")]
    [InlineData(1536, "1.50 KB")]
    [InlineData(1572864, "1.50 MB")]
    [InlineData(1610612736, "1.50 GB")]
    public void FormatBytes_Matches(long bytes, string expected)
    {
        Assert.Equal(expected, HtmlUtil.FormatBytes(bytes));
    }
}
