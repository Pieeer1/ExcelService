using System.Text.RegularExpressions;

namespace ExcelService.Enums
{
    public static class EnumExtensions
    {
        public static string FontEnumToSpacedString(this Font font)
        {
            Regex re = new Regex(@"([A-Z]+[a-z]+)", RegexOptions.Multiline);
            return string.Join(' ', re.Matches(font.ToString()).Select(x => x.Groups[1].Value));
        }
        public static string? FontEnumToSpacedString(this Font? font)
        {
            if (font is null) { return null; }
            Regex re = new Regex(@"([A-Z]+[a-z]+)", RegexOptions.Multiline);
            return string.Join(' ', re.Matches(font.ToString()!).Select(x => x.Groups[1].Value));
        }
    }
}
