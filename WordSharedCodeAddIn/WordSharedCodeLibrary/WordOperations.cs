namespace WordSharedCodeLibrary
{
    public class WordOperations
    {
        static public string GetHighlightedTextUnicode(string value)
        {
            string result = "";
            foreach (char c in value)
            {
                int unicode = c;

                result += $"{c}: {unicode}\r\n";
            }
            return result;
        }
    }
}
