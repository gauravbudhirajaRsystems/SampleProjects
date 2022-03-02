﻿namespace CellAnalyzerSharedLibrary
{
    public class CellOperations
    {
        static public string GetUnicodeFromText(string value)
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
