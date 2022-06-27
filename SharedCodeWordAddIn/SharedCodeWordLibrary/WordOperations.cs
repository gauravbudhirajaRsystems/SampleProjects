using System.Collections.Generic;

namespace SharedCodeWordLibrary
{
    public class WordOperations
    {
        static public string GetUnicode(string value)
        {
            string result = "";
            foreach (char c in value)
            {
                int unicode = c;

                result += $"{c}: {unicode}\r\n";
            }
            return result;
        }

        static public string GetCharCount(string value)
        {
            string result = "";
            Dictionary<char, int> keyValuePairs = new Dictionary<char, int>();
            foreach (char c in value)
            {
                if (keyValuePairs.ContainsKey(c))
                    keyValuePairs[c]++;
                else
                    keyValuePairs.Add(c, 1);
            }

            foreach (var keyValuePair in keyValuePairs)
                result += $"{keyValuePair.Key}: {keyValuePair.Value}\r\n";

            return result;
        }

        static public string GetWordCount(string value)
        {
            string result = "";
            string[] words = value.Split(' ');
            Dictionary<string, int> keyValuePairs = new Dictionary<string, int>();
            foreach (string word in words)
            {
                if (keyValuePairs.ContainsKey(word))
                    keyValuePairs[word]++;
                else
                    keyValuePairs.Add(word, 1);
            }

            foreach (var keyValuePair in keyValuePairs)
                result += $"{keyValuePair.Key}: {keyValuePair.Value}\r\n";

            return result;
        }
    }
}
