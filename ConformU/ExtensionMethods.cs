using ASCOM.Tools;
using Blazorise;
using System;

namespace ConformU
{
    public static class ExtensionMethods
    {
        // Constants
        public const int COLUMN_WIDTH = 30; //
        public const int OUTCOME_WIDTH = 6;

        public static bool IsNumeric(this string text) => !string.IsNullOrWhiteSpace(text) && double.TryParse(text, out _);
        public static string SpaceDup(this int n)
        {
            return new String(' ', n);
        }
        public static string ToMultiLine(this string message, int offset, int screenLogColumns)
        {
            // Reformat the text across multiple screen lines if the message does not contain HTML code. If it does contain HTML display as-is.
            int maxLineLength = screenLogColumns - COLUMN_WIDTH - OUTCOME_WIDTH
                ;
            string padString = new(' ', COLUMN_WIDTH + OUTCOME_WIDTH + offset);
            if (screenLogColumns > 0)
            {
                // Leave HTML content unchanged
                if (!message.ToLowerInvariant().Contains("<!doctype html>")) // Not HTML
                {
                    string messageTrimmed = message.Trim();
                    // Convert to multiple lines if message length is over 1 line length
                    if (messageTrimmed.Length > maxLineLength)
                    {
                        // Trim white space from the message before we start
                        string messageMultiLines = "";

                        //Console.WriteLine("012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789 Length > maxLineLength");
                        //Console.WriteLine(messageTrimmed);

                        int thisBreakPosition = 0;
                        int lastBreakPosition = 0;
                        while (messageTrimmed.Length > lastBreakPosition + maxLineLength)
                        {
                            thisBreakPosition = messageTrimmed.LastIndexOf(" ", lastBreakPosition + maxLineLength);
                            string thisLine = messageTrimmed.Substring(lastBreakPosition, thisBreakPosition - lastBreakPosition).Trim();
                            messageMultiLines = $"{messageMultiLines}\r\n{padString}{thisLine}";
                            //Console.WriteLine("012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789 Loop");
                            //Console.WriteLine(messageTrimmed.Substring(lastBreakPosition, thisBreakPosition - lastBreakPosition));
                            //Console.WriteLine($"Screen log columns: {screenLogColumns}, Max line length: {maxLineLength}, This break position: {thisBreakPosition}, Last break position: {lastBreakPosition}, Message length: {message.Length}");
                            //Console.WriteLine("012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789 Loop");
                            //Console.WriteLine(messageMultiLines);
                            lastBreakPosition = thisBreakPosition;
                        }

                        //Console.WriteLine("012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789 Final");
                        message = $"{messageMultiLines}\r\n{padString}{messageTrimmed.Substring(thisBreakPosition + 1)}"; // Add the remaining characters
                    }
                }
            }
            return message;
        }

        public static string ToHMS(this double rightAscension)
        {
            return $"{(rightAscension >= 0.0 ? "+" : "")}{Utilities.HoursToHMS(rightAscension, ":", ":", "", 2)}";
        }

        public static string ToDMS(this double declination)
        {
            return $"{(declination >= 0.0 ? "+" : "")}{Utilities.DegreesToDMS(declination, ":", ":", "", 1)}";
        }


    }
}
