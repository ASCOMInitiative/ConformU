using System.Collections.Generic;

namespace ConformU
{
    public class ConformResults
    {
        public ConformResults()
        {
            Errors = new();
            Issues = new();
        }

        public int ErrorCount
        {
            get
            {
                return Errors.Count;
            }
        }

        public int IssueCount
        {
            get
            {
                return Issues.Count;
            }
        }

        // public int OkCount { get; set; }
        public List<KeyValuePair<string, string>> Errors { get; set; }
        public List<KeyValuePair<string, string>> Issues { get; set; }
    }
}
