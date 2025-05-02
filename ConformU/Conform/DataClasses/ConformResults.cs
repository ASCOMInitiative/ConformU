using System.Collections.Generic;

namespace ConformU
{
    public class ConformResults
    {
        public ConformResults()
        {
            // Initialise the class
            Errors = new();
            Issues = new();
            ConfigurationAlerts= new();
            Timings = new();
            TimingIssuesCount = 0;
            TimingCount = 0;
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

        public int ConfigurationAlertCount
        {
            get
            {
                return ConfigurationAlerts.Count;
            }
        }

        public int TimingIssuesCount { get; set; }

        public int TimingCount { get; set; }

        public List<KeyValuePair<string, string>> Errors { get; set; }

        public List<KeyValuePair<string, string>> Issues { get; set; }

        public List<KeyValuePair<string, string>> ConfigurationAlerts { get; set; }

        public List<KeyValuePair<string, string>> Timings { get; set; }
    }
}
