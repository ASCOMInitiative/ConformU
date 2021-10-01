namespace ConformU
{
    public class TelescopeTestItem
    {
        public TelescopeTestItem(string testDescription, int id, bool enabled)
        {
            TestDescription = testDescription;
            Id = id;
            Enabled = enabled;
        }

        public string TestDescription { get; set; }
        public int Id { get; set; }
        public bool Enabled { get; set; }
    }
}
