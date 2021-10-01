namespace ConformU
{
    public class ComDevice
    {
        public ComDevice()
        {
            DisplayName = "";
            ProgId = "";
        }
        public ComDevice(string displayName, string progId)
        {
            DisplayName = displayName;
            ProgId = progId;
        }

        public string DisplayName { get; set; }
        public string ProgId { get; set; }
    }
}