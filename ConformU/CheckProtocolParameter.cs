namespace ConformU
{
    public class CheckProtocolParameter
    {
        public CheckProtocolParameter(string parameterName, string parameterValue)
        {
            ParameterName = parameterName;
            ParameterValue = parameterValue;
        }


        public string ParameterName { get; set; } = string.Empty;

        public string ParameterValue { get; set; } = string.Empty;
    }
}
