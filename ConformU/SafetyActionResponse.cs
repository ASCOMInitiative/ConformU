namespace ConformU
{
    /// <summary>
    /// Represents the result of a safety action operation
    /// </summary>
    public class SafetyActionResponse
    {
        /// <summary>
        /// Creates a response with the given outcome and message
        /// </summary>
        /// <param name="success">True for a successful outcome, otherwise false.</param>
        /// <param name="message">A human readable message describing the outcome.</param>
        public SafetyActionResponse(bool success, string message)
        {
            Success = success;
            Message = message;
        }

        /// <summary>
        /// Success or failure of the request
        /// </summary>
        public bool Success { get; set; } = false;

        /// <summary>
        /// Description of the outcome
        /// </summary>
        public string Message { get; set; } = "Message not set";
    }
}
