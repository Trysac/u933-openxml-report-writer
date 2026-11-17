namespace Trysac.OpenxmlReportWriter;

/// <summary>Thrown when a text input is empty or whitespace-only.</summary>
public sealed class BlankInputException : ReportException
{
    /// <summary>Initializes a new instance with the default message.</summary>
    public BlankInputException() : base("invalid_value", "input is blank")
    {
    }

    /// <summary>Initializes a new instance with a custom message.</summary>
    /// <param name="message">The exception message.</param>
    public BlankInputException(string message) : base("invalid_value", message)
    {
    }

    /// <summary>Initializes a new instance with a custom message and inner exception.</summary>
    /// <param name="message">The exception message.</param>
    /// <param name="innerException">The cause of this exception.</param>
    public BlankInputException(string message, Exception innerException) : base("invalid_value", message, innerException)
    {
    }
}
