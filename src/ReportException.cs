namespace Trysac.OpenxmlReportWriter;

/// <summary>Exception thrown when a report template, cell value, formula or reference violates SpreadsheetML rules.</summary>
public class ReportException : ArgumentException
{
    /// <summary>Gets the stable machine-readable error code.</summary>
    public string Code { get; }

    /// <summary>Initializes a new instance of the <see cref="ReportException"/> class.</summary>
    public ReportException() : base("Report validation failed.")
    {
        Code = "invalid_value";
    }

    /// <summary>Initializes a new instance of the <see cref="ReportException"/> class with a message.</summary>
    /// <param name="message">A short error description (at most 200 characters of input).</param>
    public ReportException(string message) : base(TruncateMessage(message))
    {
        Code = "invalid_value";
    }

    /// <summary>Initializes a new instance of the <see cref="ReportException"/> class with a message and inner exception.</summary>
    /// <param name="message">A short error description (at most 200 characters of input).</param>
    /// <param name="innerException">The inner cause of this error.</param>
    public ReportException(string message, Exception innerException) : base(TruncateMessage(message), innerException)
    {
        Code = "invalid_value";
    }

    /// <summary>Initializes a new instance of the <see cref="ReportException"/> class with an error code and message.</summary>
    /// <param name="code">The machine-readable error code.</param>
    /// <param name="message">A short error description (at most 200 characters of input).</param>
    public ReportException(string code, string message) : base(TruncateMessage(message))
    {
        ArgumentNullException.ThrowIfNull(code);
        Code = code;
    }

    /// <summary>Initializes a new instance of the <see cref="ReportException"/> class with an error code, message and inner exception.</summary>
    /// <param name="code">The machine-readable error code.</param>
    /// <param name="message">A short error description (at most 200 characters of input).</param>
    /// <param name="innerException">The inner cause of this error.</param>
    public ReportException(string code, string message, Exception innerException) : base(TruncateMessage(message), innerException)
    {
        ArgumentNullException.ThrowIfNull(code);
        Code = code;
    }

    private static string TruncateMessage(string? message)
    {
        if (string.IsNullOrEmpty(message))
        {
            return "Report validation failed.";
        }
        return message.Length <= 200 ? message : string.Concat(message.AsSpan(0, 197), "...");
    }
}
