namespace SprintReporting.Application.DTOs;

public class GeneratedReportResultDto
{
    public string FileName { get; set; } = "SprintReport.pptx";

    public string ContentType { get; set; }
        = "application/vnd.openxmlformats-officedocument.presentationml.presentation";

    public byte[] FileContent { get; set; } = Array.Empty<byte>();
}