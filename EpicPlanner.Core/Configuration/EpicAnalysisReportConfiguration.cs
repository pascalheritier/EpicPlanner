namespace EpicPlanner.Core.Configuration;

public class EpicAnalysisReportConfiguration
{
    /// <summary>Root folder containing the Sprint XX subdirectories (e.g. the "Sprint planning" folder).</summary>
    public string InputFolderPath { get; set; } = string.Empty;

    /// <summary>Full path of the generated HTML report.</summary>
    public string OutputHtmlPath { get; set; } = string.Empty;
}
