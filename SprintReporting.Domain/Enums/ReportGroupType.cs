namespace SprintReporting.Domain.Enums;

public enum ReportGroupType
{
    /// <summary>
    /// Selecting this includes every report group.
    /// </summary>
    All,
    Delivery,
    PriorityRisk,
    TeamAnalysis,
    ComponentAnalysis,
    TechnicalDebt,
    AgingBacklog
}