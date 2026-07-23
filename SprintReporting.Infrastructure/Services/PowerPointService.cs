using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using SprintReporting.Application.Interfaces;
using SprintReporting.Domain.Models;

namespace SprintReporting.Infrastructure.Services;

public class PowerPointService : IPowerPointService
{
    private const long EmuPerInch = 914400;

    private const string Navy = "0F172A";
    private const string Blue = "2563EB";
    private const string LightBlue = "DBEAFE";
    private const string Green = "16A34A";
    private const string Orange = "F97316";
    private const string Red = "DC2626";
    private const string Purple = "7C3AED";
    private const string GrayText = "475569";
    private const string LightGray = "F8FAFC";
    private const string BorderGray = "CBD5E1";
    private const string White = "FFFFFF";

    public Task<byte[]> GeneratePresentationAsync(
        SprintMetrics metrics,
        AIInsightResult aiInsights,
        ReportConfiguration configuration,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();

        using var memoryStream = new MemoryStream();

        using (var presentationDocument = PresentationDocument.Create(
            memoryStream,
            PresentationDocumentType.Presentation))
        {
            var presentationPart = presentationDocument.AddPresentationPart();
            presentationPart.Presentation = new P.Presentation();

            var slideLayoutPart = CreatePresentationParts(
                presentationPart,
                out var slideIdList);

            uint slideId = 256;

            // The 14 required slides are always generated.
            AddCoverSlide(presentationPart, slideLayoutPart, slideIdList, slideId++, metrics, aiInsights);
            AddExecutiveSummarySlide(presentationPart, slideLayoutPart, slideIdList, slideId++, aiInsights, metrics);
            AddSprintOverviewSlide(presentationPart, slideLayoutPart, slideIdList, slideId++, metrics, aiInsights);
            AddStatusDistributionSlide(presentationPart, slideLayoutPart, slideIdList, slideId++, metrics, aiInsights);
            AddPriorityAnalysisSlide(presentationPart, slideLayoutPart, slideIdList, slideId++, metrics, aiInsights);
            AddIssueTypeSlide(presentationPart, slideLayoutPart, slideIdList, slideId++, metrics);
            AddTeamWorkloadSlide(presentationPart, slideLayoutPart, slideIdList, slideId++, metrics, aiInsights);
            AddAssigneeAnalysisSlide(presentationPart, slideLayoutPart, slideIdList, slideId++, metrics, aiInsights);
            AddComponentAnalysisSlide(presentationPart, slideLayoutPart, slideIdList, slideId++, metrics, aiInsights);
            AddResolutionSummarySlide(presentationPart, slideLayoutPart, slideIdList, slideId++, metrics, aiInsights);
            AddKeyInsightsSlide(presentationPart, slideLayoutPart, slideIdList, slideId++, metrics, aiInsights);
            AddRisksObservationsSlide(presentationPart, slideLayoutPart, slideIdList, slideId++, metrics, aiInsights);
            AddRecommendationsSlide(presentationPart, slideLayoutPart, slideIdList, slideId++, aiInsights);
            AddConclusionSlide(presentationPart, slideLayoutPart, slideIdList, slideId++, metrics, aiInsights);

            presentationPart.Presentation.Save();
        }

        return Task.FromResult(memoryStream.ToArray());
    }

    private static SlideLayoutPart CreatePresentationParts(
        PresentationPart presentationPart,
        out P.SlideIdList slideIdList)
    {
        var presentation = presentationPart.Presentation!;

        // 1. Slide master part.
        var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
        var slideMaster = CreateSlideMaster();
        slideMasterPart.SlideMaster = slideMaster;

        // 2. Theme part attached to the master (required for a valid presentation).
        var themePart = slideMasterPart.AddNewPart<ThemePart>();
        themePart.Theme = CreateTheme();

        // 3. Slide layout part, with a back-relationship to its master.
        var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
        slideLayoutPart.SlideLayout = CreateSlideLayout();
        slideLayoutPart.AddPart(slideMasterPart);

        // Master -> layout id list (layout id must be >= 2147483648).
        slideMaster.SlideLayoutIdList = new P.SlideLayoutIdList(
            new P.SlideLayoutId
            {
                Id = 2147483649U,
                RelationshipId = slideMasterPart.GetIdOfPart(slideLayoutPart)
            });

        slideMasterPart.SlideMaster.Save();
        slideLayoutPart.SlideLayout.Save();
        themePart.Theme.Save();

        // Assign presentation children through typed properties so the SDK
        // serializes them in the schema-required order:
        // sldMasterIdLst -> sldIdLst -> sldSz -> notesSz.
        presentation.SlideMasterIdList = new P.SlideMasterIdList(
            new P.SlideMasterId
            {
                Id = 2147483648U,
                RelationshipId = presentationPart.GetIdOfPart(slideMasterPart)
            });

        slideIdList = new P.SlideIdList();
        presentation.SlideIdList = slideIdList;

        presentation.SlideSize = new P.SlideSize
        {
            Cx = 12192000,
            Cy = 6858000
        };

        presentation.NotesSize = new P.NotesSize
        {
            Cx = 6858000,
            Cy = 9144000
        };

        return slideLayoutPart;
    }

    private static P.SlideMaster CreateSlideMaster()
    {
        return new P.SlideMaster(
            CreateCommonSlideData("Slide Master"),
            new P.ColorMap
            {
                Background1 = A.ColorSchemeIndexValues.Light1,
                Text1 = A.ColorSchemeIndexValues.Dark1,
                Background2 = A.ColorSchemeIndexValues.Light2,
                Text2 = A.ColorSchemeIndexValues.Dark2,
                Accent1 = A.ColorSchemeIndexValues.Accent1,
                Accent2 = A.ColorSchemeIndexValues.Accent2,
                Accent3 = A.ColorSchemeIndexValues.Accent3,
                Accent4 = A.ColorSchemeIndexValues.Accent4,
                Accent5 = A.ColorSchemeIndexValues.Accent5,
                Accent6 = A.ColorSchemeIndexValues.Accent6,
                Hyperlink = A.ColorSchemeIndexValues.Hyperlink,
                FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink
            });
    }

    private static P.SlideLayout CreateSlideLayout()
    {
        return new P.SlideLayout(
            CreateCommonSlideData("Blank Slide Layout"),
            new P.ColorMapOverride(new A.MasterColorMapping()))
        {
            Type = P.SlideLayoutValues.Blank
        };
    }

    private static A.Theme CreateTheme()
    {
        const string themeXml =
            "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Office Theme\">" +
              "<a:themeElements>" +
                "<a:clrScheme name=\"Office\">" +
                  "<a:dk1><a:sysClr val=\"windowText\" lastClr=\"000000\"/></a:dk1>" +
                  "<a:lt1><a:sysClr val=\"window\" lastClr=\"FFFFFF\"/></a:lt1>" +
                  "<a:dk2><a:srgbClr val=\"44546A\"/></a:dk2>" +
                  "<a:lt2><a:srgbClr val=\"E7E6E6\"/></a:lt2>" +
                  "<a:accent1><a:srgbClr val=\"2563EB\"/></a:accent1>" +
                  "<a:accent2><a:srgbClr val=\"16A34A\"/></a:accent2>" +
                  "<a:accent3><a:srgbClr val=\"F97316\"/></a:accent3>" +
                  "<a:accent4><a:srgbClr val=\"DC2626\"/></a:accent4>" +
                  "<a:accent5><a:srgbClr val=\"7C3AED\"/></a:accent5>" +
                  "<a:accent6><a:srgbClr val=\"0F172A\"/></a:accent6>" +
                  "<a:hlink><a:srgbClr val=\"0563C1\"/></a:hlink>" +
                  "<a:folHlink><a:srgbClr val=\"954F72\"/></a:folHlink>" +
                "</a:clrScheme>" +
                "<a:fontScheme name=\"Office\">" +
                  "<a:majorFont>" +
                    "<a:latin typeface=\"Calibri Light\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/>" +
                  "</a:majorFont>" +
                  "<a:minorFont>" +
                    "<a:latin typeface=\"Calibri\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/>" +
                  "</a:minorFont>" +
                "</a:fontScheme>" +
                "<a:fmtScheme name=\"Office\">" +
                  "<a:fillStyleLst>" +
                    "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>" +
                    "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>" +
                    "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>" +
                  "</a:fillStyleLst>" +
                  "<a:lnStyleLst>" +
                    "<a:ln w=\"6350\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/></a:ln>" +
                    "<a:ln w=\"12700\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/></a:ln>" +
                    "<a:ln w=\"19050\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\"><a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill><a:prstDash val=\"solid\"/></a:ln>" +
                  "</a:lnStyleLst>" +
                  "<a:effectStyleLst>" +
                    "<a:effectStyle><a:effectLst/></a:effectStyle>" +
                    "<a:effectStyle><a:effectLst/></a:effectStyle>" +
                    "<a:effectStyle><a:effectLst/></a:effectStyle>" +
                  "</a:effectStyleLst>" +
                  "<a:bgFillStyleLst>" +
                    "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>" +
                    "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>" +
                    "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>" +
                  "</a:bgFillStyleLst>" +
                "</a:fmtScheme>" +
              "</a:themeElements>" +
            "</a:theme>";

        return new A.Theme(themeXml);
    }

    private static P.CommonSlideData CreateCommonSlideData(string name)
    {
        return new P.CommonSlideData(
            new P.ShapeTree(
                new P.NonVisualGroupShapeProperties(
                    new P.NonVisualDrawingProperties
                    {
                        Id = 1U,
                        Name = name
                    },
                    new P.NonVisualGroupShapeDrawingProperties(),
                    new P.ApplicationNonVisualDrawingProperties()),
                new P.GroupShapeProperties(
                    new A.TransformGroup(
                        new A.Offset { X = 0, Y = 0 },
                        new A.Extents { Cx = 0, Cy = 0 },
                        new A.ChildOffset { X = 0, Y = 0 },
                        new A.ChildExtents { Cx = 0, Cy = 0 }))));
    }

    private static SlideContext CreateSlide(
        PresentationPart presentationPart,
        SlideLayoutPart slideLayoutPart,
        P.SlideIdList slideIdList,
        uint slideId,
        string slideName)
    {
        var slidePart = presentationPart.AddNewPart<SlidePart>();

        slidePart.AddPart(slideLayoutPart);

        var shapeTree = new P.ShapeTree(
            new P.NonVisualGroupShapeProperties(
                new P.NonVisualDrawingProperties
                {
                    Id = 1U,
                    Name = slideName
                },
                new P.NonVisualGroupShapeDrawingProperties(),
                new P.ApplicationNonVisualDrawingProperties()),
            new P.GroupShapeProperties(
                new A.TransformGroup(
                    new A.Offset { X = 0, Y = 0 },
                    new A.Extents { Cx = 0, Cy = 0 },
                    new A.ChildOffset { X = 0, Y = 0 },
                    new A.ChildExtents { Cx = 0, Cy = 0 })));

        slidePart.Slide = new P.Slide(
            new P.CommonSlideData(shapeTree),
            new P.ColorMapOverride(new A.MasterColorMapping()));

        slidePart.Slide.Save();

        var relationshipId = presentationPart.GetIdOfPart(slidePart);

        slideIdList.Append(new P.SlideId
        {
            Id = slideId,
            RelationshipId = relationshipId
        });

        return new SlideContext(slidePart, shapeTree);
    }

    // ---------------------------------------------------------------------
    // Slide 1: Cover
    // ---------------------------------------------------------------------
    private static void AddCoverSlide(
        PresentationPart presentationPart,
        SlideLayoutPart slideLayoutPart,
        P.SlideIdList slideIdList,
        uint slideId,
        SprintMetrics metrics,
        AIInsightResult aiInsights)
    {
        var context = CreateSlide(
            presentationPart,
            slideLayoutPart,
            slideIdList,
            slideId,
            "Cover Slide");

        AddBackground(context.ShapeTree, Navy);

        AddRectangle(context.ShapeTree, 2U, "Accent Block", 0.0, 0.0, 13.333, 0.18, Blue, Blue);

        AddText(context.ShapeTree, 3U, "Title", "Sprint Report",
            0.75, 1.0, 11.8, 0.75, 4000, White, true);

        AddText(context.ShapeTree, 4U, "Subtitle", "AI-generated sprint analysis from Jira issue data",
            0.78, 1.9, 11.0, 0.4, 1800, "BFDBFE", false);

        AddText(context.ShapeTree, 5U, "Generated Date", $"Generated on {DateTime.Now:dd MMM yyyy}",
            0.78, 2.4, 6.0, 0.3, 1300, "CBD5E1", false);

        AddMetricCard(context.ShapeTree, 6U, "Total Issues", metrics.TotalIssues.ToString(),
            "Parsed from Excel", 0.78, 3.35, 2.75, 1.25, Blue);

        AddMetricCard(context.ShapeTree, 10U, "Sprints", metrics.TotalSprints.ToString(),
            "Distinct sprints", 3.82, 3.35, 2.75, 1.25, Purple);

        AddMetricCard(context.ShapeTree, 14U, "Completion", $"{metrics.CompletionPercentage}%",
            "Sprint progress", 6.86, 3.35, 2.75, 1.25, Green);

        AddMetricCard(context.ShapeTree, 18U, "Open Issues", metrics.OpenIssues.ToString(),
            "Pending work", 9.9, 3.35, 2.75, 1.25, Orange);

        AddText(context.ShapeTree, 22U, "Architecture Note",
            "Metrics are calculated deterministically in the backend. Only aggregated KPI data is sent to AI.",
            0.8, 5.3, 11.0, 0.45, 1500, "E2E8F0", false);

        AddFooter(context.ShapeTree, 23U);

        context.SlidePart.Slide.Save();
    }

    // ---------------------------------------------------------------------
    // Slide 2: Executive Summary
    // ---------------------------------------------------------------------
    private static void AddExecutiveSummarySlide(
        PresentationPart presentationPart,
        SlideLayoutPart slideLayoutPart,
        P.SlideIdList slideIdList,
        uint slideId,
        AIInsightResult aiInsights,
        SprintMetrics metrics)
    {
        var context = CreateReportSlide(
            presentationPart, slideLayoutPart, slideIdList, slideId,
            "Executive Summary",
            "AI-generated overview based on aggregated sprint KPIs");

        AddText(context.ShapeTree, 20U, "Summary Text", aiInsights.ExecutiveSummary,
            0.75, 1.55, 11.8, 2.3, 1800, Navy, false);

        AddKpiStrip(context.ShapeTree, 40U, metrics, 0.75, 4.25);

        AddFooter(context.ShapeTree, 80U);

        context.SlidePart.Slide.Save();
    }

    // ---------------------------------------------------------------------
    // Slide 3: Sprint Overview
    // ---------------------------------------------------------------------
    private static void AddSprintOverviewSlide(
        PresentationPart presentationPart,
        SlideLayoutPart slideLayoutPart,
        P.SlideIdList slideIdList,
        uint slideId,
        SprintMetrics metrics,
        AIInsightResult aiInsights)
    {
        var context = CreateReportSlide(
            presentationPart, slideLayoutPart, slideIdList, slideId,
            "Sprint Overview",
            "Issue volume and sprint-wise distribution");

        AddMetricCard(context.ShapeTree, 20U, "Total Issues", metrics.TotalIssues.ToString(),
            "All issues", 0.75, 1.35, 2.85, 1.15, Blue);

        AddMetricCard(context.ShapeTree, 24U, "Total Sprints", metrics.TotalSprints.ToString(),
            "Distinct sprints", 3.72, 1.35, 2.85, 1.15, Purple);

        AddMetricCard(context.ShapeTree, 28U, "Completed", metrics.CompletedIssues.ToString(),
            "Done issues", 6.69, 1.35, 2.85, 1.15, Green);

        AddMetricCard(context.ShapeTree, 32U, "Open", metrics.OpenIssues.ToString(),
            "Not done", 9.66, 1.35, 2.85, 1.15, Orange);

        AddDistributionPanel(context.ShapeTree, 50U, "Issues by Sprint", metrics.SprintIssueCount,
            0.75, 3.0, 6.05, 3.1, Blue, 8);

        AddInsightBox(context.ShapeTree, 80U, "AI Sprint Overview", aiInsights.SprintOverview,
            6.95, 3.0, 5.55, 1.5, LightBlue, Blue);

        AddIssueListPanel(context.ShapeTree, 90U, "Recently Created", metrics.RecentlyCreatedIssues,
            6.95, 4.65, 5.55, 1.45, Green, false);

        AddFooter(context.ShapeTree, 120U);

        context.SlidePart.Slide.Save();
    }

    // ---------------------------------------------------------------------
    // Slide 4: Issue Status Distribution
    // ---------------------------------------------------------------------
    private static void AddStatusDistributionSlide(
        PresentationPart presentationPart,
        SlideLayoutPart slideLayoutPart,
        P.SlideIdList slideIdList,
        uint slideId,
        SprintMetrics metrics,
        AIInsightResult aiInsights)
    {
        var context = CreateReportSlide(
            presentationPart, slideLayoutPart, slideIdList, slideId,
            "Issue Status Distribution",
            "Distribution of issues across workflow statuses");

        AddDistributionPanel(context.ShapeTree, 20U, "Issues by Status", metrics.StatusDistribution,
            0.75, 1.45, 6.05, 4.6, Blue, 10);

        AddInsightBox(context.ShapeTree, 60U, "Status Analysis", aiInsights.StatusAnalysis,
            6.95, 1.45, 5.55, 2.15, LightBlue, Blue);

        AddMetricCard(context.ShapeTree, 70U, "Completion", $"{metrics.CompletionPercentage}%",
            "Sprint progress", 6.95, 3.85, 2.6, 1.1, Green);

        AddMetricCard(context.ShapeTree, 74U, "Open", metrics.OpenIssues.ToString(),
            "Pending", 9.9, 3.85, 2.6, 1.1, Orange);

        AddText(context.ShapeTree, 80U, "Note",
            $"{metrics.CompletedIssues} completed of {metrics.TotalIssues} total issues.",
            6.95, 5.1, 5.55, 0.9, 1400, GrayText, false);

        AddFooter(context.ShapeTree, 90U);

        context.SlidePart.Slide.Save();
    }

    // ---------------------------------------------------------------------
    // Slide 5: Priority Analysis
    // ---------------------------------------------------------------------
    private static void AddPriorityAnalysisSlide(
        PresentationPart presentationPart,
        SlideLayoutPart slideLayoutPart,
        P.SlideIdList slideIdList,
        uint slideId,
        SprintMetrics metrics,
        AIInsightResult aiInsights)
    {
        var context = CreateReportSlide(
            presentationPart, slideLayoutPart, slideIdList, slideId,
            "Priority Analysis",
            "Issue distribution and risk by priority");

        AddDistributionPanel(context.ShapeTree, 20U, "Issues by Priority", metrics.PriorityDistribution,
            0.75, 1.45, 6.05, 4.6, Red, 10);

        AddInsightBox(context.ShapeTree, 60U, "Priority Analysis", aiInsights.PriorityAnalysis,
            6.95, 1.45, 5.55, 2.15, LightBlue, Blue);

        AddMetricCard(context.ShapeTree, 70U, "High Priority Open", metrics.HighPriorityOpenIssues.ToString(),
            "High/Highest and open", 6.95, 3.85, 5.55, 1.1, Red);

        AddText(context.ShapeTree, 80U, "Note",
            "High-priority open items should be resolved first to protect delivery confidence.",
            6.95, 5.15, 5.55, 0.9, 1400, GrayText, false);

        AddFooter(context.ShapeTree, 90U);

        context.SlidePart.Slide.Save();
    }

    // ---------------------------------------------------------------------
    // Slide 6: Issue Type Distribution
    // ---------------------------------------------------------------------
    private static void AddIssueTypeSlide(
        PresentationPart presentationPart,
        SlideLayoutPart slideLayoutPart,
        P.SlideIdList slideIdList,
        uint slideId,
        SprintMetrics metrics)
    {
        var context = CreateReportSlide(
            presentationPart, slideLayoutPart, slideIdList, slideId,
            "Issue Type Distribution",
            "Breakdown of issues by type and due-date health");

        AddDistributionPanel(context.ShapeTree, 20U, "Issues by Type", metrics.IssueTypeDistribution,
            0.75, 1.45, 6.05, 4.6, Green, 10);

        AddInsightBox(context.ShapeTree, 60U, "How to Read This",
            "Issue type distribution highlights the balance between features, bugs, and other work categories this sprint.",
            6.95, 1.45, 5.55, 2.15, LightBlue, Blue);

        AddDistributionPanel(context.ShapeTree, 70U, "Due Date Buckets", metrics.DueDateDistribution,
            6.95, 3.85, 5.55, 2.25, Orange, 6);

        AddFooter(context.ShapeTree, 100U);

        context.SlidePart.Slide.Save();
    }

    // ---------------------------------------------------------------------
    // Slide 7: Team Workload
    // ---------------------------------------------------------------------
    private static void AddTeamWorkloadSlide(
        PresentationPart presentationPart,
        SlideLayoutPart slideLayoutPart,
        P.SlideIdList slideIdList,
        uint slideId,
        SprintMetrics metrics,
        AIInsightResult aiInsights)
    {
        var context = CreateReportSlide(
            presentationPart, slideLayoutPart, slideIdList, slideId,
            "Team Workload",
            "Workload distribution across teams and reporters");

        AddDistributionPanel(context.ShapeTree, 20U, "Issues by Team", metrics.TeamDistribution,
            0.75, 1.45, 6.05, 4.6, Blue, 10);

        AddInsightBox(context.ShapeTree, 60U, "Team Workload Analysis", aiInsights.TeamWorkloadAnalysis,
            6.95, 1.45, 5.55, 2.15, LightBlue, Blue);

        AddDistributionPanel(context.ShapeTree, 70U, "Issues by Reporter", metrics.ReporterDistribution,
            6.95, 3.85, 5.55, 2.25, Purple, 6);

        AddFooter(context.ShapeTree, 100U);

        context.SlidePart.Slide.Save();
    }

    // ---------------------------------------------------------------------
    // Slide 8: Assignee Analysis
    // ---------------------------------------------------------------------
    private static void AddAssigneeAnalysisSlide(
        PresentationPart presentationPart,
        SlideLayoutPart slideLayoutPart,
        P.SlideIdList slideIdList,
        uint slideId,
        SprintMetrics metrics,
        AIInsightResult aiInsights)
    {
        var context = CreateReportSlide(
            presentationPart, slideLayoutPart, slideIdList, slideId,
            "Assignee Analysis",
            "Assigned vs completed work by assignee");

        AddDistributionPanel(context.ShapeTree, 20U, "Issues per Assignee", metrics.AssigneeDistribution,
            0.75, 1.45, 6.05, 4.6, Blue, 10);

        AddDistributionPanel(context.ShapeTree, 60U, "Completed per Assignee", metrics.CompletedWorkPerAssignee,
            6.95, 1.45, 5.55, 2.15, Green, 6);

        AddInsightBox(context.ShapeTree, 90U, "Assignee Productivity", aiInsights.AssigneeProductivitySummary,
            6.95, 3.85, 5.55, 2.25, LightBlue, Blue);

        AddFooter(context.ShapeTree, 120U);

        context.SlidePart.Slide.Save();
    }

    // ---------------------------------------------------------------------
    // Slide 9: Component Analysis
    // ---------------------------------------------------------------------
    private static void AddComponentAnalysisSlide(
        PresentationPart presentationPart,
        SlideLayoutPart slideLayoutPart,
        P.SlideIdList slideIdList,
        uint slideId,
        SprintMetrics metrics,
        AIInsightResult aiInsights)
    {
        var context = CreateReportSlide(
            presentationPart, slideLayoutPart, slideIdList, slideId,
            "Component Analysis",
            "Issue concentration by system component");

        AddDistributionPanel(context.ShapeTree, 20U, "Issues by Component", metrics.ComponentDistribution,
            0.75, 1.45, 6.05, 4.6, Purple, 10);

        AddInsightBox(context.ShapeTree, 60U, "Component Analysis", aiInsights.ComponentAnalysis,
            6.95, 1.45, 5.55, 2.15, LightBlue, Blue);

        AddText(context.ShapeTree, 70U, "Note",
            "Components with higher issue concentration may need ownership review or dependency management.",
            6.95, 3.85, 5.55, 1.2, 1400, GrayText, false);

        AddFooter(context.ShapeTree, 90U);

        context.SlidePart.Slide.Save();
    }

    // ---------------------------------------------------------------------
    // Slide 10: Resolution Summary
    // ---------------------------------------------------------------------
    private static void AddResolutionSummarySlide(
        PresentationPart presentationPart,
        SlideLayoutPart slideLayoutPart,
        P.SlideIdList slideIdList,
        uint slideId,
        SprintMetrics metrics,
        AIInsightResult aiInsights)
    {
        var context = CreateReportSlide(
            presentationPart, slideLayoutPart, slideIdList, slideId,
            "Resolution Summary",
            "Resolution outcomes across issues");

        AddDistributionPanel(context.ShapeTree, 20U, "Resolution Distribution", metrics.ResolutionDistribution,
            0.75, 1.45, 6.05, 4.6, Green, 10);

        AddInsightBox(context.ShapeTree, 60U, "Resolution Summary", aiInsights.ResolutionSummary,
            6.95, 1.45, 5.55, 2.15, LightBlue, Blue);

        AddMetricCard(context.ShapeTree, 70U, "Completed", metrics.CompletedIssues.ToString(),
            "Resolved / done", 6.95, 3.85, 2.6, 1.1, Green);

        AddMetricCard(context.ShapeTree, 74U, "Open", metrics.OpenIssues.ToString(),
            "Unresolved", 9.9, 3.85, 2.6, 1.1, Orange);

        AddFooter(context.ShapeTree, 90U);

        context.SlidePart.Slide.Save();
    }

    // ---------------------------------------------------------------------
    // Slide 11: Key Insights
    // ---------------------------------------------------------------------
    private static void AddKeyInsightsSlide(
        PresentationPart presentationPart,
        SlideLayoutPart slideLayoutPart,
        P.SlideIdList slideIdList,
        uint slideId,
        SprintMetrics metrics,
        AIInsightResult aiInsights)
    {
        var context = CreateReportSlide(
            presentationPart, slideLayoutPart, slideIdList, slideId,
            "Key Insights",
            "Notable observations and label themes");

        AddBulletedPanel(context.ShapeTree, 20U, "Key Observations", aiInsights.Observations,
            0.75, 1.45, 6.05, 4.6, Blue);

        AddDistributionPanel(context.ShapeTree, 60U, "Issues by Label", metrics.LabelDistribution,
            6.95, 1.45, 5.55, 2.15, Purple, 6);

        AddInsightBox(context.ShapeTree, 90U, "Label Analysis", aiInsights.LabelAnalysis,
            6.95, 3.85, 5.55, 2.25, LightBlue, Blue);

        AddFooter(context.ShapeTree, 120U);

        context.SlidePart.Slide.Save();
    }

    // ---------------------------------------------------------------------
    // Slide 12: Risks & Observations
    // ---------------------------------------------------------------------
    private static void AddRisksObservationsSlide(
        PresentationPart presentationPart,
        SlideLayoutPart slideLayoutPart,
        P.SlideIdList slideIdList,
        uint slideId,
        SprintMetrics metrics,
        AIInsightResult aiInsights)
    {
        var context = CreateReportSlide(
            presentationPart, slideLayoutPart, slideIdList, slideId,
            "Risks & Observations",
            "Delivery risks, overdue work, and due-date health");

        AddBulletedPanel(context.ShapeTree, 20U, "Potential Risks", aiInsights.Risks,
            0.75, 1.45, 6.05, 2.35, Red);

        AddBulletedPanel(context.ShapeTree, 60U, "Observations", aiInsights.Observations,
            0.75, 4.0, 6.05, 2.35, Blue);

        AddDistributionPanel(context.ShapeTree, 100U, "Due Date Distribution", metrics.DueDateDistribution,
            6.95, 1.45, 5.55, 2.35, Orange, 6);

        AddIssueListPanel(context.ShapeTree, 130U, "Overdue Issues", metrics.OverdueIssues,
            6.95, 4.0, 5.55, 2.35, Red, true);

        AddFooter(context.ShapeTree, 160U);

        context.SlidePart.Slide.Save();
    }

    // ---------------------------------------------------------------------
    // Slide 13: AI Recommendations
    // ---------------------------------------------------------------------
    private static void AddRecommendationsSlide(
        PresentationPart presentationPart,
        SlideLayoutPart slideLayoutPart,
        P.SlideIdList slideIdList,
        uint slideId,
        AIInsightResult aiInsights)
    {
        var context = CreateReportSlide(
            presentationPart, slideLayoutPart, slideIdList, slideId,
            "AI Recommendations",
            "Recommended actions for sprint stakeholders");

        AddBulletedPanel(context.ShapeTree, 20U, "Recommended Actions", aiInsights.Recommendations,
            0.75, 1.45, 7.0, 4.6, Green);

        AddInsightBox(context.ShapeTree, 60U, "AI Usage Note",
            $"Provider used: {aiInsights.ProviderUsed}. The AI receives aggregated metrics only, not raw Excel rows.",
            8.1, 1.45, 4.4, 2.15, LightBlue, Blue);

        AddText(context.ShapeTree, 70U, "Final Note",
            "Use these recommendations as a starting point for sprint planning discussions.",
            8.1, 3.85, 4.4, 1.2, 1400, GrayText, false);

        AddFooter(context.ShapeTree, 90U);

        context.SlidePart.Slide.Save();
    }

    // ---------------------------------------------------------------------
    // Slide 14: Conclusion
    // ---------------------------------------------------------------------
    private static void AddConclusionSlide(
        PresentationPart presentationPart,
        SlideLayoutPart slideLayoutPart,
        P.SlideIdList slideIdList,
        uint slideId,
        SprintMetrics metrics,
        AIInsightResult aiInsights)
    {
        var context = CreateReportSlide(
            presentationPart, slideLayoutPart, slideIdList, slideId,
            "Conclusion",
            "Next steps and closing summary");

        AddBulletedPanel(context.ShapeTree, 20U, "Next Sprint Suggestions", aiInsights.NextSprintSuggestions,
            0.75, 1.45, 6.05, 4.6, Purple);

        AddIssueListPanel(context.ShapeTree, 60U, "Recently Updated", metrics.RecentlyUpdatedIssues,
            6.95, 1.45, 5.55, 2.35, Blue, false);

        AddInsightBox(context.ShapeTree, 90U, "Summary",
            $"{metrics.CompletedIssues} of {metrics.TotalIssues} issues completed ({metrics.CompletionPercentage}%). {metrics.OverdueIssueCount} issue(s) overdue.",
            6.95, 4.0, 5.55, 2.35, LightBlue, Green);

        AddFooter(context.ShapeTree, 120U);

        context.SlidePart.Slide.Save();
    }

    private static SlideContext CreateReportSlide(
        PresentationPart presentationPart,
        SlideLayoutPart slideLayoutPart,
        P.SlideIdList slideIdList,
        uint slideId,
        string title,
        string subtitle)
    {
        var context = CreateSlide(
            presentationPart,
            slideLayoutPart,
            slideIdList,
            slideId,
            title);

        AddBackground(context.ShapeTree, White);

        AddRectangle(context.ShapeTree, 2U, "Top Accent", 0.0, 0.0, 13.333, 0.14, Blue, Blue);

        AddText(context.ShapeTree, 3U, "Slide Title", title,
            0.72, 0.35, 11.8, 0.45, 2800, Navy, true);

        AddText(context.ShapeTree, 4U, "Slide Subtitle", subtitle,
            0.74, 0.86, 11.8, 0.3, 1200, GrayText, false);

        return context;
    }

    private static void AddKpiStrip(
        P.ShapeTree shapeTree,
        uint startId,
        SprintMetrics metrics,
        double x,
        double y)
    {
        AddMetricCard(shapeTree, startId, "Total", metrics.TotalIssues.ToString(),
            "Total issues", x, y, 2.75, 1.15, Blue);

        AddMetricCard(shapeTree, startId + 4, "Completed", metrics.CompletedIssues.ToString(),
            "Done issues", x + 3.0, y, 2.75, 1.15, Green);

        AddMetricCard(shapeTree, startId + 8, "Open", metrics.OpenIssues.ToString(),
            "Not Done", x + 6.0, y, 2.75, 1.15, Orange);

        AddMetricCard(shapeTree, startId + 12, "Completion", $"{metrics.CompletionPercentage}%",
            "Sprint progress", x + 9.0, y, 2.75, 1.15, Purple);
    }

    private static void AddMetricCard(
        P.ShapeTree shapeTree,
        uint startId,
        string label,
        string value,
        string caption,
        double x,
        double y,
        double width,
        double height,
        string accentColor)
    {
        AddRectangle(shapeTree, startId, $"{label} Card", x, y, width, height, LightGray, BorderGray);

        AddRectangle(shapeTree, startId + 1, $"{label} Accent", x, y, 0.08, height, accentColor, accentColor);

        AddText(shapeTree, startId + 2, $"{label} Value", value,
            x + 0.22, y + 0.18, width - 0.35, 0.36, 2500, accentColor, true);

        AddText(shapeTree, startId + 3, $"{label} Label", $"{label}\n{caption}",
            x + 0.22, y + 0.58, width - 0.35, 0.42, 1000, GrayText, false);
    }

    private static void AddDistributionPanel(
        P.ShapeTree shapeTree,
        uint startId,
        string title,
        Dictionary<string, int> values,
        double x,
        double y,
        double width,
        double height,
        string accentColor,
        int take = 8)
    {
        AddRectangle(shapeTree, startId, $"{title} Panel", x, y, width, height, LightGray, BorderGray);

        AddRectangle(shapeTree, startId + 1, $"{title} Header", x, y, width, 0.42, accentColor, accentColor);

        AddText(shapeTree, startId + 2, $"{title} Header Text", title,
            x + 0.22, y + 0.1, width - 0.45, 0.25, 1300, White, true);

        var lines = BuildDistributionLines(values, take);

        AddText(shapeTree, startId + 3, $"{title} Body", string.Join("\n", lines),
            x + 0.28, y + 0.65, width - 0.55, height - 0.85, 1300, Navy, false);
    }

    private static void AddBulletedPanel(
        P.ShapeTree shapeTree,
        uint startId,
        string title,
        List<string> values,
        double x,
        double y,
        double width,
        double height,
        string accentColor)
    {
        AddRectangle(shapeTree, startId, $"{title} Panel", x, y, width, height, LightGray, BorderGray);

        AddRectangle(shapeTree, startId + 1, $"{title} Header", x, y, width, 0.42, accentColor, accentColor);

        AddText(shapeTree, startId + 2, $"{title} Header Text", title,
            x + 0.22, y + 0.1, width - 0.45, 0.25, 1300, White, true);

        var text = values.Count == 0
            ? "No insights available."
            : string.Join("\n", values.Select(value => $"• {value}"));

        AddText(shapeTree, startId + 3, $"{title} Body", text,
            x + 0.28, y + 0.65, width - 0.55, height - 0.85, 1250, Navy, false);
    }

    private static void AddInsightBox(
        P.ShapeTree shapeTree,
        uint startId,
        string title,
        string body,
        double x,
        double y,
        double width,
        double height,
        string fillColor,
        string accentColor)
    {
        AddRectangle(shapeTree, startId, $"{title} Box", x, y, width, height, fillColor, accentColor);

        AddText(shapeTree, startId + 1, $"{title} Heading", title,
            x + 0.25, y + 0.18, width - 0.5, 0.28, 1400, accentColor, true);

        AddText(shapeTree, startId + 2, $"{title} Body", body,
            x + 0.25, y + 0.58, width - 0.5, height - 0.75, 1200, Navy, false);
    }

    private static void AddIssueListPanel(
        P.ShapeTree shapeTree,
        uint startId,
        string title,
        List<IssueSummaryMetric> issues,
        double x,
        double y,
        double width,
        double height,
        string accentColor,
        bool showOverdue)
    {
        AddRectangle(shapeTree, startId, $"{title} Panel", x, y, width, height, LightGray, BorderGray);

        AddRectangle(shapeTree, startId + 1, $"{title} Header", x, y, width, 0.42, accentColor, accentColor);

        AddText(shapeTree, startId + 2, $"{title} Header Text", title,
            x + 0.22, y + 0.1, width - 0.45, 0.25, 1300, White, true);

        string text;

        if (issues is null || issues.Count == 0)
        {
            text = "No issues to display.";
        }
        else
        {
            text = string.Join(
                "\n",
                issues.Take(5).Select(issue =>
                {
                    var suffix = showOverdue
                        ? $"{issue.DaysOverdue ?? 0}d overdue"
                        : (issue.Date.HasValue ? issue.Date.Value.ToString("dd MMM") : string.Empty);

                    return $"{issue.IssueKey}  {issue.Status}  {suffix}".TrimEnd();
                }));
        }

        AddText(shapeTree, startId + 3, $"{title} Body", text,
            x + 0.28, y + 0.7, width - 0.55, height - 0.9, 1200, Navy, false);
    }

    private static void AddFooter(
        P.ShapeTree shapeTree,
        uint id)
    {
        AddRectangle(shapeTree, id, "Footer Line", 0.72, 6.95, 11.9, 0.01, BorderGray, BorderGray);

        AddText(shapeTree, id + 1, "Footer Text",
            $"AI Sprint Reporting Agent  |  Generated {DateTime.Now:dd MMM yyyy}",
            0.75, 7.05, 7.5, 0.22, 850, GrayText, false);

        AddText(shapeTree, id + 2, "Footer Note",
            "Metrics calculated in C#; AI used only for narrative insights.",
            8.2, 7.05, 4.25, 0.22, 850, GrayText, false);
    }

    private static void AddBackground(
        P.ShapeTree shapeTree,
        string color)
    {
        AddRectangle(shapeTree, 1000U, "Background", 0, 0, 13.333, 7.5, color, color);
    }

    private static void AddRectangle(
        P.ShapeTree shapeTree,
        uint id,
        string name,
        double x,
        double y,
        double width,
        double height,
        string fillColor,
        string outlineColor)
    {
        shapeTree.Append(
            new P.Shape(
                new P.NonVisualShapeProperties(
                    new P.NonVisualDrawingProperties
                    {
                        Id = id,
                        Name = name
                    },
                    new P.NonVisualShapeDrawingProperties(
                        new A.ShapeLocks
                        {
                            NoGrouping = true
                        }),
                    new P.ApplicationNonVisualDrawingProperties()),
                new P.ShapeProperties(
                    new A.Transform2D(
                        new A.Offset
                        {
                            X = InchesToEmu(x),
                            Y = InchesToEmu(y)
                        },
                        new A.Extents
                        {
                            Cx = InchesToEmu(width),
                            Cy = InchesToEmu(height)
                        }),
                    new A.PresetGeometry(new A.AdjustValueList())
                    {
                        Preset = A.ShapeTypeValues.Rectangle
                    },
                    new A.SolidFill(
                        new A.RgbColorModelHex
                        {
                            Val = fillColor
                        }),
                    new A.Outline(
                        new A.SolidFill(
                            new A.RgbColorModelHex
                            {
                                Val = outlineColor
                            })))));
    }

    private static void AddText(
        P.ShapeTree shapeTree,
        uint id,
        string name,
        string text,
        double x,
        double y,
        double width,
        double height,
        int fontSize,
        string color,
        bool bold)
    {
        var textBody = new P.TextBody();

        textBody.Append(
            new A.BodyProperties
            {
                Wrap = A.TextWrappingValues.Square
            });

        textBody.Append(new A.ListStyle());

        foreach (var paragraph in CreateParagraphs(text, fontSize, color, bold))
        {
            textBody.Append(paragraph);
        }

        shapeTree.Append(
            new P.Shape(
                new P.NonVisualShapeProperties(
                    new P.NonVisualDrawingProperties
                    {
                        Id = id,
                        Name = name
                    },
                    new P.NonVisualShapeDrawingProperties(
                        new A.ShapeLocks
                        {
                            NoGrouping = true
                        }),
                    new P.ApplicationNonVisualDrawingProperties()),
                new P.ShapeProperties(
                    new A.Transform2D(
                        new A.Offset
                        {
                            X = InchesToEmu(x),
                            Y = InchesToEmu(y)
                        },
                        new A.Extents
                        {
                            Cx = InchesToEmu(width),
                            Cy = InchesToEmu(height)
                        }),
                    new A.PresetGeometry(new A.AdjustValueList())
                    {
                        Preset = A.ShapeTypeValues.Rectangle
                    },
                    new A.NoFill(),
                    new A.Outline(new A.NoFill())),
                textBody));
    }

    private static IEnumerable<A.Paragraph> CreateParagraphs(
        string text,
        int fontSize,
        string color,
        bool bold)
    {
        var lines = (text ?? string.Empty)
            .Replace("\r\n", "\n")
            .Split('\n');

        foreach (var line in lines)
        {
            var runProperties = new A.RunProperties
            {
                FontSize = fontSize,
                Bold = bold
            };

            // Font color must live INSIDE the run properties (a:rPr),
            // not as a direct child of the run (a:r).
            runProperties.Append(
                new A.SolidFill(
                    new A.RgbColorModelHex
                    {
                        Val = color
                    }));

            yield return new A.Paragraph(
                new A.Run(
                    runProperties,
                    new A.Text(line)));
        }
    }

    private static List<string> BuildDistributionLines(
        Dictionary<string, int> values,
        int take)
    {
        if (values.Count == 0)
        {
            return new List<string>
            {
                "No data available."
            };
        }

        return values
            .Take(take)
            .Select(item => $"{item.Key}:  {item.Value}")
            .ToList();
    }

    private static long InchesToEmu(double inches)
    {
        return (long)(inches * EmuPerInch);
    }

    private sealed record SlideContext(
        SlidePart SlidePart,
        P.ShapeTree ShapeTree);
}
