using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using SprintReporting.Application.Interfaces;
using SprintReporting.Domain.Enums;
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

            presentationPart.Presentation = new P.Presentation
            {
                SlideSize = new P.SlideSize
                {
                    Cx = 12192000,
                    Cy = 6858000
                },
                NotesSize = new P.NotesSize
                {
                    Cx = 6858000,
                    Cy = 9144000
                }
            };

            var slideLayoutPart = CreatePresentationParts(presentationPart);

            var slideIdList = new P.SlideIdList();
            presentationPart.Presentation.Append(slideIdList);

            uint slideId = 256;

            AddTitleSlide(
                presentationPart,
                slideLayoutPart,
                slideIdList,
                slideId++,
                metrics,
                aiInsights);

            AddExecutiveSummarySlide(
                presentationPart,
                slideLayoutPart,
                slideIdList,
                slideId++,
                aiInsights,
                metrics);

            if (configuration.SelectedGroups.Contains(ReportGroupType.Delivery))
            {
                AddDeliverySlide(
                    presentationPart,
                    slideLayoutPart,
                    slideIdList,
                    slideId++,
                    metrics);
            }

            if (configuration.SelectedGroups.Contains(ReportGroupType.PriorityRisk))
            {
                AddPriorityRiskSlide(
                    presentationPart,
                    slideLayoutPart,
                    slideIdList,
                    slideId++,
                    metrics,
                    aiInsights);
            }

            if (configuration.SelectedGroups.Contains(ReportGroupType.TeamAnalysis))
            {
                AddTeamAnalysisSlide(
                    presentationPart,
                    slideLayoutPart,
                    slideIdList,
                    slideId++,
                    metrics);
            }

            if (configuration.SelectedGroups.Contains(ReportGroupType.ComponentAnalysis))
            {
                AddComponentAnalysisSlide(
                    presentationPart,
                    slideLayoutPart,
                    slideIdList,
                    slideId++,
                    metrics);
            }

            if (configuration.SelectedGroups.Contains(ReportGroupType.TechnicalDebt))
            {
                AddTechnicalDebtSlide(
                    presentationPart,
                    slideLayoutPart,
                    slideIdList,
                    slideId++,
                    metrics);
            }

            if (configuration.SelectedGroups.Contains(ReportGroupType.AgingBacklog))
            {
                AddAgingBacklogSlide(
                    presentationPart,
                    slideLayoutPart,
                    slideIdList,
                    slideId++,
                    metrics);
            }

            AddRecommendationsSlide(
                presentationPart,
                slideLayoutPart,
                slideIdList,
                slideId++,
                aiInsights);

            presentationPart.Presentation.Save();
        }

        return Task.FromResult(memoryStream.ToArray());
    }

    private static SlideLayoutPart CreatePresentationParts(PresentationPart presentationPart)
    {
        var slideMasterPart = presentationPart.AddNewPart<SlideMasterPart>();
        var slideMaster = CreateSlideMaster();
        slideMasterPart.SlideMaster = slideMaster;

        var slideLayoutPart = slideMasterPart.AddNewPart<SlideLayoutPart>();
        var slideLayout = CreateSlideLayout();
        slideLayoutPart.SlideLayout = slideLayout;

        slideMaster.Append(
            new P.SlideLayoutIdList(
                new P.SlideLayoutId
                {
                    Id = 1U,
                    RelationshipId = slideMasterPart.GetIdOfPart(slideLayoutPart)
                }));

        slideMaster.Save();
        slideLayout.Save();

        var slideMasterIdList = new P.SlideMasterIdList(
            new P.SlideMasterId
            {
                Id = 2147483648U,
                RelationshipId = presentationPart.GetIdOfPart(slideMasterPart)
            });

        presentationPart.Presentation!.Append(slideMasterIdList);

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

    private static void AddTitleSlide(
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
            "Title Slide");

        AddBackground(context.ShapeTree, Navy);

        AddRectangle(
            context.ShapeTree,
            2U,
            "Accent Block",
            0.0,
            0.0,
            13.333,
            0.18,
            Blue,
            Blue);

        AddText(
            context.ShapeTree,
            3U,
            "Title",
            "Sprint Reporting Agent",
            0.75,
            1.0,
            8.6,
            0.65,
            3600,
            White,
            true);

        AddText(
            context.ShapeTree,
            4U,
            "Subtitle",
            "Excel to AI-powered stakeholder PowerPoint",
            0.78,
            1.75,
            8.6,
            0.35,
            1700,
            "BFDBFE",
            false);

        AddText(
            context.ShapeTree,
            5U,
            "Generated Date",
            $"Generated on {DateTime.Now:dd MMM yyyy}",
            0.78,
            2.2,
            5.5,
            0.3,
            1300,
            "CBD5E1",
            false);

        AddMetricCard(
            context.ShapeTree,
            6U,
            "Total Issues",
            metrics.TotalIssues.ToString(),
            "Parsed from uploaded Excel",
            0.78,
            3.25,
            2.75,
            1.25,
            Blue);

        AddMetricCard(
            context.ShapeTree,
            10U,
            "Completion",
            $"{metrics.CompletionPercentage}%",
            "Calculated in C#",
            3.82,
            3.25,
            2.75,
            1.25,
            Green);

        AddMetricCard(
            context.ShapeTree,
            14U,
            "Open Issues",
            metrics.OpenIssues.ToString(),
            "Pending delivery items",
            6.86,
            3.25,
            2.75,
            1.25,
            Orange);

        AddMetricCard(
            context.ShapeTree,
            18U,
            "AI Provider",
            aiInsights.ProviderUsed,
            "Insights generated after KPI aggregation",
            9.9,
            3.25,
            2.75,
            1.25,
            Purple);

        AddText(
            context.ShapeTree,
            22U,
            "Architecture Note",
            "Metrics are calculated deterministically in the backend. Only aggregated KPI data is sent to AI.",
            0.8,
            5.25,
            10.8,
            0.45,
            1500,
            "E2E8F0",
            false);

        AddFooter(context.ShapeTree, 23U);

        context.SlidePart.Slide.Save();
    }

    private static void AddExecutiveSummarySlide(
        PresentationPart presentationPart,
        SlideLayoutPart slideLayoutPart,
        P.SlideIdList slideIdList,
        uint slideId,
        AIInsightResult aiInsights,
        SprintMetrics metrics)
    {
        var context = CreateReportSlide(
            presentationPart,
            slideLayoutPart,
            slideIdList,
            slideId,
            "Executive Summary",
            "AI-generated view based on aggregated sprint KPIs");

        AddText(
            context.ShapeTree,
            20U,
            "Summary Text",
            aiInsights.ExecutiveSummary,
            0.75,
            1.55,
            7.45,
            1.9,
            1900,
            Navy,
            false);

        AddKpiStrip(context.ShapeTree, 30U, metrics, 0.75, 4.1);

        AddBulletedPanel(
            context.ShapeTree,
            50U,
            "Key Observations",
            aiInsights.Observations,
            8.55,
            1.55,
            3.95,
            2.15,
            Blue);

        AddBulletedPanel(
            context.ShapeTree,
            70U,
            "Primary Risks",
            aiInsights.Risks,
            8.55,
            4.05,
            3.95,
            2.15,
            Red);

        AddFooter(context.ShapeTree, 90U);

        context.SlidePart.Slide.Save();
    }

    private static void AddDeliverySlide(
        PresentationPart presentationPart,
        SlideLayoutPart slideLayoutPart,
        P.SlideIdList slideIdList,
        uint slideId,
        SprintMetrics metrics)
    {
        var context = CreateReportSlide(
            presentationPart,
            slideLayoutPart,
            slideIdList,
            slideId,
            "Delivery Metrics",
            "Completion, status distribution, and issue type distribution");

        AddKpiStrip(context.ShapeTree, 20U, metrics, 0.75, 1.35);

        AddDistributionPanel(
            context.ShapeTree,
            50U,
            "Status Distribution",
            metrics.StatusDistribution,
            0.75,
            3.2,
            5.85,
            2.9,
            Blue);

        AddDistributionPanel(
            context.ShapeTree,
            80U,
            "Issue Type Distribution",
            metrics.IssueTypeDistribution,
            6.95,
            3.2,
            5.55,
            2.9,
            Green);

        AddFooter(context.ShapeTree, 110U);

        context.SlidePart.Slide.Save();
    }

    private static void AddPriorityRiskSlide(
        PresentationPart presentationPart,
        SlideLayoutPart slideLayoutPart,
        P.SlideIdList slideIdList,
        uint slideId,
        SprintMetrics metrics,
        AIInsightResult aiInsights)
    {
        var context = CreateReportSlide(
            presentationPart,
            slideLayoutPart,
            slideIdList,
            slideId,
            "Priority & Risk",
            "Risk indicators derived from priority, age, and open work");

        AddMetricCard(
            context.ShapeTree,
            20U,
            "High Priority Open",
            metrics.HighPriorityOpenIssues.ToString(),
            "High or Highest priority and not Done",
            0.75,
            1.35,
            3.0,
            1.15,
            Red);

        AddMetricCard(
            context.ShapeTree,
            24U,
            "Open Issues",
            metrics.OpenIssues.ToString(),
            "All non-Done records",
            4.05,
            1.35,
            3.0,
            1.15,
            Orange);

        AddMetricCard(
            context.ShapeTree,
            28U,
            "Avg Open Age",
            $"{metrics.AverageIssueAgeDays}",
            "Days since created",
            7.35,
            1.35,
            3.0,
            1.15,
            Purple);

        AddDistributionPanel(
            context.ShapeTree,
            40U,
            "Priority Distribution",
            metrics.PriorityDistribution,
            0.75,
            3.0,
            5.75,
            3.1,
            Red);

        AddBulletedPanel(
            context.ShapeTree,
            70U,
            "Risk Notes",
            aiInsights.Risks,
            6.85,
            3.0,
            5.65,
            3.1,
            Orange);

        AddFooter(context.ShapeTree, 100U);

        context.SlidePart.Slide.Save();
    }

    private static void AddTeamAnalysisSlide(
        PresentationPart presentationPart,
        SlideLayoutPart slideLayoutPart,
        P.SlideIdList slideIdList,
        uint slideId,
        SprintMetrics metrics)
    {
        var context = CreateReportSlide(
            presentationPart,
            slideLayoutPart,
            slideIdList,
            slideId,
            "Team Analysis",
            "Workload and completed work distribution by assignee");

        AddDistributionPanel(
            context.ShapeTree,
            20U,
            "Issues Per Assignee",
            metrics.AssigneeDistribution,
            0.75,
            1.45,
            5.85,
            4.8,
            Blue,
            8);

        AddDistributionPanel(
            context.ShapeTree,
            60U,
            "Completed Work Per Assignee",
            metrics.CompletedWorkPerAssignee,
            6.95,
            1.45,
            5.55,
            4.8,
            Green,
            8);

        AddFooter(context.ShapeTree, 100U);

        context.SlidePart.Slide.Save();
    }

    private static void AddComponentAnalysisSlide(
        PresentationPart presentationPart,
        SlideLayoutPart slideLayoutPart,
        P.SlideIdList slideIdList,
        uint slideId,
        SprintMetrics metrics)
    {
        var context = CreateReportSlide(
            presentationPart,
            slideLayoutPart,
            slideIdList,
            slideId,
            "Component Analysis",
            "Issue concentration by system component");

        AddDistributionPanel(
            context.ShapeTree,
            20U,
            "Component Distribution",
            metrics.ComponentDistribution,
            0.75,
            1.45,
            6.25,
            4.8,
            Purple);

        AddInsightBox(
            context.ShapeTree,
            60U,
            "How to Read This",
            "Components with higher issue concentration should be reviewed for ownership, dependencies, or recurring implementation risks.",
            7.35,
            1.45,
            5.15,
            2.0,
            LightBlue,
            Blue);

        AddText(
            context.ShapeTree,
            70U,
            "Component Note",
            "This slide is useful for sprint review discussions because it shows where effort is concentrated without exposing raw ticket rows.",
            7.35,
            4.0,
            5.1,
            1.2,
            1500,
            GrayText,
            false);

        AddFooter(context.ShapeTree, 100U);

        context.SlidePart.Slide.Save();
    }

    private static void AddTechnicalDebtSlide(
        PresentationPart presentationPart,
        SlideLayoutPart slideLayoutPart,
        P.SlideIdList slideIdList,
        uint slideId,
        SprintMetrics metrics)
    {
        var context = CreateReportSlide(
            presentationPart,
            slideLayoutPart,
            slideIdList,
            slideId,
            "Technical Debt",
            "Technical debt indicators derived from issue labels");

        AddMetricCard(
            context.ShapeTree,
            20U,
            "tech-debt",
            GetDictionaryValue(metrics.LabelDistribution, "tech-debt").ToString(),
            "Debt-related labels",
            0.75,
            1.35,
            2.75,
            1.15,
            Purple);

        AddMetricCard(
            context.ShapeTree,
            24U,
            "performance",
            GetDictionaryValue(metrics.LabelDistribution, "performance").ToString(),
            "Performance related work",
            3.75,
            1.35,
            2.75,
            1.15,
            Orange);

        AddMetricCard(
            context.ShapeTree,
            28U,
            "hotfix",
            GetDictionaryValue(metrics.LabelDistribution, "hotfix").ToString(),
            "Urgent correction work",
            6.75,
            1.35,
            2.75,
            1.15,
            Red);

        AddMetricCard(
            context.ShapeTree,
            32U,
            "enhancement",
            GetDictionaryValue(metrics.LabelDistribution, "enhancement").ToString(),
            "Improvement work",
            9.75,
            1.35,
            2.75,
            1.15,
            Green);

        AddDistributionPanel(
            context.ShapeTree,
            50U,
            "All Label Distribution",
            metrics.LabelDistribution,
            0.75,
            3.1,
            11.75,
            3.0,
            Purple,
            10);

        AddFooter(context.ShapeTree, 90U);

        context.SlidePart.Slide.Save();
    }

    private static void AddAgingBacklogSlide(
        PresentationPart presentationPart,
        SlideLayoutPart slideLayoutPart,
        P.SlideIdList slideIdList,
        uint slideId,
        SprintMetrics metrics)
    {
        var context = CreateReportSlide(
            presentationPart,
            slideLayoutPart,
            slideIdList,
            slideId,
            "Aging & Backlog",
            "Open issue age and backlog health indicators");

        AddMetricCard(
            context.ShapeTree,
            20U,
            "Average Open Age",
            $"{metrics.AverageIssueAgeDays}",
            "Days",
            0.75,
            1.35,
            3.25,
            1.15,
            Purple);

        AddMetricCard(
            context.ShapeTree,
            24U,
            "Backlog Size",
            metrics.BacklogSize.ToString(),
            "Backlog status items",
            4.35,
            1.35,
            3.25,
            1.15,
            Orange);

        AddMetricCard(
            context.ShapeTree,
            28U,
            "Open Issues",
            metrics.OpenIssues.ToString(),
            "Pending work",
            7.95,
            1.35,
            3.25,
            1.15,
            Red);

        AddOldestIssuesPanel(
            context.ShapeTree,
            40U,
            "Oldest Open Issues",
            metrics.OldestOpenIssues,
            0.75,
            3.05,
            11.75,
            3.1,
            Red);

        AddFooter(context.ShapeTree, 90U);

        context.SlidePart.Slide.Save();
    }

    private static void AddRecommendationsSlide(
        PresentationPart presentationPart,
        SlideLayoutPart slideLayoutPart,
        P.SlideIdList slideIdList,
        uint slideId,
        AIInsightResult aiInsights)
    {
        var context = CreateReportSlide(
            presentationPart,
            slideLayoutPart,
            slideIdList,
            slideId,
            "Recommendations",
            "AI-generated actions for sprint stakeholders");

        AddBulletedPanel(
            context.ShapeTree,
            20U,
            "Recommended Actions",
            aiInsights.Recommendations,
            0.75,
            1.45,
            7.0,
            4.5,
            Green);

        AddInsightBox(
            context.ShapeTree,
            60U,
            "AI Usage Note",
            $"Provider used: {aiInsights.ProviderUsed}. The AI receives aggregated metrics only, not the full Excel dataset.",
            8.1,
            1.45,
            4.4,
            2.0,
            LightBlue,
            Blue);

        AddText(
            context.ShapeTree,
            70U,
            "Final Note",
            "Use this report as a sprint review starting point. Numeric KPIs are calculated in code, while narrative insights are generated by AI.",
            8.1,
            4.05,
            4.4,
            1.2,
            1500,
            GrayText,
            false);

        AddFooter(context.ShapeTree, 90U);

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

        AddRectangle(
            context.ShapeTree,
            2U,
            "Top Accent",
            0.0,
            0.0,
            13.333,
            0.14,
            Blue,
            Blue);

        AddText(
            context.ShapeTree,
            3U,
            "Slide Title",
            title,
            0.72,
            0.35,
            7.8,
            0.45,
            2800,
            Navy,
            true);

        AddText(
            context.ShapeTree,
            4U,
            "Slide Subtitle",
            subtitle,
            0.74,
            0.86,
            8.7,
            0.3,
            1200,
            GrayText,
            false);

        return context;
    }

    private static void AddKpiStrip(
        P.ShapeTree shapeTree,
        uint startId,
        SprintMetrics metrics,
        double x,
        double y)
    {
        AddMetricCard(
            shapeTree,
            startId,
            "Total",
            metrics.TotalIssues.ToString(),
            "Total issues",
            x,
            y,
            2.75,
            1.15,
            Blue);

        AddMetricCard(
            shapeTree,
            startId + 4,
            "Completed",
            metrics.CompletedIssues.ToString(),
            "Done issues",
            x + 3.0,
            y,
            2.75,
            1.15,
            Green);

        AddMetricCard(
            shapeTree,
            startId + 8,
            "Open",
            metrics.OpenIssues.ToString(),
            "Not Done",
            x + 6.0,
            y,
            2.75,
            1.15,
            Orange);

        AddMetricCard(
            shapeTree,
            startId + 12,
            "Completion",
            $"{metrics.CompletionPercentage}%",
            "Sprint progress",
            x + 9.0,
            y,
            2.75,
            1.15,
            Purple);
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
        AddRectangle(
            shapeTree,
            startId,
            $"{label} Card",
            x,
            y,
            width,
            height,
            LightGray,
            BorderGray);

        AddRectangle(
            shapeTree,
            startId + 1,
            $"{label} Accent",
            x,
            y,
            0.08,
            height,
            accentColor,
            accentColor);

        AddText(
            shapeTree,
            startId + 2,
            $"{label} Value",
            value,
            x + 0.22,
            y + 0.18,
            width - 0.35,
            0.36,
            2500,
            accentColor,
            true);

        AddText(
            shapeTree,
            startId + 3,
            $"{label} Label",
            $"{label}\n{caption}",
            x + 0.22,
            y + 0.58,
            width - 0.35,
            0.42,
            1000,
            GrayText,
            false);
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
        AddRectangle(
            shapeTree,
            startId,
            $"{title} Panel",
            x,
            y,
            width,
            height,
            LightGray,
            BorderGray);

        AddRectangle(
            shapeTree,
            startId + 1,
            $"{title} Header",
            x,
            y,
            width,
            0.42,
            accentColor,
            accentColor);

        AddText(
            shapeTree,
            startId + 2,
            $"{title} Header Text",
            title,
            x + 0.22,
            y + 0.1,
            width - 0.45,
            0.25,
            1300,
            White,
            true);

        var lines = BuildDistributionLines(values, take);

        AddText(
            shapeTree,
            startId + 3,
            $"{title} Body",
            string.Join("\n", lines),
            x + 0.28,
            y + 0.65,
            width - 0.55,
            height - 0.85,
            1300,
            Navy,
            false);
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
        AddRectangle(
            shapeTree,
            startId,
            $"{title} Panel",
            x,
            y,
            width,
            height,
            LightGray,
            BorderGray);

        AddRectangle(
            shapeTree,
            startId + 1,
            $"{title} Header",
            x,
            y,
            width,
            0.42,
            accentColor,
            accentColor);

        AddText(
            shapeTree,
            startId + 2,
            $"{title} Header Text",
            title,
            x + 0.22,
            y + 0.1,
            width - 0.45,
            0.25,
            1300,
            White,
            true);

        var text = values.Count == 0
            ? "No insights available."
            : string.Join("\n", values.Select(value => $"• {value}"));

        AddText(
            shapeTree,
            startId + 3,
            $"{title} Body",
            text,
            x + 0.28,
            y + 0.65,
            width - 0.55,
            height - 0.85,
            1250,
            Navy,
            false);
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
        AddRectangle(
            shapeTree,
            startId,
            $"{title} Box",
            x,
            y,
            width,
            height,
            fillColor,
            accentColor);

        AddText(
            shapeTree,
            startId + 1,
            $"{title} Heading",
            title,
            x + 0.25,
            y + 0.18,
            width - 0.5,
            0.28,
            1400,
            accentColor,
            true);

        AddText(
            shapeTree,
            startId + 2,
            $"{title} Body",
            body,
            x + 0.25,
            y + 0.58,
            width - 0.5,
            height - 0.75,
            1200,
            Navy,
            false);
    }

    private static void AddOldestIssuesPanel(
        P.ShapeTree shapeTree,
        uint startId,
        string title,
        List<OldestOpenIssueMetric> issues,
        double x,
        double y,
        double width,
        double height,
        string accentColor)
    {
        AddRectangle(
            shapeTree,
            startId,
            $"{title} Panel",
            x,
            y,
            width,
            height,
            LightGray,
            BorderGray);

        AddRectangle(
            shapeTree,
            startId + 1,
            $"{title} Header",
            x,
            y,
            width,
            0.42,
            accentColor,
            accentColor);

        AddText(
            shapeTree,
            startId + 2,
            $"{title} Header Text",
            title,
            x + 0.22,
            y + 0.1,
            width - 0.45,
            0.25,
            1300,
            White,
            true);

        var text = issues.Count == 0
            ? "No open aging issues available."
            : string.Join(
                "\n",
                issues.Take(5).Select(issue =>
                    $"{issue.IssueKey}    {issue.Priority}    {issue.Status}    {issue.AgeDays} days"));

        AddText(
            shapeTree,
            startId + 3,
            $"{title} Body",
            text,
            x + 0.28,
            y + 0.7,
            width - 0.55,
            height - 0.9,
            1250,
            Navy,
            false);
    }

    private static void AddFooter(
        P.ShapeTree shapeTree,
        uint id)
    {
        AddRectangle(
            shapeTree,
            id,
            "Footer Line",
            0.72,
            6.95,
            11.9,
            0.01,
            BorderGray,
            BorderGray);

        AddText(
            shapeTree,
            id + 1,
            "Footer Text",
            $"AI Sprint Reporting Agent  |  Generated {DateTime.Now:dd MMM yyyy}",
            0.75,
            7.05,
            7.5,
            0.22,
            850,
            GrayText,
            false);

        AddText(
            shapeTree,
            id + 2,
            "Footer Note",
            "Metrics calculated in C#; AI used only for narrative insights.",
            8.2,
            7.05,
            4.25,
            0.22,
            850,
            GrayText,
            false);
    }

    private static void AddBackground(
        P.ShapeTree shapeTree,
        string color)
    {
        AddRectangle(
            shapeTree,
            1000U,
            "Background",
            0,
            0,
            13.333,
            7.5,
            color,
            color);
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
        var lines = text
            .Replace("\r\n", "\n")
            .Split('\n');

        foreach (var line in lines)
        {
            yield return new A.Paragraph(
                new A.Run(
                    new A.RunProperties
                    {
                        FontSize = fontSize,
                        Bold = bold
                    },
                    new A.SolidFill(
                        new A.RgbColorModelHex
                        {
                            Val = color
                        }),
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

        var maxValue = values.Values.Max();

        return values
            .Take(take)
            .Select(item =>
            {
                var barLength = maxValue == 0
                    ? 0
                    : Math.Max(1, (int)Math.Round((double)item.Value / maxValue * 14));

                var bar = new string('█', barLength);

                return $"{item.Key}: {item.Value}  {bar}";
            })
            .ToList();
    }

    private static int GetDictionaryValue(
        Dictionary<string, int> values,
        string key)
    {
        return values.TryGetValue(key, out var value)
            ? value
            : 0;
    }

    private static long InchesToEmu(double inches)
    {
        return (long)(inches * EmuPerInch);
    }

    private sealed record SlideContext(
        SlidePart SlidePart,
        P.ShapeTree ShapeTree);
}