namespace WEB_PHANTICHLOI.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class InitialCreate : DbMigration
    {
        public override void Up()
        {
            CreateTable(
                "dbo.ErrorAnalyses",
                c => new
                    {
                        Id = c.Int(nullable: false, identity: true),
                        Phenomenon = c.String(),
                        ProblemContent = c.String(),
                        Category = c.String(),
                        Model = c.String(),
                        StageClassification = c.String(),
                        OccurrenceProcess = c.String(),
                        Line = c.String(),
                        OccurrenceCount = c.Int(),
                        OccurrenceDate = c.DateTime(),
                        ShipmentStop = c.String(),
                        Investigator = c.String(),
                        BAction = c.String(),
                        ChronicDefect = c.String(),
                        LineStopTime = c.String(),
                        InvestigationContent = c.String(),
                        PartCode1 = c.String(),
                        PartName1 = c.String(),
                        PartCode2 = c.String(),
                        PartName2 = c.String(),
                        PartCode3 = c.String(),
                        PartName3 = c.String(),
                        InvestigationProgress = c.String(),
                        CauseClassification = c.String(),
                        DetailedCause = c.String(),
                        InvestigationStartTime = c.DateTime(),
                        InvestigationEndTime = c.DateTime(),
                        DaysOff = c.Int(),
                        NightShiftDays = c.Int(),
                        InvestigationHours = c.Double(),
                        AnalysisContent = c.String(),
                        NextAction = c.String(),
                        TemporaryMeasureClassification = c.String(),
                        TemporaryMeasureDate = c.DateTime(),
                        TemporaryMeasureDetail = c.String(),
                        PermanentMeasureClassification = c.String(),
                        PermanentMeasureDate = c.DateTime(),
                        PermanentMeasureDetail = c.String(),
                        AttachmentPath = c.String(),
                        Factory = c.String(),
                        Status = c.String(),
                        OccurrencePeriod = c.String(),
                        Team = c.String(),
                        PersonInCharge = c.String(),
                        ImageStatus = c.String(),
                        CauseDescription = c.String(),
                        NextActionSummary = c.String(),
                        DefectClassification = c.String(),
                        TemporaryMeasureSummary = c.String(),
                        PermanentMeasureSummary = c.String(),
                        InvestigationCompletionTime = c.String(),
                        MeasureCompletionTime = c.String(),
                        BActionStatus = c.String(),
                        Sharing = c.String(),
                        Report = c.String(),
                        UpdatedDate = c.DateTime(),
                        UpdatedBy = c.String(),
                    })
                .PrimaryKey(t => t.Id);
            
            CreateTable(
                "dbo.ErrorAnalysisLookups",
                c => new
                    {
                        Id = c.Int(nullable: false, identity: true),
                        GroupCode = c.String(nullable: false, maxLength: 100),
                        Value = c.String(nullable: false, maxLength: 200),
                        DisplayText = c.String(nullable: false, maxLength: 200),
                        SortOrder = c.Int(nullable: false),
                        IsActive = c.Boolean(nullable: false),
                    })
                .PrimaryKey(t => t.Id);
            
        }
        
        public override void Down()
        {
            DropTable("dbo.ErrorAnalysisLookups");
            DropTable("dbo.ErrorAnalyses");
        }
    }
}
