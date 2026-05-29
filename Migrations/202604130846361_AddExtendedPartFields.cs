namespace WEB_PHANTICHLOI.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class AddExtendedPartFields : DbMigration
    {
        public override void Up()
        {
            AddColumn("dbo.ErrorAnalyses", "PartSupplier1", c => c.String());
            AddColumn("dbo.ErrorAnalyses", "PartCav1", c => c.String());
            AddColumn("dbo.ErrorAnalyses", "PartLot1", c => c.String());
            AddColumn("dbo.ErrorAnalyses", "PartSupplier2", c => c.String());
            AddColumn("dbo.ErrorAnalyses", "PartCav2", c => c.String());
            AddColumn("dbo.ErrorAnalyses", "PartLot2", c => c.String());
            AddColumn("dbo.ErrorAnalyses", "PartSupplier3", c => c.String());
            AddColumn("dbo.ErrorAnalyses", "PartCav3", c => c.String());
            AddColumn("dbo.ErrorAnalyses", "PartLot3", c => c.String());
            AddColumn("dbo.ErrorAnalyses", "PartCode4", c => c.String());
            AddColumn("dbo.ErrorAnalyses", "PartName4", c => c.String());
            AddColumn("dbo.ErrorAnalyses", "PartSupplier4", c => c.String());
            AddColumn("dbo.ErrorAnalyses", "PartCav4", c => c.String());
            AddColumn("dbo.ErrorAnalyses", "PartLot4", c => c.String());
            AddColumn("dbo.ErrorAnalyses", "PartCode5", c => c.String());
            AddColumn("dbo.ErrorAnalyses", "PartName5", c => c.String());
            AddColumn("dbo.ErrorAnalyses", "PartSupplier5", c => c.String());
            AddColumn("dbo.ErrorAnalyses", "PartCav5", c => c.String());
            AddColumn("dbo.ErrorAnalyses", "PartLot5", c => c.String());
        }
        
        public override void Down()
        {
            DropColumn("dbo.ErrorAnalyses", "PartLot5");
            DropColumn("dbo.ErrorAnalyses", "PartCav5");
            DropColumn("dbo.ErrorAnalyses", "PartSupplier5");
            DropColumn("dbo.ErrorAnalyses", "PartName5");
            DropColumn("dbo.ErrorAnalyses", "PartCode5");
            DropColumn("dbo.ErrorAnalyses", "PartLot4");
            DropColumn("dbo.ErrorAnalyses", "PartCav4");
            DropColumn("dbo.ErrorAnalyses", "PartSupplier4");
            DropColumn("dbo.ErrorAnalyses", "PartName4");
            DropColumn("dbo.ErrorAnalyses", "PartCode4");
            DropColumn("dbo.ErrorAnalyses", "PartLot3");
            DropColumn("dbo.ErrorAnalyses", "PartCav3");
            DropColumn("dbo.ErrorAnalyses", "PartSupplier3");
            DropColumn("dbo.ErrorAnalyses", "PartLot2");
            DropColumn("dbo.ErrorAnalyses", "PartCav2");
            DropColumn("dbo.ErrorAnalyses", "PartSupplier2");
            DropColumn("dbo.ErrorAnalyses", "PartLot1");
            DropColumn("dbo.ErrorAnalyses", "PartCav1");
            DropColumn("dbo.ErrorAnalyses", "PartSupplier1");
        }
    }
}
