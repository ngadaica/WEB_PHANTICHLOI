namespace WEB_PHANTICHLOI.Migrations
{
    using System;
    using System.Data.Entity.Migrations;
    
    public partial class AddMeasurementWaitHours : DbMigration
    {
        public override void Up()
        {
            AddColumn("dbo.ErrorAnalyses", "MeasurementWaitHours", c => c.Double());
        }
        
        public override void Down()
        {
            DropColumn("dbo.ErrorAnalyses", "MeasurementWaitHours");
        }
    }
}
