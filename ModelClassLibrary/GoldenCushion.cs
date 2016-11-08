using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DotNet.Highcharts;
using System.Web.Mvc;

namespace SPPricing.Models
{
    public class GoldenCushion
    {

        public GoldenCushion()
        {
            GoldenCushionDetailsList = new List<GoldenCushionDetails>();
            StatusList = new List<LookupMaster>();
        }

        public Int32 GoldenCushionID { get; set; }
        public string ProductID { get; set; }
        public string Distributor { get; set; }
        public double EdelweissBuiltIn { get; set; }
        public double DistributorBuiltIn { get; set; }

        public Int32 UnderlyingID { get; set; }
        public string UnderlyingName { get; set; }
        public List<Underlying> UnderlyingList { get; set; }
        public string FilterUnderlying { get; set; }

        public double Remaining { get; set; }

        public double FixedCoupon { get; set; }
        public double FixedCouponIRR { get; set; }
        public double LowerCoupon { get; set; }
        public double LowerCouponIRR { get; set; }

        public double DeploymentRate { get; set; }
        public double CustomDeploymentRate { get; set; }

        public bool IsPrincipalProtected { get; set; }

        public Int32 OptionTenure { get; set; }
        public Int32 RedemptionPeriodMonth { get; set; }
        public Int32 RedemptionPeriodDays { get; set; }
        public string ProductTenure { get; set; }

        public Int32 InitialAveragingMonth { get; set; }
        public Int32 InitialAveragingDaysDiff { get; set; }

        public Int32 FinalAveragingMonth { get; set; }
        public Int32 FinalAveragingDaysDiff { get; set; }

        public string CouponScenario1 { get; set; }
        public string CouponScenario2 { get; set; }

        public double TotalOptionPrice { get; set; }
        public double NetRemaining { get; set; }

        public string SalesComments { get; set; }
        public string TradingComments { get; set; }

        public Int32 CreatedBy { get; set; }
        public DateTime CreatedOn { get; set; }
        public string ConfirmedOn { get; set; }
        public Int32 ModifiedBy { get; set; }
        public DateTime ModifiedOn { get; set; }

        public double Strike1 { get; set; }
        public double Strike2 { get; set; }

        public DateTime FromDate { get; set; }
        public DateTime ToDate { get; set; }

        public List<GoldenCushionDetails> GoldenCushionDetailsList { get; set; }

        public Highcharts GoldenCushionChart { get; set; }
        public string ParentProductID { get; set; }

        public List<LookupMaster> StatusList { get; set; }
        public string FilterStatus { get; set; }

        public bool IsChildQuote { get; set; }

        [AllowHtml]
        public string ExportPutSpreadStrike1Summary { get; set; }

        [AllowHtml]
        public string ExportPutSpreadStrike2Summary { get; set; }

        [AllowHtml]
        public string ExportPutStrike1Summary { get; set; }

        [AllowHtml]
        public string ExportPutStrike2Summary { get; set; }

        public class GoldenCushionDetails
        {
            public Int32 GoldenCushionDetailsID { get; set; }
            public Int32 GoldenCushionID { get; set; }
            public Int32 OptionTypeID { get; set; }

            public double Strike1 { get; set; }
            public double Strike2 { get; set; }
            public double ParticipatoryRratio { get; set; }
            public double Price { get; set; }
            public double DiscountedPrice { get; set; }
            public double PRAdjustedPrice { get; set; }

            public double IV1 { get; set; }
            public double CustomIV1 { get; set; }
            public double RF1 { get; set; }
            public double CustomRF1 { get; set; }

            public double IV2 { get; set; }
            public double CustomIV2 { get; set; }
            public double RF2 { get; set; }
            public double CustomRF2 { get; set; }
        }

    }
}