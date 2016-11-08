using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using DotNet.Highcharts;
using System.Web.Mvc;

namespace SPPricing.Models
{
    public class CallBinary
    {
        public CallBinary()
        {
            CallBinaryDetailsList = new List<CallBinaryDetails>();
            StatusList = new List<LookupMaster>();
        }

        public Int32 CallBinaryID { get; set; }
        public string ProductID { get; set; }
        public string Distributor { get; set; }
        public double EdelweissBuiltIn { get; set; }
        public double DistributorBuiltIn { get; set; }
        public Int32 UnderlyingID { get; set; }
        public string UnderlyingName { get; set; }
        public double Remaining { get; set; }

        public string FilterUnderlying { get; set; }

        public bool IsChildQuote { get; set; }

        public double FixedCoupon { get; set; }
        public double FixedCouponIRR { get; set; }
        public double MaxCoupon { get; set; }
        public double MaxCouponIRR { get; set; }

        public double DeploymentRate { get; set; }
        public double CustomDeploymentRate { get; set; }

        public double OptionTenure { get; set; }
        public double RedemptionPeriodMonth { get; set; }
        public Int32 RedemptionPeriodDays { get; set; }
        public string ProductTenure { get; set; }

        public Int32 InitialAveragingMonth { get; set; }
        public Int32 InitialAveragingDaysDiff { get; set; }

        public Int32 FinalAveragingMonth { get; set; }
        public Int32 FinalAveragingDaysDiff { get; set; }

        public string SalesComments { get; set; }
        public string TradingComments { get; set; }
        public string CouponScenario { get; set; }

        public double TotalOptionPrice { get; set; }
        public double NetRemaining { get; set; }

        public Int32 CreatedBy { get; set; }
        public DateTime CreatedOn { get; set; }
        public string ConfirmedOn { get; set; }
        public Int32 ModifiedBy { get; set; }
        public DateTime ModifiedOn { get; set; }

        public DateTime FromDate { get; set; }
        public DateTime ToDate { get; set; }

        public Highcharts CallBinaryChart { get; set; }

        public List<CallBinaryDetails> CallBinaryDetailsList { get; set; }
        public string ParentProductID { get; set; }

        public List<LookupMaster> StatusList { get; set; }
        public string FilterStatus { get; set; }

        public double Strike1 { get; set; }
        public double CouponRise { get; set; }

        [AllowHtml]
        public string ExportCallStrike1Summary { get; set; }

        [AllowHtml]
        public string ExportCallStrike2Summary { get; set; }
    }

    public class CallBinaryDetails
    {
        public Int32 CallBinaryDetailsID { get; set; }
        public Int32 CallBinaryID { get; set; }
        public Int32 OptionTypeID { get; set; }

        public double Strike1 { get; set; }
        public double CouponRise { get; set; }
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