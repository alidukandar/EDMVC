using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ModelClassLibrary
{
    public class PricerAPIDetails
    {
        public string ProductID { get; set; }
        public string PricerType { get; set; }
        public double FixedCoupon { get; set; }
        public double EdelweissBuiltIn { get; set; }
        public double DistributorBuiltIn { get; set; }

        public Int32 InitialAveragingMonth { get; set; }
        public Int32 InitialAveragingDaysDiff { get; set; }

        public Int32 FinalAveragingMonth { get; set; }
        public Int32 FinalAveragingDaysDiff { get; set; }

        public Int32 RedemptionPeriodDays { get; set; }
        public double OptionTenureMonth { get; set; }
        public string OptionType { get; set; }

        public string CouponScenario1 { get; set; }
        public string CouponScenario2 { get; set; }

        public double ContingentCouponRate { get; set; }
        public Int32 ObservationDays { get; set; }
        public Int32 EarlyContingentRedemptionDays { get; set; }
        public double AutocallLevel { get; set; }

        public string UnderlyingName { get; set; }
        public Int32 UnderlyingID { get; set; }

        public double PR1 { get; set; }
        public double PR2 { get; set; }
    }
}