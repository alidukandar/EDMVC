using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SPPricing.Models
{
    public class Autocall
    {
        public Int32 AutocallID { get; set; }
        public string ProductID { get; set; }
        public string Distributor { get; set; }
        public double EdelweissBuiltIn { get; set; }
        public double DistributorBuiltIn { get; set; }

        public Int32 UnderlyingID { get; set; }
        public List<Underlying> UnderlyingList { get; set; }
        public string FilterUnderlying { get; set; }

        public double FixedCoupon { get; set; }
        public double FixedCouponIRR { get; set; }

        public double DeploymentRate { get; set; }
        public double CustomDeploymentRate { get; set; }

        public double Remaining { get; set; }

        public Int32 ImpliedVolatilityID { get; set; }
        public List<LookupMaster> ImpliedVolatilityList { get; set; }
        public string FilterImpliedVolatility { get; set; }

        public Int32 ObservationFrequencyID { get; set; }
        public List<LookupMaster> ObservationFrequencyList { get; set; }
        public string FilterObservationFrequency { get; set; }

        public bool IsPrincipalProtected { get; set; }
        public bool IsDiscountingApplicable { get; set; }

        public double NonPPLevel { get; set; }
        public double RollCost { get; set; }

        public DateTime ObservationStartDate { get; set; }
        public DateTime ObservationEndDate { get; set; }

        public double OptionTenure { get; set; }
        public double RedemptionPeriodMonth { get; set; }
        public Int32 RedemptionPeriodDays { get; set; }
        public string ProductTenure { get; set; }

        public double BondPrice { get; set; }
        public double OptionPrice { get; set; }
        public double TotalBuiltIn { get; set; }

        public Int32 NoOfSimulations { get; set; }
        public Int32 Count { get; set; }

        public Int32 EarlyRedemptionPaymentGap { get; set; }
        public double InterestRateHit { get; set; }
        public double InterestRateHitCalculation { get; set; }
        public double AutocallLevel { get; set; }
        //public double NonPPLevel { get; set; }
        public double CouponIfHit { get; set; }

        public double ExpectedTimeToMaturity { get; set; }
        public double AverageDeploymentRate { get; set; }
        public double ExpectedBondPrice { get; set; }
        public double BondPriceCalculation { get; set; }
        public double PriceInterestRateHit { get; set; }
        public double Coupon { get; set; }

        public string SalesComments { get; set; }
        public string TradingComments { get; set; }
        public string CouponScenario1 { get; set; }
        public string CouponScenario2 { get; set; }
        public double IRR { get; set; }
    }

    public class AutocallSimulation
    {
        public Int32 Month { get; set; }
        public double AutocallLevel { get; set; }
        public Int32 Buffer { get; set; }
        public double UnderlyingLevel { get; set; }
        public bool IsAutocalled { get; set; }
        public double DeploymentRate { get; set; }
        public double InterestRateHit { get; set; }
        public double CouponIfHit { get; set; }
    }
}