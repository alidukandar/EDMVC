using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ModelClassLibrary
{
    public class PricerGraphAPIDetails
    {
        public string ProductID { get; set; }

        public double Strike { get; set; }
        public double Strike11 { get; set; }
        public double Strike12 { get; set; }
        public double Strike21 { get; set; }
        public double Strike22 { get; set; }

        public double PutSpreadStrike1 { get; set; }
        public double PutSpreadStrike2 { get; set; }

        public double PutShortStrike1 { get; set; }
        public double PutShortStrike2 { get; set; }

        public double PutSpreadPR { get; set; }
        public double PutPR { get; set; }

        public double BelowStrikeCoupon { get; set; }
        public double AfterStrikeCoupon { get; set; }

        public Int32 RedemptionDays { get; set; }
        public string GraphType { get; set; }
        public double PR { get; set; }

        public string ProductType { get; set; }
        public string CallOptionType { get; set; }
        public string PutOptionType { get; set; }
        public double FixedCouponValue { get; set; }
    }
}
