using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SPPricing.Models
{
    public class BasketInstruments
    {
        public Int32 UnderlyingID { get; set; }
        public string BasketInstrument { get; set; }
        public Int32 Weightage { get; set; }
    }
}