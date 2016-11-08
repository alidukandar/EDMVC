using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ModelClassLibrary
{
    public class ResponseDetails
    {

        public int ResponseCode { get; set; }
        public string ResponseMsg { get; set; }
        public string Result { get; set; }

        public UnderlyingList UnderlyingListData { get; set; }
        public PricerAPIDetailsList PricerAPIDetailsListData { get; set; }
        public PricerGraphAPIDetailsList PricerGraphAPIDetailsListData { get; set; }
    }

    public class UnderlyingList
    {
        public List<Underlying> lstUnderlyingList { get; set; }
    }

    public class PricerAPIDetailsList
    {
        public List<PricerAPIDetails> lstPricerAPIDetailsList { get; set; }
    }

    public class PricerGraphAPIDetailsList
    {
        public List<PricerGraphAPIDetails> lstPricerGraphAPIDetailsList { get; set; }
    }
}