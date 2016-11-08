using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ModelClassLibrary
{
    public class LookupMaster
    {
        public Int32 LookupID { get; set; }
        public string LookupCode { get; set; }
        public Int32 LookupType { get; set; }
        public string LookupDescription { get; set; }
        public bool IsActive { get; set; }
    }
}