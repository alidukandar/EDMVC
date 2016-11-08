using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ModelClassLibrary
{
    public class Underlying
    {
        public Underlying()
        {
            UnderlyingTypeList = new List<LookupMaster>();
            StandardList = new List<LookupMaster>();
            SubTypeList = new List<LookupMaster>();
            IVRFCategoryList = new List<LookupMaster>();
            UnderLyingList = new List<Underlying>();
            NameList = new List<Underlying>();
            ShortNameList = new List<Underlying>();
            TypeList = new List<Underlying>();
            UnderlyingStandardList = new List<Underlying>();
            UnderlyingSubTypeList = new List<Underlying>();
        }

        public Int32 UnderlyingID { get; set; }
        public string UnderlyingShortName { get; set; }
        public string UnderlyingName { get; set; }

        public Int32 UnderlyingTypeID { get; set; }
        public string UnderlyingType { get; set; }
        public List<LookupMaster> UnderlyingTypeList { get; set; }
        public Int32 FilterUnderlyingType { get; set; }

        public Int32 Standard { get; set; }
        public string StandardName { get; set; }
        public List<LookupMaster> StandardList { get; set; }
        public Int32 FilterStandard { get; set; }

        public Int32 SubType { get; set; }
        public string SubTypeName { get; set; }
        public List<LookupMaster> SubTypeList { get; set; }
        public Int32 FilterSubType { get; set; }


        public List<LookupMaster> CSUCategoryList { get; set; }
        public Int32 FilterCSUCategory { get; set; }

        public List<LookupMaster> PSUCategoryList { get; set; }
        public Int32 FilterPSUCategory { get; set; }

        public List<LookupMaster> IVRFCategoryList { get; set; }
        public Int32 FilterIVRFCategory { get; set; }

        public List<Underlying> UnderLyingList { get; set; }
        public Int32 FilterUnderLyingCategory { get; set; }

        public List<LookupMaster> RCTypeList { get; set; }
        public Int32 FilterRCType { get; set; }

        public List<Underlying> RCUnderLyingList { get; set; }
        public Int32 FilterRCUnderLyingCategory { get; set; }

        public List<LookupMaster> LVTypeList { get; set; }
        public Int32 FilterLVType { get; set; }

        public List<Underlying> LVUnderLyingList { get; set; }
        public Int32 FilterLVUnderLyingCategory { get; set; }

        public List<Underlying> NameList { get; set; }
        public string FilterName { get; set; }

        public List<Underlying> ShortNameList { get; set; }
        public string FilterShortName { get; set; }

        public List<Underlying> TypeList { get; set; }
        public string FilterType { get; set; }

        public List<Underlying> UnderlyingStandardList { get; set; }
        public string FilterUnderlyingStandard { get; set; }

        public List<Underlying> UnderlyingSubTypeList { get; set; }
        public string FilterUnderlyingSubType { get; set; }

        #region CSU


        public List<Underlying> CSUAUnderLyingList { get; set; }
        public Int32 FilterCSUAdjustment { get; set; }

        public List<Underlying> CSUTUnderLyingList { get; set; }
        public Int32 FilterCSUThreshold { get; set; }

        public List<Underlying> CSUMinimumUnderLyingList { get; set; }
        public Int32 FilterCSUMinimum { get; set; }

        #endregion

        #region PSU


        public List<Underlying> PSUAUnderLyingList { get; set; }
        public Int32 FilterPSUAdjustment { get; set; }

        public List<Underlying> PSUSkewUnderLyingList { get; set; }
        public Int32 FilterPSUSkew { get; set; }

        public List<Underlying> PSUMinimumUnderLyingList { get; set; }
        public Int32 FilterPSUMinimum { get; set; }

        #endregion

        //public List<Underlying> 

        public UploadFileMaster BasketCorrelation { get; set; }
        public UploadFileMaster ImpliedVolatility { get; set; }
        public UploadFileMaster RollCost { get; set; }
        public UploadFileMaster LocaleVolatilitySurface { get; set; }
        public UploadFileMaster CallAdjustmentSurface { get; set; }
        public UploadFileMaster CallThreshold { get; set; }
        public UploadFileMaster CallMinimumIV { get; set; }
        public UploadFileMaster PutAdjustmentSurface { get; set; }
        public UploadFileMaster PutSkewAdjustment { get; set; }
        public UploadFileMaster PutMinimumIV { get; set; }



    }
}