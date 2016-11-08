using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ModelClassLibrary
{
    public class UploadFileMaster
    {
        public Int32 UploadFileID { get; set; }
        public Int32 UploadType { get; set; }
        public string UploadTypeCode { get; set; }
        public string TemplateFileName { get; set; }
        public string SourceColumn { get; set; }
        public string DestinationColumn { get; set; }
        public string MandatoryColumn { get; set; }
        public string TableName { get; set; }
        public string SheetName { get; set; }
        public bool UploadFlag { get; set; }
        public string ExtraProcedure { get; set; }
        public bool DisplayFlag { get; set; }
    }
}