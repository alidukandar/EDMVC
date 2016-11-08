using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using SPPricing.Models;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Data;
using System.Runtime.InteropServices;
using System.Data.SqlClient;
using System.Data.Objects;
using OfficeOpenXml;

namespace SPPricing.Controllers
{
    public class SPNoteMTMController : Controller
    {
        //
        // GET: /SPNoteMTM/

        SP_PRICINGEntities objSP_PRICINGEntities = new SP_PRICINGEntities();

        public ActionResult Index()
        {
            return View();
        }

        public ActionResult CustomIV()
        {
            return View();
        }

        [HttpGet]
        public ActionResult MTMReport()
        {
            try
            {
                if (ValidateSession())
                {
                    SPNoteMTM obj = new SPNoteMTM();

                    #region Menu Access By on Role

                    Int32 intResult = 0;
                    // bool PPorNonPP = false;

                    UserMaster objUserMaster = new UserMaster();
                    objUserMaster = (UserMaster)Session["LoggedInUser"];

                    var Result = objSP_PRICINGEntities.VALIDATE_MENU_BY_ROLE(objUserMaster.RoleID, "MTMR");
                    intResult = Convert.ToInt32(Result.SingleOrDefault());

                    if (intResult == 0)
                        return RedirectToAction("UserNotAuthorize", "Login");

                    #endregion

                    obj.IsFormula = false;
                    return View();
                }
                else
                {
                    return RedirectToAction("Login", "Login");
                }
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "SPNoteMTMController", "MTMReport Get", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        [HttpPost]
        public ActionResult MTMReport(SPNoteMTM objSPNoteMTM, string Command, FormCollection objFormCollection)
        {
            try
            {
                if (ValidateSession())
                {
                    if (Command == "GenerateReport")
                    {
                        FetchUploadFileMasterList();

                        string strReportDate = objFormCollection["ReportDate"];

                        objSPNoteMTM.ReportDate = Convert.ToDateTime(strReportDate.Substring(6, 4) + "-" + strReportDate.Substring(0, 2) + "-" + strReportDate.Substring(3, 2));

                        #region Fetch SP Portal Details
                        System.Net.ServicePointManager.ServerCertificateValidationCallback += (se, cert, chain, sslerror) =>
                        {
                            return true;
                        };

                        //System.Data.DataTable dtProductData;
                        List<UploadFileMaster> UploadFileMasterList = (List<UploadFileMaster>)Session["UploadFileMasterList"];
                        //UploadFileMaster objUploadFileMaster;
                        //string strSourceColumn;
                        //string[] arrSourceColumn = null;
                        //string strDestinationColumn;
                        //string[] arrDestinationColumn = null;
                        //string strTableName;
                        //System.Data.DataTable dtData;

                        string strMyConnection = Convert.ToString(System.Configuration.ConfigurationManager.ConnectionStrings["SP_PRICINGConnectionString"]);

                        #region SP Protal Product Details
                        //using (var client = new SPPricingProductDetails.GetDataForPricerSoapClient("GetDataForPricerSoap"))
                        //{
                        //    dtProductData = client.GetRiskDataForPricer();
                        //}

                        //objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "SPPD"; });

                        //strSourceColumn = objUploadFileMaster.SourceColumn;

                        //if (strSourceColumn != "")
                        //    arrSourceColumn = strSourceColumn.Split('|');

                        //dtData = new System.Data.DataTable();

                        //for (int i = 0; i < arrSourceColumn.Length; i++)
                        //{
                        //    dtData.Columns.Add(arrSourceColumn[i]);
                        //}

                        //strDestinationColumn = objUploadFileMaster.DestinationColumn;

                        //if (strDestinationColumn != "")
                        //    arrDestinationColumn = strDestinationColumn.Split('|');

                        //strTableName = objUploadFileMaster.TableName;

                        //if (arrSourceColumn != null && arrDestinationColumn != null && arrSourceColumn.Length == arrDestinationColumn.Length)
                        //{
                        //    using (SqlBulkCopy bulkCopy = new SqlBulkCopy(strMyConnection))
                        //    {
                        //        bulkCopy.DestinationTableName = strTableName;

                        //        for (int i = 0; i < arrSourceColumn.Length; i++)
                        //        {
                        //            bulkCopy.ColumnMappings.Add(arrSourceColumn[i], arrDestinationColumn[i]);
                        //        }
                        //        bulkCopy.WriteToServer(dtProductData);
                        //    }

                        //    DataSet dsIV = new DataSet();
                        //    dsIV = General.ExecuteDataSet(objUploadFileMaster.ExtraProcedure);
                        //}
                        #endregion

                        #region SP Protal Buyback Details
                        //using (var client = new SPPricingProductDetails.GetDataForPricerSoapClient("GetDataForPricerSoap"))
                        //{
                        //    dtProductData = client.GetBuybackDataForPricer();
                        //}

                        //objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "SPBD"; });

                        //strSourceColumn = objUploadFileMaster.SourceColumn;

                        //if (strSourceColumn != "")
                        //    arrSourceColumn = strSourceColumn.Split('|');

                        //dtData = new System.Data.DataTable();

                        //for (int i = 0; i < arrSourceColumn.Length; i++)
                        //{
                        //    dtData.Columns.Add(arrSourceColumn[i]);
                        //}

                        //strDestinationColumn = objUploadFileMaster.DestinationColumn;

                        //if (strDestinationColumn != "")
                        //    arrDestinationColumn = strDestinationColumn.Split('|');

                        //strTableName = objUploadFileMaster.TableName;

                        //if (arrSourceColumn != null && arrDestinationColumn != null && arrSourceColumn.Length == arrDestinationColumn.Length)
                        //{
                        //    using (SqlBulkCopy bulkCopy = new SqlBulkCopy(strMyConnection))
                        //    {
                        //        bulkCopy.DestinationTableName = strTableName;

                        //        for (int i = 0; i < arrSourceColumn.Length; i++)
                        //        {
                        //            bulkCopy.ColumnMappings.Add(arrSourceColumn[i], arrDestinationColumn[i]);
                        //        }
                        //        bulkCopy.WriteToServer(dtProductData);
                        //    }

                        //    DataSet dsIV = new DataSet();
                        //    dsIV = General.ExecuteDataSet(objUploadFileMaster.ExtraProcedure);
                        //}
                        #endregion
                        #endregion

                        #region Export MTM Report
                        Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                        System.Globalization.CultureInfo CurrentCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                        System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                        Microsoft.Office.Interop.Excel.Workbooks workbooks = xlApp.Workbooks;
                        Microsoft.Office.Interop.Excel.Workbook xlWorkbook = workbooks.Add(Microsoft.Office.Interop.Excel.XlWBATemplate.xlWBATWorksheet);

                        Sheets xlSheets = null;
                        Worksheet xlNewSheet = null;

                        xlApp.ErrorCheckingOptions.BackgroundChecking = false;
                        xlApp.Visible = false;
                        xlApp.DisplayAlerts = false;

                        string TargetFolder = System.Web.HttpContext.Current.Server.MapPath("~/OutputFiles/");

                        string TemplateFile = System.Web.HttpContext.Current.Server.MapPath("~/Templates/");
                        TemplateFile += @"\SPNoteMTMReportTemplate.xlsx";
                        
                        string strYear = Convert.ToString(objSPNoteMTM.ReportDate.Year);
                        string strMonth = Convert.ToString(objSPNoteMTM.ReportDate.Month.ToString());
                        string strDay = Convert.ToString(objSPNoteMTM.ReportDate.Day.ToString());

                        if (strMonth.Length == 1)
                            strMonth = "0" + strMonth;
                        if (strDay.Length == 1)
                            strDay = "0" + strDay;

                        string strFolderName = strYear + "-" + strMonth + "-" + strDay;

                        string ParentFolderLocation = TargetFolder + "\\" + strFolderName;

                        if (!Directory.Exists(ParentFolderLocation))
                        {
                            Directory.CreateDirectory(ParentFolderLocation);
                        }

                        string FileName = Path.GetFileNameWithoutExtension(TemplateFile) + DateTime.Now.ToString("dd_MMM_yyyy_hh_mm_ss") + "." + Path.GetExtension(TemplateFile);
                        string strTargetFilePath = "";
                        
                        strTargetFilePath = ParentFolderLocation + "\\" + FileName;

                        if (!System.IO.File.Exists(strTargetFilePath))
                        {
                            System.IO.File.Copy(TemplateFile, strTargetFilePath);
                        }

                        if (!System.IO.File.Exists(strTargetFilePath))
                        {
                            xlWorkbook.SaveAs(strTargetFilePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        }

                        xlApp.Quit();

                        xlApp = new Microsoft.Office.Interop.Excel.Application();
                        xlWorkbook = xlApp.Workbooks.Open(strTargetFilePath, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing,
                                       false, XlPlatform.xlWindows, Type.Missing,
                                       true, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        xlSheets = xlWorkbook.Sheets as Sheets;
                        DataSet dsResult = null;

                        if (objSPNoteMTM.IsFormula == true)
                            dsResult = General.ExecuteDataSet("SP_FETCH_SP_NOTE_MTM_REPORT_FORMULA_EXPORT", objSPNoteMTM.ReportDate);
                        else
                            dsResult = General.ExecuteDataSet("SP_FETCH_SP_NOTE_MTM_REPORT", objSPNoteMTM.ReportDate);

                        int intSQLRow = 0;
                        
                        if (dsResult != null && dsResult.Tables.Count == 6)
                        {
                            #region NewMethod
                            intSQLRow = 0;

                            xlNewSheet = xlSheets.get_Item("New method");
                            xlNewSheet.Name = "New method";

                            int heightNewMethod = dsResult.Tables[0].Rows.Count + 5;
                            int widthNewMethod = dsResult.Tables[0].Columns.Count;

                            object[,] retListNewMethod = new object[heightNewMethod, widthNewMethod];

                            retListNewMethod = new object[heightNewMethod, widthNewMethod];
                            for (int intCol = 1; intCol <= widthNewMethod; intCol++)
                            {
                                retListNewMethod[0, intCol - 1] = dsResult.Tables[0].Columns[intCol - 1].ColumnName;
                            }

                            //Write MS SQL Data
                            for (int intRow = 0; intRow < dsResult.Tables[0].Rows.Count; intRow++)
                            {
                                DataRow r = dsResult.Tables[0].Rows[intRow];
                                for (int intCol = 1; intCol <= widthNewMethod; intCol++)
                                    retListNewMethod[intRow + 1, intCol - 1] = r.ItemArray[intCol - 1];

                                intSQLRow = intRow + 1;
                            }

                            if (intSQLRow == 0)
                                intSQLRow = 1;

                            var startCellNewMethod = (Range)xlNewSheet.Cells[1, 1];
                            var endCellNewMethod = new object();
                            endCellNewMethod = (Range)xlNewSheet.Cells[intSQLRow + 1, dsResult.Tables[0].Columns.Count];

                            var writeRangeNewMethod = xlNewSheet.get_Range(startCellNewMethod, endCellNewMethod);
                            writeRangeNewMethod.set_Value(Type.Missing, retListNewMethod);

                            //xlWorkbook.Save();
                            //xlWorkbook.Close(Type.Missing, Type.Missing, Type.Missing);
                            //xlApp.Quit();
                            //if (xlNewSheet != null)
                            //{
                            //    Marshal.ReleaseComObject(xlNewSheet);
                            //    xlNewSheet = null;
                            //}

                            //if (xlSheets != null)
                            //{
                            //    Marshal.ReleaseComObject(xlSheets);
                            //    xlSheets = null;
                            //}

                            //if (xlWorkbook != null)
                            //{
                            //    Marshal.ReleaseComObject(xlWorkbook);
                            //    xlWorkbook = null;
                            //}

                            //if (xlApp != null)
                            //{
                            //    Marshal.ReleaseComObject(xlApp);
                            //    xlApp = null;
                            //}

                            //xlApp = new Microsoft.Office.Interop.Excel.Application();
                            //xlWorkbook = xlApp.Workbooks.Open(strTargetFilePath, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing,
                            //               false, XlPlatform.xlWindows, Type.Missing,
                            //               true, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                            //xlSheets = xlWorkbook.Sheets as Sheets;
                            #endregion

                            #region Summary
                            intSQLRow = 0;

                            xlNewSheet = xlSheets.get_Item("Summary");
                            xlNewSheet.Name = "Summary";

                            int heightSummary = dsResult.Tables[1].Rows.Count + 5;
                            int widthSummary = dsResult.Tables[1].Columns.Count;

                            object[,] retListSummary = new object[heightSummary, widthSummary];

                            retListSummary = new object[heightSummary, widthSummary];
                            for (int intCol = 1; intCol <= widthSummary; intCol++)
                            {
                                retListSummary[0, intCol - 1] = dsResult.Tables[1].Columns[intCol - 1].ColumnName;
                            }

                            //Write MS SQL Data
                            for (int intRow = 0; intRow < dsResult.Tables[1].Rows.Count; intRow++)
                            {
                                DataRow r = dsResult.Tables[1].Rows[intRow];
                                for (int intCol = 1; intCol <= widthSummary; intCol++)
                                    retListSummary[intRow + 1, intCol - 1] = r.ItemArray[intCol - 1];

                                intSQLRow = intRow + 1;
                            }

                            if (intSQLRow == 0)
                                intSQLRow = 1;

                            var startCellSummary = (Range)xlNewSheet.Cells[1, 1];
                            var endCellSummary = new object();
                            endCellSummary = (Range)xlNewSheet.Cells[intSQLRow + 1, dsResult.Tables[1].Columns.Count];

                            var writeRangeSummary = xlNewSheet.get_Range(startCellSummary, endCellSummary);
                            writeRangeSummary.set_Value(Type.Missing, retListSummary);

                            //xlWorkbook.Save();
                            //xlWorkbook.Close(Type.Missing, Type.Missing, Type.Missing);
                            //xlApp.Quit();
                            //if (xlNewSheet != null)
                            //{
                            //    Marshal.ReleaseComObject(xlNewSheet);
                            //    xlNewSheet = null;
                            //}

                            //if (xlSheets != null)
                            //{
                            //    Marshal.ReleaseComObject(xlSheets);
                            //    xlSheets = null;
                            //}

                            //if (xlWorkbook != null)
                            //{
                            //    Marshal.ReleaseComObject(xlWorkbook);
                            //    xlWorkbook = null;
                            //}

                            //if (xlApp != null)
                            //{
                            //    Marshal.ReleaseComObject(xlApp);
                            //    xlApp = null;
                            //}

                            //xlApp = new Microsoft.Office.Interop.Excel.Application();
                            //xlWorkbook = xlApp.Workbooks.Open(strTargetFilePath, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing,
                            //               false, XlPlatform.xlWindows, Type.Missing,
                            //               true, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                            //xlSheets = xlWorkbook.Sheets as Sheets;
                            #endregion

                            #region AveragingSheet
                            intSQLRow = 0;

                            xlNewSheet = xlSheets.get_Item("Averaging sheet");
                            xlNewSheet.Name = "Averaging sheet";

                            int heightAveragingSheet = dsResult.Tables[2].Rows.Count + 5;
                            int widthAveragingSheet = dsResult.Tables[2].Columns.Count;

                            object[,] retListAveragingSheet = new object[heightAveragingSheet, widthAveragingSheet];

                            retListAveragingSheet = new object[heightAveragingSheet, widthAveragingSheet];
                            for (int intCol = 1; intCol <= widthAveragingSheet; intCol++)
                            {
                                retListAveragingSheet[0, intCol - 1] = dsResult.Tables[2].Columns[intCol - 1].ColumnName;
                            }

                            //Write MS SQL Data
                            for (int intRow = 0; intRow < dsResult.Tables[2].Rows.Count; intRow++)
                            {
                                DataRow r = dsResult.Tables[2].Rows[intRow];
                                for (int intCol = 1; intCol <= widthAveragingSheet; intCol++)
                                    retListAveragingSheet[intRow + 1, intCol - 1] = r.ItemArray[intCol - 1];

                                intSQLRow = intRow + 1;
                            }

                            if (intSQLRow == 0)
                                intSQLRow = 1;

                            var startCellAveragingSheet = (Range)xlNewSheet.Cells[1, 1];
                            var endCellAveragingSheet = new object();
                            endCellAveragingSheet = (Range)xlNewSheet.Cells[intSQLRow + 1, dsResult.Tables[2].Columns.Count];

                            var writeRangeAveragingSheet = xlNewSheet.get_Range(startCellAveragingSheet, endCellAveragingSheet);
                            writeRangeAveragingSheet.set_Value(Type.Missing, retListAveragingSheet);

                            //xlWorkbook.Save();
                            //xlWorkbook.Close(Type.Missing, Type.Missing, Type.Missing);
                            //xlApp.Quit();
                            //if (xlNewSheet != null)
                            //{
                            //    Marshal.ReleaseComObject(xlNewSheet);
                            //    xlNewSheet = null;
                            //}

                            //if (xlSheets != null)
                            //{
                            //    Marshal.ReleaseComObject(xlSheets);
                            //    xlSheets = null;
                            //}

                            //if (xlWorkbook != null)
                            //{
                            //    Marshal.ReleaseComObject(xlWorkbook);
                            //    xlWorkbook = null;
                            //}

                            //if (xlApp != null)
                            //{
                            //    Marshal.ReleaseComObject(xlApp);
                            //    xlApp = null;
                            //}

                            //xlApp = new Microsoft.Office.Interop.Excel.Application();
                            //xlWorkbook = xlApp.Workbooks.Open(strTargetFilePath, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing,
                            //               false, XlPlatform.xlWindows, Type.Missing,
                            //               true, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                            //xlSheets = xlWorkbook.Sheets as Sheets;
                            #endregion

                            #region ActiveUnderlyingsheet
                            intSQLRow = 0;

                            xlNewSheet = xlSheets.get_Item("Active underlying sheet");
                            xlNewSheet.Name = "Active underlying sheet";

                            int heightActiveUnderlyingsheet = dsResult.Tables[3].Rows.Count + 5;
                            int widthActiveUnderlyingsheet = dsResult.Tables[3].Columns.Count;

                            object[,] retListActiveUnderlyingsheet = new object[heightActiveUnderlyingsheet, widthActiveUnderlyingsheet];

                            retListActiveUnderlyingsheet = new object[heightActiveUnderlyingsheet, widthActiveUnderlyingsheet];
                            for (int intCol = 1; intCol <= widthActiveUnderlyingsheet; intCol++)
                            {
                                retListActiveUnderlyingsheet[0, intCol - 1] = dsResult.Tables[3].Columns[intCol - 1].ColumnName;
                            }

                            //Write MS SQL Data
                            for (int intRow = 0; intRow < dsResult.Tables[3].Rows.Count; intRow++)
                            {
                                DataRow r = dsResult.Tables[3].Rows[intRow];
                                for (int intCol = 1; intCol <= widthActiveUnderlyingsheet; intCol++)
                                    retListActiveUnderlyingsheet[intRow + 1, intCol - 1] = r.ItemArray[intCol - 1];

                                intSQLRow = intRow + 1;
                            }

                            if (intSQLRow == 0)
                                intSQLRow = 1;

                            var startCellActiveUnderlyingsheet = (Range)xlNewSheet.Cells[1, 1];
                            var endCellActiveUnderlyingsheet = new object();
                            endCellActiveUnderlyingsheet = (Range)xlNewSheet.Cells[intSQLRow + 1, dsResult.Tables[3].Columns.Count];

                            var writeRangeActiveUnderlyingsheet = xlNewSheet.get_Range(startCellActiveUnderlyingsheet, endCellActiveUnderlyingsheet);
                            writeRangeActiveUnderlyingsheet.set_Value(Type.Missing, retListActiveUnderlyingsheet);

                            //xlWorkbook.Save();
                            //xlWorkbook.Close(Type.Missing, Type.Missing, Type.Missing);
                            //xlApp.Quit();
                            //if (xlNewSheet != null)
                            //{
                            //    Marshal.ReleaseComObject(xlNewSheet);
                            //    xlNewSheet = null;
                            //}

                            //if (xlSheets != null)
                            //{
                            //    Marshal.ReleaseComObject(xlSheets);
                            //    xlSheets = null;
                            //}

                            //if (xlWorkbook != null)
                            //{
                            //    Marshal.ReleaseComObject(xlWorkbook);
                            //    xlWorkbook = null;
                            //}

                            //if (xlApp != null)
                            //{
                            //    Marshal.ReleaseComObject(xlApp);
                            //    xlApp = null;
                            //}

                            //xlApp = new Microsoft.Office.Interop.Excel.Application();
                            //xlWorkbook = xlApp.Workbooks.Open(strTargetFilePath, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing,
                            //               false, XlPlatform.xlWindows, Type.Missing,
                            //               true, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                            //xlSheets = xlWorkbook.Sheets as Sheets;
                            #endregion

                            #region ExpiredUnderlyingsheet
                            intSQLRow = 0;

                            xlNewSheet = xlSheets.get_Item("Expired underlying sheet");
                            xlNewSheet.Name = "Expired underlying sheet";

                            int heightExpiredUnderlyingsheet = dsResult.Tables[4].Rows.Count + 5;
                            int widthExpiredUnderlyingsheet = dsResult.Tables[4].Columns.Count;

                            object[,] retListExpiredUnderlyingsheet = new object[heightExpiredUnderlyingsheet, widthExpiredUnderlyingsheet];

                            retListExpiredUnderlyingsheet = new object[heightExpiredUnderlyingsheet, widthExpiredUnderlyingsheet];
                            for (int intCol = 1; intCol <= widthExpiredUnderlyingsheet; intCol++)
                            {
                                retListExpiredUnderlyingsheet[0, intCol - 1] = dsResult.Tables[4].Columns[intCol - 1].ColumnName;
                            }

                            //Write MS SQL Data
                            for (int intRow = 0; intRow < dsResult.Tables[4].Rows.Count; intRow++)
                            {
                                DataRow r = dsResult.Tables[4].Rows[intRow];
                                for (int intCol = 1; intCol <= widthExpiredUnderlyingsheet; intCol++)
                                    retListExpiredUnderlyingsheet[intRow + 1, intCol - 1] = r.ItemArray[intCol - 1];

                                intSQLRow = intRow + 1;
                            }

                            if (intSQLRow == 0)
                                intSQLRow = 1;

                            var startCellExpiredUnderlyingsheet = (Range)xlNewSheet.Cells[1, 1];
                            var endCellExpiredUnderlyingsheet = new object();
                            endCellExpiredUnderlyingsheet = (Range)xlNewSheet.Cells[intSQLRow + 1, dsResult.Tables[4].Columns.Count];

                            var writeRangeExpiredUnderlyingsheet = xlNewSheet.get_Range(startCellExpiredUnderlyingsheet, endCellExpiredUnderlyingsheet);
                            writeRangeExpiredUnderlyingsheet.set_Value(Type.Missing, retListExpiredUnderlyingsheet);

                            //xlWorkbook.Save();
                            //xlWorkbook.Close(Type.Missing, Type.Missing, Type.Missing);
                            //xlApp.Quit();
                            //if (xlNewSheet != null)
                            //{
                            //    Marshal.ReleaseComObject(xlNewSheet);
                            //    xlNewSheet = null;
                            //}

                            //if (xlSheets != null)
                            //{
                            //    Marshal.ReleaseComObject(xlSheets);
                            //    xlSheets = null;
                            //}

                            //if (xlWorkbook != null)
                            //{
                            //    Marshal.ReleaseComObject(xlWorkbook);
                            //    xlWorkbook = null;
                            //}

                            //if (xlApp != null)
                            //{
                            //    Marshal.ReleaseComObject(xlApp);
                            //    xlApp = null;
                            //}

                            //xlApp = new Microsoft.Office.Interop.Excel.Application();
                            //xlWorkbook = xlApp.Workbooks.Open(strTargetFilePath, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing,
                            //               false, XlPlatform.xlWindows, Type.Missing,
                            //               true, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                            //xlSheets = xlWorkbook.Sheets as Sheets;
                            #endregion

                            #region DebenturesBuyback
                            intSQLRow = 0;

                            xlNewSheet = xlSheets.get_Item("Debentures Buyback");
                            xlNewSheet.Name = "Debentures Buyback";

                            int heightDebenturesBuyback = dsResult.Tables[5].Rows.Count + 5;
                            int widthDebenturesBuyback = dsResult.Tables[5].Columns.Count;

                            object[,] retListDebenturesBuyback = new object[heightDebenturesBuyback, widthDebenturesBuyback];

                            retListDebenturesBuyback = new object[heightDebenturesBuyback, widthDebenturesBuyback];
                            for (int intCol = 1; intCol <= widthDebenturesBuyback; intCol++)
                            {
                                retListDebenturesBuyback[0, intCol - 1] = dsResult.Tables[5].Columns[intCol - 1].ColumnName;
                            }

                            //Write MS SQL Data
                            for (int intRow = 0; intRow < dsResult.Tables[5].Rows.Count; intRow++)
                            {
                                DataRow r = dsResult.Tables[5].Rows[intRow];
                                for (int intCol = 1; intCol <= widthDebenturesBuyback; intCol++)
                                    retListDebenturesBuyback[intRow + 1, intCol - 1] = r.ItemArray[intCol - 1];

                                intSQLRow = intRow + 1;
                            }

                            if (intSQLRow == 0)
                                intSQLRow = 1;

                            var startCellDebenturesBuyback = (Range)xlNewSheet.Cells[1, 1];
                            var endCellDebenturesBuyback = new object();
                            endCellDebenturesBuyback = (Range)xlNewSheet.Cells[intSQLRow + 1, dsResult.Tables[5].Columns.Count];

                            var writeRangeDebenturesBuyback = xlNewSheet.get_Range(startCellDebenturesBuyback, endCellDebenturesBuyback);
                            writeRangeDebenturesBuyback.set_Value(Type.Missing, retListDebenturesBuyback);
                            #endregion

                            xlWorkbook.Save();
                            xlWorkbook.Close(Type.Missing, Type.Missing, Type.Missing);
                            xlApp.Quit();

                            if (xlNewSheet != null)
                            {
                                Marshal.ReleaseComObject(xlNewSheet);
                                xlNewSheet = null;
                            }

                            if (xlSheets != null)
                            {
                                Marshal.ReleaseComObject(xlSheets);
                                xlSheets = null;
                            }

                            if (xlWorkbook != null)
                            {
                                Marshal.ReleaseComObject(xlWorkbook);
                                xlWorkbook = null;
                            }

                            if (xlApp != null)
                            {
                                Marshal.ReleaseComObject(xlApp);
                                xlApp = null;
                            }

                            if (System.IO.File.Exists(strTargetFilePath))
                            {
                                FileInfo objFileInfo = new FileInfo(strTargetFilePath);

                                Response.Clear();
                                Response.ClearHeaders();
                                Response.ClearContent();
                                Response.AddHeader("content-disposition", "attachment; filename=" + Path.GetFileName(strTargetFilePath));
                                Response.AddHeader("Content-Type", "application/Excel");
                                Response.ContentType = "application/vnd.xls";
                                Response.AddHeader("Content-Length", objFileInfo.Length.ToString());
                                Response.WriteFile(objFileInfo.FullName);
                                Response.End();
                            }
                        }
                        else
                        {
                            if (dsResult != null && dsResult.Tables.Count == 1)
                            {
                                ViewBag.Message = "Price is not available for " + objSPNoteMTM.ReportDate.ToString("dd-M-yyyy") + " ";
                                Session["UnderlyingName"] = dsResult.Tables[0];
                            }
                        }
                        #endregion

                        return View();
                    }

                    return View();
                }
                else
                {
                    return RedirectToAction("Login", "Login");
                }
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "SPNoteMTMController", "MTMReport Post", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        private void FetchUploadFileMasterList()
        {
            if (Session["UploadFileMasterList"] == null)
            {
                ObjectResult<UploadFileMasterResult> objUploadFileMasterResult = objSP_PRICINGEntities.SP_FETCH_UPLOAD_FILE_MASTER_DETAILS();
                List<UploadFileMasterResult> UploadFileMasterResultList = objUploadFileMasterResult.ToList();

                List<UploadFileMaster> UploadFileMasterList = new List<UploadFileMaster>();

                if (UploadFileMasterResultList != null && UploadFileMasterResultList.Count > 0)
                {
                    foreach (UploadFileMasterResult oUploadFileMasterResult in UploadFileMasterResultList)
                    {
                        UploadFileMaster objUploadFileMaster = new UploadFileMaster();
                        General.ReflectSingleData(objUploadFileMaster, oUploadFileMasterResult);
                        UploadFileMasterList.Add(objUploadFileMaster);
                    }
                }


                Session["UploadFileMasterList"] = UploadFileMasterList;
            }
        }

        public JsonResult FetchUnderlyingName()
        {
            try
            {
                List<SPNoteMTM> SPNoteMTMList = new List<SPNoteMTM>();

                DataSet dsResult = new DataSet();
                System.Data.DataTable a = (System.Data.DataTable)Session["UnderlyingName"];

                if (a.Rows.Count > 0)
                {
                    foreach (DataRow dr in a.Rows)
                    {
                        SPNoteMTM obj = new SPNoteMTM();
                        obj.Underlying = Convert.ToString(dr["Underlying"]);

                        SPNoteMTMList.Add(obj);
                    }
                }

                var SPNoteMTMListData = SPNoteMTMList.ToList();
                return Json(SPNoteMTMListData, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "SPNoteMTMController", "FetchUnderlyingName", objUserMaster.UserID);
                return Json("");
            }
        }


        #region MTM Upload
        [HttpGet]
        public ActionResult MTMUpload(string Msg)
        {
            LoginController objLoginController = new LoginController();
            MTMUpload objMTMUpload = new MTMUpload();

            try
            {
                if (ValidateSession())
                {

                    #region Menu Access By on Role

                    Int32 intResult = 0;
                    // bool PPorNonPP = false;

                    UserMaster objUserMaster = new UserMaster();
                    objUserMaster = (UserMaster)Session["LoggedInUser"];

                    var Result = objSP_PRICINGEntities.VALIDATE_MENU_BY_ROLE(objUserMaster.RoleID, "MTMU");
                    intResult = Convert.ToInt32(Result.SingleOrDefault());

                    if (intResult == 0)
                        return RedirectToAction("UserNotAuthorize", "Login");

                    #endregion

                    if (Msg == "1")
                    {
                        ViewBag.Message = "Imported successfully";
                    }
                    return View(objMTMUpload);
                }
                else
                {
                    return RedirectToAction("Login", "Login");
                }
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "SPNoteMTMController", "MTMUpload Get", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        [HttpPost]
        public ActionResult MTMUpload(MTMUpload objMTMUpload, string Command, FormCollection collection, HttpPostedFileBase file)
        {
            LoginController objLoginController = new LoginController();
            List<UploadFileMaster> UploadFileMasterList = (List<UploadFileMaster>)Session["UploadFileMasterList"];

            try
            {
                if (ValidateSession())
                {
                    UserMaster objUserMaster = (UserMaster)Session["LoggedInUser"];
                    if (Command == "Upload")
                    {
                        bool blnUploadStatus = false;
                        bool blnUploadDataStatus = true;

                        if (file != null && file.ContentLength > 0)
                        {
                            UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "MTMU"; });
                            string strFilePath = System.Web.HttpContext.Current.Server.MapPath("~/Uploads/");
                            string strFileName = Path.GetFileNameWithoutExtension(file.FileName) + DateTime.Now.ToString("dd_mmm_yyyy_hh_mm");
                            string strExtension = Path.GetExtension(file.FileName);

                            strFilePath += strFileName + strExtension;
                            file.SaveAs(strFilePath);

                            FileInfo newFile = new FileInfo(strFilePath);

                            #region Source and Destination Column
                            string strSourceColumn = objUploadFileMaster.SourceColumn;
                            string[] arrSourceColumn = null;
                            if (strSourceColumn != "")
                                arrSourceColumn = strSourceColumn.Split('|');

                            System.Data.DataTable dtData = new System.Data.DataTable();

                            for (int i = 0; i < arrSourceColumn.Length; i++)
                            {
                                dtData.Columns.Add(arrSourceColumn[i]);
                            }

                            string strDestinationColumn = objUploadFileMaster.DestinationColumn;
                            string[] arrDestinationColumn = null;

                            if (strDestinationColumn != "")
                                arrDestinationColumn = strDestinationColumn.Split('|');

                            string strTableName = objUploadFileMaster.TableName;
                            #endregion

                            using (ExcelPackage xlPackage = new ExcelPackage(newFile))
                            {
                                ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets[objUploadFileMaster.SheetName];

                                DataRow drNew;

                                for (int iRow = 2; iRow < 1000; iRow++)
                                {
                                    if (Convert.ToString(worksheet.Cell(iRow, 1).Value) != "")
                                    {
                                        drNew = dtData.NewRow();

                                        drNew["PRODUCT_CODE"] = worksheet.Cell(iRow, 1).Value;
                                        drNew["UNDERLYING_TYPE"] = worksheet.Cell(iRow, 2).Value;
                                        drNew["PRODUCT_TYPE"] = worksheet.Cell(iRow, 3).Value;
                                        drNew["UNDERLYING"] = worksheet.Cell(iRow, 4).Value;
                                        drNew["AUM"] = worksheet.Cell(iRow, 5).Value;
                                        drNew["ACTUAL_DEPLOYMENT_RATE"] = worksheet.Cell(iRow, 6).Value;
                                        drNew["FIXED_RETURN"] = worksheet.Cell(iRow, 7).Value;
                                        drNew["START_DATE"] = DateTime.FromOADate(Convert.ToDouble(worksheet.Cell(iRow, 8).Value)).ToString("yyyy-MM-dd");
                                        drNew["REDEMPTION_DATE"] = DateTime.FromOADate(Convert.ToDouble(worksheet.Cell(iRow, 9).Value)).ToString("yyyy-MM-dd");
                                        drNew["EARLY_REDEMPTION_DATE"] = worksheet.Cell(iRow, 10).Value;
                                        drNew["BOND_PRICE"] = worksheet.Cell(iRow, 11).Value;
                                        drNew["OPTION_PRICE"] = worksheet.Cell(iRow, 12).Value;
                                        drNew["SP_VALUE"] = worksheet.Cell(iRow, 13).Value;
                                        drNew["SP_NOTE_MTM"] = worksheet.Cell(iRow, 14).Value;
                                        drNew["DEBENTURES_BUYBACK"] = worksheet.Cell(iRow, 15).Value;
                                        drNew["BOND_VALUE_BASE_ON_100"] = worksheet.Cell(iRow, 16).Value;
                                        drNew["OPTION_VALUE_BASE_ON_100"] = worksheet.Cell(iRow, 17).Value;
                                        drNew["SP_VALUE_BASE_ON_100"] = worksheet.Cell(iRow, 18).Value;
                                        drNew["SP_FINAL_FIXING_DATE"] = DateTime.FromOADate(Convert.ToDouble(worksheet.Cell(iRow, 19).Value)).ToString("yyyy-MM-dd");
                                        drNew["ACTUAL_DEPLOYMENT_RATE1"] = worksheet.Cell(iRow, 20).Value;
                                        drNew["DISCOUNT_RATE"] = worksheet.Cell(iRow, 21).Value;
                                        drNew["BOND_DISCOUNTING"] = worksheet.Cell(iRow, 22).Value;
                                        drNew["OPTION_DISCOUNTED_TILL"] = DateTime.FromOADate(Convert.ToDouble(worksheet.Cell(iRow, 23).Value)).ToString("yyyy-MM-dd");
                                        drNew["DISCOUNTED_OPTION"] = worksheet.Cell(iRow, 24).Value;
                                        drNew["REPORT_DATE"] = DateTime.FromOADate(Convert.ToDouble(worksheet.Cell(iRow, 25).Value)).ToString("yyyy-MM-dd");

                                        dtData.Rows.Add(drNew);
                                    }
                                    else
                                        break;
                                }
                            }

                            string strMyConnection = Convert.ToString(System.Configuration.ConfigurationManager.ConnectionStrings["SP_PRICINGConnectionString"]);

                            if (arrSourceColumn != null && arrDestinationColumn != null && arrSourceColumn.Length == arrDestinationColumn.Length)
                            {
                                using (SqlBulkCopy bulkCopy = new SqlBulkCopy(strMyConnection))
                                {
                                    bulkCopy.DestinationTableName = strTableName;

                                    for (int i = 0; i < arrSourceColumn.Length; i++)
                                    {
                                        bulkCopy.ColumnMappings.Add(arrSourceColumn[i], arrDestinationColumn[i]);
                                    }
                                    bulkCopy.WriteToServer(dtData);
                                }
                                blnUploadStatus = true;

                                DataSet dsIV = new DataSet();
                                dsIV = General.ExecuteDataSet(objUploadFileMaster.ExtraProcedure, objUserMaster.UserID);
                            }
                            else
                                blnUploadStatus = false;

                            if (blnUploadStatus)
                            {
                                ManageUploadFileInfo(0, file.FileName, strFilePath, blnUploadStatus, blnUploadDataStatus);
                                ViewBag.Message = "Imported successfully";
                            }
                            return RedirectToAction("MTMUpload", new { Msg = "1" });
                            //return RedirectToAction("MTMUpload");
                        }
                    }

                    #region MTMDownload
                    else if (Command == "Download")
                    {
                        UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "MTMU"; });
                        string strFilePath = System.Web.HttpContext.Current.Server.MapPath(objUploadFileMaster.TemplateFileName);

                        if (System.IO.File.Exists(strFilePath))
                        {
                            FileInfo fileinfo = new FileInfo(strFilePath);

                            Response.Clear();
                            Response.ClearHeaders();
                            Response.ClearContent();
                            Response.AddHeader("content-disposition", "attachment; filename=" + Path.GetFileName(strFilePath));
                            Response.AddHeader("Content-Type", "application/Excel");
                            Response.ContentType = "application/vnd.xls";
                            Response.AddHeader("Content-Length", fileinfo.Length.ToString());
                            Response.WriteFile(fileinfo.FullName);
                            Response.End();


                            return RedirectToAction("MTMUpload");
                        }
                        else
                            return View();
                    }
                    #endregion

                    return View();
                }
                else
                {
                    return RedirectToAction("Login", "Login");
                }
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "SPNoteMTMController", "MTMUpload Post", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        #endregion


        public void LogError(string strErrorDescription, string strStackTrace, string strClassName, string strMethodName, Int32 intUserId)
        {
            SP_PRICINGEntities objSP_PRICINGEntities = new SP_PRICINGEntities();
            var Count = objSP_PRICINGEntities.SP_ERROR_LOG(strErrorDescription, strStackTrace, strClassName, strMethodName, intUserId);
        }

        public bool ValidateSession()
        {
            LoginController objLoginController = new LoginController();

            try
            {
                if (Session["LoggedInUser"] != null)
                    return true;
                else
                    return false;
            }
            catch (Exception ex)
            {
                objLoginController.LogError(ex.Message, ex.StackTrace, "SPNoteMTMController", "ValidateSession", -1);
                return false;
            }
        }

        public void ManageUploadFileInfo(Int32 intUnderlyingID, string strOriginalFileName, string strFilePath, bool blnUploadStatus, bool blnUploadDataStatus)
        {
            Int32 intUploadType = 0;
            UserMaster objUserMaster = new UserMaster();
            objUserMaster = (UserMaster)Session["LoggedInUser"];


            Int32 intResult = 0;
            var Count = objSP_PRICINGEntities.SP_MANAGE_UPLOAD_FILE_INFO(intUnderlyingID, intUploadType, strOriginalFileName, Path.GetFileName(strFilePath), strFilePath, blnUploadStatus, blnUploadDataStatus, objUserMaster.UserID);
            intResult = Count.SingleOrDefault().Value;
        }
    }
}