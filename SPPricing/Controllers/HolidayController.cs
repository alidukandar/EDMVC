using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using SPPricing.Models;
using System.Data;
using System.IO;
using System.Data.Objects;
using OfficeOpenXml;
using System.Data.SqlClient;

namespace SPPricing.Controllers
{
    public class HolidayController : Controller
    {

        SP_PRICINGEntities objSP_PRICINGEntities = new SP_PRICINGEntities();
        //
        // GET: /Holiday/

        public ActionResult Index()
        {
            return View();
        }

        [HttpGet]
        public ActionResult HolidayMaster(int? holidayID)
        {
            HolidayMaster hMaster = new HolidayMaster();
            if (ValidateSession())
            {
            }
            if (holidayID != null)
            {
                List<HolidayMaster> hm = FetchHolidayMasterData();
                var hMasterResult = hm.FirstOrDefault(h => h.ID.Equals(holidayID));
                hMaster = hMasterResult;
            }
            return View(hMaster);
        }


        [HttpPost]
        public Int32 DeleteHolidayMaster(string HolidayID, string HolidayName, string HolidayDate)
        {
            HolidayMaster hMaster = new HolidayMaster();
            if (ValidateSession())
            {
            }
            if (HolidayID != null)
            {
                Int32 IsActive = 0;
                ManageHoliday(HolidayID, HolidayName, HolidayDate, IsActive);
            }
            return 1;
        }


        [HttpPost]
        public ActionResult HolidayMaster(string Command, HttpPostedFileBase file)
        {
            if (ValidateSession())
            {
            }
            FetchUploadFileMasterList();
            List<UploadFileMaster> UploadFileMasterList = (List<UploadFileMaster>)Session["UploadFileMasterList"];
            HolidayMaster hMaster = new HolidayMaster();
            if (Command == "PDDownload")
            {

                UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(
                                                        delegate(UploadFileMaster oUploadFileMaster)
                                                        {
                                                            return oUploadFileMaster.UploadTypeCode == "HM";
                                                        });
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
                }
            }
            if (Command == "PDUpload")
            {
                if (file != null && file.ContentLength > 0)
                {
                    UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "HM"; });
                    string strFilePath = System.Web.HttpContext.Current.Server.MapPath("~/Uploads/");
                    string strFileName = Path.GetFileNameWithoutExtension(file.FileName) + DateTime.Now.ToString("dd_mmm_yyyy_hh_mm");
                    string strExtension = Path.GetExtension(file.FileName);

                    strFilePath += strFileName + strExtension;
                    file.SaveAs(strFilePath);

                    FileInfo newFile = new FileInfo(strFilePath);

                    string strSourceColumn = objUploadFileMaster.SourceColumn;
                    string[] arrSourceColumn = null;
                    if (strSourceColumn != "")
                        arrSourceColumn = strSourceColumn.Split('|');

                    DataTable dtData = new DataTable();

                    for (int i = 0; i < arrSourceColumn.Length; i++)
                    {
                        dtData.Columns.Add(arrSourceColumn[i]);
                    }

                    string strDestinationColumn = objUploadFileMaster.DestinationColumn;
                    string[] arrDestinationColumn = null;

                    if (strDestinationColumn != "")
                        arrDestinationColumn = strDestinationColumn.Split('|');

                    string strTableName = objUploadFileMaster.TableName;

                    using (ExcelPackage xlPackage = new ExcelPackage(newFile))
                    {
                        ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets[objUploadFileMaster.SheetName];

                        string strVersion = worksheet.Cell(1, 2).Value.Trim();
                        if (strVersion == "")
                            strVersion = "0";

                        DataRow drNew;

                        for (int iRow = 2; iRow < 36; iRow++)
                        {
                            if (Convert.ToString(worksheet.Cell(iRow, 1).Value) != "")
                            {
                                drNew = dtData.NewRow();

                                var Reason = worksheet.Cell(iRow, 1).Value;
                                var Holiday_Date = DateTime.FromOADate(double.Parse(worksheet.Cell(iRow, 2).Value));

                                drNew["REASON"] = Reason;//worksheet.Cell(iRow, 1).Value;
                                drNew["HOLIDAY_DATE"] = Convert.ToDateTime(Holiday_Date);
                                drNew["VERSION"] = 1;
                                drNew["ISACTIVE"] = 1;
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
                    }
                    else
                    {
                        blnUploadStatus = false;
                    }
                    if (blnUploadStatus)
                    {
                        ViewBag.Message = "Imported successfully";
                    }

                    return View(hMaster);
                }
            }
            return View(hMaster);
        }


        public JsonResult FetchHolidayMaster()
        {
            try
            {
                return Json(FetchHolidayMasterData(), JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "FetchHolidayMaster", objUserMaster.UserID);
                return Json("");
            }
        }

        public List<HolidayMaster> FetchHolidayMasterData()
        {
            try
            {
                List<HolidayMaster> HolidayMasterList = new List<HolidayMaster>();
                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("FETCH_HOLIDAY_DETAILS");

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in dsResult.Tables[0].Rows)
                    {
                        HolidayMaster obj = new HolidayMaster();
                        obj.ID = Convert.ToInt32(dr["ID"]);
                        obj.Version = Convert.ToInt32(dr["VERSION"]);
                        obj.HolidayDate = Convert.ToDateTime(dr["HOLIDAY_DATE"]);
                        obj.Reason = Convert.ToString(dr["REASON"]);
                        HolidayMasterList.Add(obj);
                    }
                }

                var HolidayMasterListData = HolidayMasterList.ToList();
                return HolidayMasterListData;
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "FetchHolidayMaster", objUserMaster.UserID);
                return null;
            }
        }


        public Int32 ManageHoliday(string HolidayID, string HolidayName, string HolidayDate, Int32 IsActive)
        {
            UserMaster objUserMaster = new UserMaster();
            objUserMaster = (UserMaster)Session["LoggedInUser"];
            Int32 intResult = 0;
            DateTime dtHolidayDate = Convert.ToDateTime("1900-01-01");

            if (IsActive != 0 && HolidayDate != "")
                HolidayDate = HolidayDate.Substring(6, 4) + '-' + HolidayDate.Substring(0, 2) + '-' + HolidayDate.Substring(3, 2);
            try
            {
                if (ValidateSession())
                {
                    var Result = objSP_PRICINGEntities.SP_MANAGE_HOLIDAY_MASTER_LIST(Convert.ToInt32(HolidayID), HolidayName, Convert.ToDateTime(HolidayDate), IsActive);
                    intResult = Convert.ToInt32(Result.SingleOrDefault());
                }
                return intResult;
            }
            catch (Exception ex)
            {
                LogError(ex.Message, ex.StackTrace, "HolidayController", "ManageHoliday", objUserMaster.UserID);
                return 0;
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
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "HolidayController", "ManageHoliday", objUserMaster.UserID);
                return false;
            }
        }


        public void LogError(string strErrorDescription, string strStackTrace, string strClassName, string strMethodName, Int32 intUserId)
        {
            SP_PRICINGEntities objSP_PRICINGEntities = new SP_PRICINGEntities();
            var Count = objSP_PRICINGEntities.SP_ERROR_LOG(strErrorDescription, strStackTrace, strClassName, strMethodName, intUserId);
        }


        public void ManageUploadFileInfo(Int32 intUnderlyingID, string strOriginalFileName, string strFilePath, bool blnUploadStatus, bool blnUploadDataStatus)
        {
            Int32 intUploadType = 0;
            Int32 intResult = 0;
            UserMaster objUserMaster = new UserMaster();
            objUserMaster = (UserMaster)Session["LoggedInUser"];
            var Count = objSP_PRICINGEntities.SP_MANAGE_UPLOAD_FILE_INFO(intUnderlyingID, intUploadType, strOriginalFileName, Path.GetFileName(strFilePath), strFilePath, blnUploadStatus, blnUploadDataStatus, objUserMaster.UserID);
            intResult = Count.SingleOrDefault().Value;
        }

        public bool blnUploadStatus { get; set; }

        public bool blnUploadDataStatus { get; set; }
    }
}
