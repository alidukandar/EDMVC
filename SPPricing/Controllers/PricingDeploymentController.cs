using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using SPPricing.Models;
using System.Data;
using System.Data.Objects;
using System.IO;
using OfficeOpenXml;
using System.Data.SqlClient;

namespace SPPricing.Controllers
{
    public class PricingDeploymentController : Controller
    {
        //
        // GET: /PricingDeployment/

        SP_PRICINGEntities objSP_PRICINGEntities = new SP_PRICINGEntities();

        public ActionResult Index()
        {
            return View();
        }

        [HttpGet]
        public ActionResult PricingDeployment(int? ID)
        {

            LoginController objLoginController = new LoginController();
            PricingDeployment objPricingDeployment = new PricingDeployment();

            try
            {
                if (ValidateSession())
                {
                    #region Menu Access By on Role

                    Int32 intResult = 0;
                    UserMaster objUserMaster = new UserMaster();
                    objUserMaster = (UserMaster)Session["LoggedInUser"];

                    var Result = objSP_PRICINGEntities.VALIDATE_MENU_BY_ROLE(objUserMaster.RoleID, "UCPD");
                    intResult = Convert.ToInt32(Result.SingleOrDefault());

                    if (intResult == 0)
                        return RedirectToAction("UserNotAuthorize", "Login");

                    #endregion

                    FetchUploadFileMasterList();

                    return View(objPricingDeployment);
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
                LogError(ex.Message, ex.StackTrace, "PricingDeploymentController", "PricingDeployment Get", objUserMaster.UserID);
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

        [HttpPost]
        public ActionResult PricingDeployment(PricingDeployment objPricingDeployment, string Command, FormCollection collection, HttpPostedFileBase file, HttpPostedFileBase file1)
        {
            LoginController objLoginController = new LoginController();
            List<UploadFileMaster> UploadFileMasterList = (List<UploadFileMaster>)Session["UploadFileMasterList"];

            string strEntityID = collection["EntityID"];
            string strIsSecuredID = collection["IsSecuredID"];

            Int32 intEntityID = Convert.ToInt32(strEntityID.Split(',')[0]);
            Int32 intIsSecuredID = Convert.ToInt32(strIsSecuredID.Split(',')[0]);

            try
            {
                if (ValidateSession())
                {
                    bool blnUploadStatus = false;
                    bool blnUploadDataStatus = true;

                    #region Price Deployment
                    if (Command == "PDUpload")
                    {
                        if (file != null && file.ContentLength > 0)
                        {

                            UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "PDR"; });
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

                                for (int iRow = 4; iRow < 36; iRow++)
                                {
                                    if (Convert.ToString(worksheet.Cell(iRow, 1).Value) != "")
                                    {
                                        drNew = dtData.NewRow();

                                        var Days = worksheet.Cell(iRow, 1).Value.Split('-');
                                        var MinDays = Days[0].Trim();
                                        var MaxDays = Days[1].Trim();

                                        drNew["VERSION"] = strVersion;
                                        drNew["ENTITY_ID"] = intEntityID;
                                        drNew["IS_SECURED_ID"] = intIsSecuredID;
                                        drNew["MIN_DAYS"] = MinDays;//worksheet.Cell(iRow, 1).Value;
                                        drNew["MAX_DAYS"] = MaxDays;
                                        drNew["DEPLOYMENT_RATE"] = worksheet.Cell(iRow, 2).Value;
                                        drNew["CREATED_DATE"] = DateTime.Now;
                                        drNew["CREATED_BY"] = 1;

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
                                dsIV = General.ExecuteDataSet(objUploadFileMaster.ExtraProcedure);
                            }
                            else
                            {
                                blnUploadStatus = false;
                            }

                            if (blnUploadStatus)
                            {
                                ManageUploadFileInfo(0, file.FileName, strFilePath, blnUploadStatus, blnUploadDataStatus);

                                ViewBag.Message = "Imported successfully";
                            }

                            return View(objPricingDeployment);
                        }
                    }
                    #endregion

                    #region Actual Deployment
                    else if (Command == "ADUpload")
                    {
                        if (file1 != null && file1.ContentLength > 0)
                        {
                            UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "ADR"; });

                            string strFilePath = System.Web.HttpContext.Current.Server.MapPath("~/Uploads/");
                            string strFileName = Path.GetFileNameWithoutExtension(file1.FileName) + DateTime.Now.ToString("dd_mmm_yyyy_hh_mm");
                            string strExtension = Path.GetExtension(file1.FileName);

                            strFilePath += strFileName + strExtension;
                            file1.SaveAs(strFilePath);

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

                                for (int iRow = 4; iRow < 100; iRow++)
                                {
                                    if (Convert.ToString(worksheet.Cell(iRow, 1).Value) != "")
                                    {
                                        drNew = dtData.NewRow();

                                        var Days = worksheet.Cell(iRow, 1).Value.Split(' ');
                                        var MinDays = Days[0].Trim();
                                        var MaxDays = Days[2].Trim();

                                        drNew["VERSION"] = strVersion;
                                        drNew["MIN_DAYS"] = MinDays;//worksheet.Cell(iRow, 1).Value;
                                        drNew["MAX_DAYS"] = MaxDays;
                                        drNew["DEPLOYMENT_RATE"] = worksheet.Cell(iRow, 2).Value;
                                        drNew["CREATED_DATE"] = DateTime.Now;
                                        drNew["CREATED_BY"] = 1;

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
                                dsIV = General.ExecuteDataSet(objUploadFileMaster.ExtraProcedure);
                            }
                            else
                            {
                                blnUploadStatus = false;
                            }

                            if (blnUploadStatus)
                            {
                                ManageUploadFileInfo(0, file1.FileName, strFilePath, blnUploadStatus, blnUploadDataStatus);

                                ViewBag.Message = "Imported successfully";
                            }

                            return View(objPricingDeployment);
                        }
                    }
                    #endregion

                    #region Pricing Deployment Download
                    else if (Command == "PDDownload")
                    {
                        UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "PDR"; });
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

                        return View(objPricingDeployment);
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
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "PricingDeployment Post", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        public JsonResult FetchPricingDeployment(string Entity, string IsSecured)
        {
            try
            {
                List<PricingDeployment> PricingDeploymentList = new List<PricingDeployment>();

                if (Entity == "")
                    Entity = "-1";

                if (IsSecured == "")
                    IsSecured = "-1";

                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("FETCH_PRICING_DEPLOYMENT", Convert.ToInt32(Entity), Convert.ToInt32(IsSecured));

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in dsResult.Tables[0].Rows)
                    {
                        PricingDeployment obj = new PricingDeployment();

                        obj.Version = Convert.ToInt32(dr["VERSION"]);
                        obj.EntityName = Convert.ToString(dr["ENTITY"]);
                        obj.IsSecuredName = Convert.ToString(dr["IS_SECURED"]);
                        obj.MinDays = Convert.ToInt32(dr["MIN_DAYS"]);
                        obj.MaxDays = Convert.ToInt32(dr["MAX_DAYS"]);
                        obj.DeploymentRate = Convert.ToDouble(dr["DEPLOYMENT_RATE"]);
                        obj.CreatedDate = Convert.ToDateTime(dr["CREATED_DATE"]);

                        PricingDeploymentList.Add(obj);
                    }
                }

                var PricingDeploymentListData = PricingDeploymentList.ToList();
                return Json(PricingDeploymentListData, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "FetchPricingDeployment", objUserMaster.UserID);
                return Json("");
            }
        }

        public JsonResult FetchActualDeployment(string Entity, string IsSecured)
        {
            try
            {
                List<ActualDeployment> ActualDeploymentList = new List<ActualDeployment>();

                if (Entity == "")
                    Entity = "-1";

                if (IsSecured == "")
                    IsSecured = "-1";

                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("FETCH_ACTUAL_DEPLOYMENT", Convert.ToInt32(Entity), Convert.ToInt32(IsSecured));

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in dsResult.Tables[0].Rows)
                    {
                        ActualDeployment obj = new ActualDeployment();

                        obj.Version = Convert.ToInt32(dr["Version"]);
                        obj.MinDays = Convert.ToInt32(dr["MinDays"]);
                        obj.MaxDays = Convert.ToInt32(dr["MaxDays"]);
                        obj.DeploymentRate = Convert.ToDouble(dr["DeploymentRate"]);
                        obj.CreatedDate = Convert.ToDateTime(dr["CreatedDate"]);
                        obj.EntityName = Convert.ToString(dr["EntityName"]);
                        obj.IsSecured = Convert.ToString(dr["IsSecured"]);

                        ActualDeploymentList.Add(obj);
                    }
                }

                var ActualDeploymentListData = ActualDeploymentList.ToList();
                return Json(ActualDeploymentListData, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "FetchActualDeployment", objUserMaster.UserID);
                return Json("");
            }
        }

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
                objLoginController.LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "ValidateSession", -1);
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

        public Underlying FetchDefaultDetails()
        {
            Underlying objUnderlying = new Underlying();
            try
            {

                #region Standard List
                List<LookupMaster> StandardList = new List<LookupMaster>();

                ObjectResult<LookupResult> objLookupResult;
                List<LookupResult> LookupResultList;

                objLookupResult = objSP_PRICINGEntities.SP_FETCH_LOOKUP_VALUES("US", false);
                LookupResultList = objLookupResult.ToList();

                if (LookupResultList != null && LookupResultList.Count > 0)
                {
                    foreach (var LookupResult in LookupResultList)
                    {
                        LookupMaster objLookupMaster = new LookupMaster();
                        General.ReflectSingleData(objLookupMaster, LookupResult);
                        StandardList.Add(objLookupMaster);
                    }
                }
                objUnderlying.StandardList = StandardList;
                #endregion

                #region IV RF List
                List<LookupMaster> IVRFList = new List<LookupMaster>();

                objLookupResult = objSP_PRICINGEntities.SP_FETCH_LOOKUP_VALUES("IRC", false);
                LookupResultList = objLookupResult.ToList();

                if (LookupResultList != null && LookupResultList.Count > 0)
                {
                    foreach (var LookupResult in LookupResultList)
                    {
                        LookupMaster objLookupMaster = new LookupMaster();
                        General.ReflectSingleData(objLookupMaster, LookupResult);

                        IVRFList.Add(objLookupMaster);
                    }
                }

                objUnderlying.IVRFCategoryList = IVRFList;
                #endregion

                #region Call Spread Uploads
                List<LookupMaster> CSUList = new List<LookupMaster>();

                objLookupResult = objSP_PRICINGEntities.SP_FETCH_LOOKUP_VALUES("CSU", false);
                LookupResultList = objLookupResult.ToList();

                if (LookupResultList != null && LookupResultList.Count > 0)
                {
                    foreach (var LookupResult in LookupResultList)
                    {
                        LookupMaster objLookupMaster = new LookupMaster();
                        General.ReflectSingleData(objLookupMaster, LookupResult);

                        CSUList.Add(objLookupMaster);
                    }
                }


                objUnderlying.CSUCategoryList = CSUList;
                #endregion

                #region Put Spread Uploads
                List<LookupMaster> PSUList = new List<LookupMaster>();

                objLookupResult = objSP_PRICINGEntities.SP_FETCH_LOOKUP_VALUES("PSU", false);
                LookupResultList = objLookupResult.ToList();

                if (LookupResultList != null && LookupResultList.Count > 0)
                {
                    foreach (var LookupResult in LookupResultList)
                    {
                        LookupMaster objLookupMaster = new LookupMaster();
                        General.ReflectSingleData(objLookupMaster, LookupResult);

                        PSUList.Add(objLookupMaster);
                    }
                }


                objUnderlying.PSUCategoryList = PSUList;
                #endregion

                #region SubType List
                List<LookupMaster> SubTypeList = new List<LookupMaster>();

                objLookupResult = objSP_PRICINGEntities.SP_FETCH_LOOKUP_VALUES("USub", false);
                LookupResultList = objLookupResult.ToList();

                if (LookupResultList != null && LookupResultList.Count > 0)
                {
                    foreach (var LookupResult in LookupResultList)
                    {
                        LookupMaster objLookupMaster = new LookupMaster();
                        General.ReflectSingleData(objLookupMaster, LookupResult);

                        SubTypeList.Add(objLookupMaster);
                    }
                }
                objUnderlying.SubTypeList = SubTypeList;
                #endregion

                #region Underlying Type List
                List<LookupMaster> UnderlyingTypeList = new List<LookupMaster>();

                objLookupResult = objSP_PRICINGEntities.SP_FETCH_LOOKUP_VALUES("UT", true);
                LookupResultList = objLookupResult.ToList();

                if (LookupResultList != null && LookupResultList.Count > 0)
                {
                    foreach (var LookupResult in LookupResultList)
                    {
                        LookupMaster objLookupMaster = new LookupMaster();
                        General.ReflectSingleData(objLookupMaster, LookupResult);

                        UnderlyingTypeList.Add(objLookupMaster);
                    }
                }
                objUnderlying.UnderlyingTypeList = UnderlyingTypeList;
                #endregion

                return objUnderlying;
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "FetchDefaultDetails", objUserMaster.UserID);
                return objUnderlying;
            }
        }

        public JsonResult FetchEntityList()
        {
            try
            {
                ObjectResult<LookupResult> objLookupResult;
                List<LookupResult> LookupResultList;
                List<LookupMaster> EntityList = new List<LookupMaster>();

                objLookupResult = objSP_PRICINGEntities.SP_FETCH_LOOKUP_VALUES("EM", false);
                LookupResultList = objLookupResult.ToList();

                if (LookupResultList != null && LookupResultList.Count > 0)
                {
                    foreach (var LookupResult in LookupResultList)
                    {
                        LookupMaster objLookupMaster = new LookupMaster();
                        General.ReflectSingleData(objLookupMaster, LookupResult);

                        EntityList.Add(objLookupMaster);
                    }
                }

                var EntityListData = EntityList.ToList();
                return Json(EntityListData, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "FetchUnderlyingList", objUserMaster.UserID);
                return Json("");
            }

            //return Json(UnderlyingListData);
        }

        public JsonResult FetchIsSecuredList()
        {
            try
            {
                ObjectResult<LookupResult> objLookupResult;
                List<LookupResult> LookupResultList;
                List<LookupMaster> IsSecuredList = new List<LookupMaster>();

                objLookupResult = objSP_PRICINGEntities.SP_FETCH_LOOKUP_VALUES("IS", false);
                LookupResultList = objLookupResult.ToList();

                if (LookupResultList != null && LookupResultList.Count > 0)
                {
                    foreach (var LookupResult in LookupResultList)
                    {
                        LookupMaster objLookupMaster = new LookupMaster();
                        General.ReflectSingleData(objLookupMaster, LookupResult);

                        IsSecuredList.Add(objLookupMaster);
                    }
                }

                var IsSecuredListData = IsSecuredList.ToList();
                return Json(IsSecuredListData, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "FetchUnderlyingList", objUserMaster.UserID);
                return Json("");
            }

            //return Json(UnderlyingListData);
        }
    }
}