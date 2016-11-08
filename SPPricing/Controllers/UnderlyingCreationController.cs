using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using SPPricing.Models;
using System.Data.Objects;
using System.IO;
using System.Data;
using OfficeOpenXml;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Data.Common;

namespace SPPricing.Controllers
{
    public class UnderlyingCreationController : Controller
    {
        //
        // GET: /UnderlyingCreation/

        SP_PRICINGEntities objSP_PRICINGEntities = new SP_PRICINGEntities();

        //List<UploadFileMaster> UploadFileMasterList = new List<UploadFileMaster>();

        public ActionResult Index()
        {
            return View();
        }

        [HttpGet]
        public ActionResult UnderlyingCreation()
        {
            LoginController objLoginController = new LoginController();
            try
            {
                if (ValidateSession())
                {
                    Underlying objUnderlying = FetchDefaultDetails();

                    #region Menu Access By on Role

                    Int32 intResult = 0;
                    // bool PPorNonPP = false;

                    UserMaster objUserMaster = new UserMaster();
                    objUserMaster = (UserMaster)Session["LoggedInUser"];

                    var Result = objSP_PRICINGEntities.VALIDATE_MENU_BY_ROLE(objUserMaster.RoleID, "UC");
                    intResult = Convert.ToInt32(Result.SingleOrDefault());

                    if (intResult == 0)
                        return RedirectToAction("UserNotAuthorize", "Login");

                    #endregion

                    var a = TempData["CreateUnderlying"];
                    if (a != null)
                    {
                        ViewBag.CreateUnderlying = a;
                    }

                    FetchUploadFileMasterList();
                    objUnderlying.UnderlyingID = Convert.ToInt32(Session["UnderlyingID"]);
                    //objUnderlying.UnderlyingID = underlyingID;

                    return View(objUnderlying);
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
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "UnderlyingCreation Get", objUserMaster.UserID);
                return RedirectToAction("ErrorDetails", "Login");
            }
        }

        [HttpGet]
        public ActionResult UnderlyingCreationList()
        {
            LoginController objLoginController = new LoginController();
            try
            {
                if (ValidateSession())
                {

                    Underlying objUnderlying = FetchDefaultDetails();

                    #region Menu Access By on Role

                    Int32 intResult = 0;
                    // bool PPorNonPP = false;

                    UserMaster objUserMaster = new UserMaster();
                    objUserMaster = (UserMaster)Session["LoggedInUser"];

                    var Result = objSP_PRICINGEntities.VALIDATE_MENU_BY_ROLE(objUserMaster.RoleID, "UCL");
                    intResult = Convert.ToInt32(Result.SingleOrDefault());

                    if (intResult == 0)
                        return RedirectToAction("UserNotAuthorize", "Login");

                    #endregion

                    FetchUploadFileMasterList();

                    DataSet dsResult = new DataSet();
                    dsResult = General.ExecuteDataSet("FETCH_UNDERLYING_FILTER");

                    if (dsResult != null && dsResult.Tables.Count == 5)
                    {
                        //----------Fetch Underlying Name List-------------START----------
                        List<Underlying> underlyingNameList = new List<Underlying>();
                        foreach (DataRow dr in dsResult.Tables[0].Rows)
                        {
                            Underlying objName = new Underlying();
                            objName.UnderlyingName = Convert.ToString(dr["UNDERLYING_NAME"]);
                            underlyingNameList.Add(objName);
                        }

                        objUnderlying.NameList = underlyingNameList;
                        //----------Fetch UNDERLYING_NAME List-------------END------------

                        //----------Fetch UNDERLYING_SHORTNAME List-------------START----------
                        List<Underlying> underlyingShortNameList = new List<Underlying>();
                        foreach (DataRow dr in dsResult.Tables[1].Rows)
                        {
                            Underlying objShortName = new Underlying();
                            objShortName.UnderlyingShortName = Convert.ToString(dr["UNDERLYING_SHORTNAME"]);
                            underlyingShortNameList.Add(objShortName);
                        }

                        objUnderlying.ShortNameList = underlyingShortNameList;
                        //----------Fetch UNDERLYING_SHORTNAME List-------------END------------

                        //----------Fetch UNDERLYING_TYPE List-------------START----------
                        List<Underlying> underlyingTypeList = new List<Underlying>();
                        foreach (DataRow dr in dsResult.Tables[2].Rows)
                        {
                            Underlying objType = new Underlying();
                            objType.UnderlyingType = Convert.ToString(dr["UNDERLYING_TYPE"]);
                            underlyingTypeList.Add(objType);
                        }

                        objUnderlying.TypeList = underlyingTypeList;
                        //----------Fetch UNDERLYING_TYPE List-------------END------------

                        //----------Fetch STANDARD List-------------START----------
                        List<Underlying> underlyingStandardList = new List<Underlying>();
                        foreach (DataRow dr in dsResult.Tables[3].Rows)
                        {
                            Underlying objStandard = new Underlying();
                            objStandard.StandardName = Convert.ToString(dr["STANDARD"]);
                            underlyingStandardList.Add(objStandard);
                        }

                        objUnderlying.UnderlyingStandardList = underlyingStandardList;
                        //----------Fetch STANDARD List-------------END------------

                        //----------Fetch SUB_TYPE List-------------START----------
                        List<Underlying> underlyingSubTypeList = new List<Underlying>();
                        foreach (DataRow dr in dsResult.Tables[4].Rows)
                        {
                            Underlying objSubType = new Underlying();
                            objSubType.SubTypeName = Convert.ToString(dr["SUB_TYPE"]);
                            underlyingSubTypeList.Add(objSubType);
                        }

                        objUnderlying.UnderlyingSubTypeList = underlyingSubTypeList;
                        //----------Fetch SUB_TYPE List-------------END------------
                    }

                    objUnderlying.UnderlyingID = 0;
                    Session["UnderlyingID"] = objUnderlying.UnderlyingID;

                    return View(objUnderlying);
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
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "UnderlyingCreation Post", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        public JsonResult DeleteUnderlying(string UnderlyingID, string UnderlyingName)
        {
            try
            {
                Int32 intResult = 0;
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];

                var Result = objSP_PRICINGEntities.SP_DELETE_UNDERLYING(UnderlyingID, UnderlyingName, objUserMaster.UserID);
                intResult = Convert.ToInt32(Result.SingleOrDefault());

                if (intResult == 1)
                    ViewBag.DeleteMessage = "Data Deleted Successfully";
                return Json(intResult);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "MasterController", "DeleteLineOfQuote", objUserMaster.UserID);
                return Json("");
            }
        }

        [HttpGet]
        public ActionResult UnderlyingCreationEdit(int underlyingID, bool? blnUploadStatus = false)
        {
            LoginController objLoginController = new LoginController();
            try
            {
                if (ValidateSession())
                {
                    Underlying objUnderlying = FetchDefaultDetails();

                    #region Menu Access By on Role

                    Int32 intResult = 0;
                    // bool PPorNonPP = false;

                    UserMaster objUserMaster = new UserMaster();
                    objUserMaster = (UserMaster)Session["LoggedInUser"];

                    var Result = objSP_PRICINGEntities.VALIDATE_MENU_BY_ROLE(objUserMaster.RoleID, "UCE");
                    intResult = Convert.ToInt32(Result.SingleOrDefault());

                    if (intResult == 0)
                        return RedirectToAction("UserNotAuthorize", "Login");

                    #endregion

                    FetchUploadFileMasterList();
                    DataSet dsResult = new DataSet();
                    dsResult = General.ExecuteDataSet("FETCH_UNDERLYING_CREATION_BYID", underlyingID);

                    if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                    {
                        objUnderlying.UnderlyingName = Convert.ToString(dsResult.Tables[0].Rows[0]["UNDERLYING_NAME"]);
                        objUnderlying.UnderlyingShortName = Convert.ToString(dsResult.Tables[0].Rows[0]["UNDERLYING_SHORTNAME"]);
                        objUnderlying.FilterUnderlyingType = Convert.ToInt32(dsResult.Tables[0].Rows[0]["UNDERLYING_TYPE"]);
                        objUnderlying.FilterStandard = Convert.ToInt32(dsResult.Tables[0].Rows[0]["STANDARD"]);
                        objUnderlying.FilterSubType = Convert.ToInt32(dsResult.Tables[0].Rows[0]["SUB_TYPE"]);
                        objUnderlying.Tickers = Convert.ToString(dsResult.Tables[0].Rows[0]["TICKERS"]);                        
                        objUnderlying.AutocallIV = Convert.ToDouble(dsResult.Tables[0].Rows[0]["AUTOCALL_IV"]);
                        objUnderlying.AutocallRFR = Convert.ToDouble(dsResult.Tables[0].Rows[0]["AUTOCALL_RFR"]);
                    }
                    objUnderlying.UnderlyingID = underlyingID;

                    if (blnUploadStatus == true)
                    {
                        ViewBag.Message = "Imported Successfully";
                    }

                    // return View("UnderlyingCreationEdit", objUnderlying);


                    return View(objUnderlying);
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
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "UnderlyingCreationEdit Get", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        public JsonResult FetchUnderlyingCreationList(string UnderlyingName, string UnderlyingShortName, string UnderlyingType, string UnderlyingStandard, string UnderlyingSubType)
        {
            try
            {
                List<Underlying> UnderlyingList = new List<Underlying>();

                if (UnderlyingName == "" || UnderlyingName == "--Select--")
                    UnderlyingName = "ALL";

                if (UnderlyingShortName == "" || UnderlyingShortName == "--Select--")
                    UnderlyingShortName = "ALL";

                if (UnderlyingType == "" || UnderlyingType == "--Select--")
                    UnderlyingType = "ALL";

                if (UnderlyingStandard == "" || UnderlyingStandard == "--Select--")
                    UnderlyingStandard = "ALL";

                if (UnderlyingSubType == "" || UnderlyingSubType == "--Select--")
                    UnderlyingSubType = "ALL";

                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("FETCH_UNDERLYING_CREATION", UnderlyingName, UnderlyingShortName, UnderlyingType, UnderlyingStandard, UnderlyingSubType);

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in dsResult.Tables[0].Rows)
                    {
                        Underlying obj = new Underlying();

                        obj.UnderlyingID = Convert.ToInt32(dr["UnderlyingID"]);
                        obj.UnderlyingName = Convert.ToString(dr["UnderlyingName"]);
                        obj.UnderlyingShortName = Convert.ToString(dr["UnderlyingShortName"]);
                        obj.UnderlyingType = Convert.ToString(dr["UnderlyingType"]);
                        obj.StandardName = Convert.ToString(dr["StandardName"]);
                        obj.SubTypeName = Convert.ToString(dr["SubTypeName"]);

                        UnderlyingList.Add(obj);
                    }
                }

                var UnderlyingListData = UnderlyingList.ToList();
                return Json(UnderlyingListData, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "UnderlyingCreation", objUserMaster.UserID);
                return Json("");
            }
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

        protected bool ReadDataFromExcel(string strFilePath, Int32 intUnderlyingID)
        {
            try
            {
                FileInfo newFile = new FileInfo(strFilePath);

                DataTable dt = new DataTable();

                dt.Columns.Add("UNDERLYING_ID");
                dt.Columns.Add("Instrument");
                dt.Columns.Add("Weightage");

                using (ExcelPackage xlPackage = new ExcelPackage(newFile))
                {
                    ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets[1];
                    DataRow dr;

                    for (int iRow = 2; iRow < 1000; iRow++)
                    {
                        if (Convert.ToString(worksheet.Cell(iRow, 1).Value) != "")
                        {
                            dr = dt.NewRow();

                            dr["UNDERLYING_ID"] = intUnderlyingID;
                            dr["Instrument"] = worksheet.Cell(iRow, 1).Value;
                            dr["Weightage"] = worksheet.Cell(iRow, 2).Value;

                            dt.Rows.Add(dr);
                        }
                        else
                            break;
                    }
                }

                string strMyConnection = Convert.ToString(System.Configuration.ConfigurationManager.ConnectionStrings["SP_PRICINGConnectionString"]);

                string strSourceFileColumn = "UNDERLYING_ID|Instrument|Weightage";
                string strDestinationFileColumn = "UNDERLYING_ID|BASKET_INSTRUMENT|WEIGHTAGE";
                string strTableName = "TBL_BASKET_CORRELATION_TEMP";
                string[] arrSourceColumn = null;
                string[] arrDestinationColumn = null;

                if (strSourceFileColumn != "")
                    arrSourceColumn = strSourceFileColumn.Split('|');

                if (strDestinationFileColumn != "")
                    arrDestinationColumn = strDestinationFileColumn.Split('|');

                if (arrSourceColumn != null && arrDestinationColumn != null && arrSourceColumn.Length == arrDestinationColumn.Length)
                {
                    using (SqlBulkCopy bulkCopy = new SqlBulkCopy(strMyConnection))
                    {
                        bulkCopy.DestinationTableName = strTableName;
                        bulkCopy.BulkCopyTimeout = 1000;

                        for (int i = 0; i < arrSourceColumn.Length; i++)
                        {
                            bulkCopy.ColumnMappings.Add(arrSourceColumn[i], arrDestinationColumn[i]);
                        }
                        bulkCopy.WriteToServer(dt);
                    }

                    return true;
                }
                else
                    return false;
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "FetchDefaultDetails", objUserMaster.UserID);
                return false;
            }
        }

        [HttpPost]
        public void UploadHandler()
        {
            var files = Request.Files;

            if (files != null && files.Count > 0)
            {
                DataTable dtData = new DataTable();
                foreach (string filename in files)
                {
                    var file = files[filename];
                    if (file.ContentLength > 0)
                    {
                        //TODO : Save Logic
                        List<UploadFileMaster> UploadFileMasterList = (List<UploadFileMaster>)Session["UploadFileMasterList"];
                        UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "BC"; });
                        string strFilePath = System.Web.HttpContext.Current.Server.MapPath("~/Uploads/");
                        string strFilename = Path.GetFileName(file.FileName);
                        // string strExtension = Path.GetExtension(file.FileName);

                        strFilePath += strFilename;// +strExtension;
                        file.SaveAs(strFilePath);

                        FileInfo newFile = new FileInfo(strFilePath);

                        #region Source and Destination Column
                        string strSourceColumn = objUploadFileMaster.SourceColumn;
                        string[] arrSourceColumn = null;
                        if (strSourceColumn != "")
                            arrSourceColumn = strSourceColumn.Split('|');



                        for (int i = 0; i < arrSourceColumn.Length; i++)
                        {
                            dtData.Columns.Add(arrSourceColumn[i]);
                        }

                        DataTable dtColumnList = new DataTable();
                        dtColumnList.Columns.Add("ColumnName");

                        string strDestinationColumn = objUploadFileMaster.DestinationColumn;
                        string[] arrDestinationColumn = null;

                        if (strDestinationColumn != "")
                            arrDestinationColumn = strDestinationColumn.Split('|');

                        string strTableName = objUploadFileMaster.TableName;
                        #endregion

                        using (ExcelPackage xlPackage = new ExcelPackage(newFile))
                        {
                            ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets[objUploadFileMaster.SheetName];

                            for (int intCol = 2; intCol < 20; intCol++)
                            {
                                if (Convert.ToString(worksheet.Cell(1, intCol).Value) != "")
                                {
                                    DataRow dr = dtColumnList.NewRow();
                                    dr["ColumnName"] = worksheet.Cell(1, intCol).Value;

                                    dtColumnList.Rows.Add(dr);
                                }
                                else
                                    break;
                            }

                            DataRow drNew;

                            for (int iRow = 2; iRow < 50; iRow++)
                            {
                                if (Convert.ToString(worksheet.Cell(iRow, 1).Value) != "")
                                {
                                    for (int intCol = 2; intCol < dtColumnList.Rows.Count + 2; intCol++)
                                    {
                                        drNew = dtData.NewRow();

                                        drNew["UNDERLYING_ID"] = 0;
                                        drNew["UNDERLYING_1"] = worksheet.Cell(iRow, 1).Value;
                                        drNew["UNDERLYING_2"] = dtColumnList.Rows[intCol - 2][0];

                                        if (worksheet.Cell(iRow, intCol).Value.Trim() != "")
                                            drNew["VALUE"] = worksheet.Cell(iRow, intCol).Value;
                                        else
                                            drNew["VALUE"] = "0";

                                        drNew["CREATED_BY"] = "1";
                                        drNew["CREATED_ON"] = DateTime.Now;

                                        dtData.Rows.Add(drNew);
                                    }
                                }
                                else
                                    break;
                            }
                            //}
                        }
                    }
                }
                Session["BasketCorrelationFileDetails"] = dtData;
            }
        }

        [HttpPost]
        public ActionResult UnderlyingCreation(Underlying objUnderlying, string Command, FormCollection collection, HttpPostedFileBase file, HttpPostedFileBase file1, HttpPostedFileBase RCfile,
            HttpPostedFileBase LVSfile, HttpPostedFileBase CSUfile, HttpPostedFileBase PSUfile)
        {
            LoginController objLoginController = new LoginController();
            try
            {
                if (ValidateSession())
                {
                    //var httpPostedFile = HttpContext.Request.Files[0];
                    bool blnUploadStatus = false;
                    bool blnUploadDataStatus = true;
                    List<UploadFileMaster> UploadFileMasterList = (List<UploadFileMaster>)Session["UploadFileMasterList"];

                    UserMaster objUserMaster = new UserMaster();
                    objUserMaster = (UserMaster)Session["LoggedInUser"];


                    #region Submit
                    if (Command == "Submit")
                    {
                        int Standard = Convert.ToInt32(collection.Get("STANDARD"));
                        int SubType = Convert.ToInt32(collection.Get("SUBTYPE"));
                        string Tickers = Convert.ToString(collection.Get("TICKERS"));
                        ObjectResult<UnderlyingResult> objUnderlyingResult = objSP_PRICINGEntities.SP_MANAGE_UNDERLYING_DETAILS(objUnderlying.UnderlyingID, objUnderlying.UnderlyingName, objUnderlying.UnderlyingShortName, objUnderlying.FilterUnderlyingType, Standard, SubType, Tickers, objUnderlying.AutocallIV, objUnderlying.AutocallRFR, objUserMaster.UserID);
                        List<UnderlyingResult> UnderlyingResultList = objUnderlyingResult.ToList();
                        var tickerExist = objSP_PRICINGEntities.SP_MANAGE_UNDERLYING_TICKER(Tickers).ToList();


                        objUnderlying = new Underlying();
                        objUnderlying = FetchDefaultDetails();


                        if (tickerExist[0] == 5)
                        {
                            ViewBag.TickerAlready = "Ticker Already Exist";
                            return View(objUnderlying);
                        }

                        if (UnderlyingResultList[0].RESULT == -1)
                        {
                            ViewBag.UnderlyingAlready = "Underlying Already Exist";
                            return View(objUnderlying);
                        }

                        if (UnderlyingResultList != null && UnderlyingResultList.Count > 0)
                            objUnderlying.UnderlyingID = Convert.ToInt32(UnderlyingResultList[0].UNDERLYING_ID);
                        Session["UnderlyingID"] = objUnderlying.UnderlyingID;
                        ViewBag.UnderlyingID = objUnderlying.UnderlyingID;

                        var BusinessInstrument = collection.Get("hdnBasketInstruments");
                        var Weightage = collection.Get("hdnWeightage");

                        if (BusinessInstrument != "0" && Weightage != "0")
                        {
                            var NewProductID = objSP_PRICINGEntities.SP_MANAGE_BASKET_INSTRUMENT(objUnderlying.UnderlyingID, BusinessInstrument, Weightage);
                        }

                        blnUploadStatus = false;
                        blnUploadDataStatus = true;

                        if (file != null && file.ContentLength > 0)
                        {
                            UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "BC"; });
                            string strFilePath = System.Web.HttpContext.Current.Server.MapPath("~/Uploads/");
                            string strFileName = Path.GetFileNameWithoutExtension(file.FileName) + DateTime.Now.ToString("dd_mmm_yyyy_hh_mm");
                            string strExtension = Path.GetExtension(file.FileName);
                            string Type = "";

                            Type = collection.Get("IV");

                            strFilePath += strFileName + strExtension;
                            file.SaveAs(strFilePath);

                            FileInfo newFile = new FileInfo(strFilePath);

                            #region Source and Destination Column
                            string strSourceColumn = objUploadFileMaster.SourceColumn;
                            string[] arrSourceColumn = null;
                            if (strSourceColumn != "")
                                arrSourceColumn = strSourceColumn.Split('|');

                            DataTable dtData = new DataTable();

                            for (int i = 0; i < arrSourceColumn.Length; i++)
                            {
                                dtData.Columns.Add(arrSourceColumn[i]);
                            }

                            DataTable dtColumnList = new DataTable();
                            dtColumnList.Columns.Add("ColumnName");

                            string strDestinationColumn = objUploadFileMaster.DestinationColumn;
                            string[] arrDestinationColumn = null;

                            if (strDestinationColumn != "")
                                arrDestinationColumn = strDestinationColumn.Split('|');

                            string strTableName = objUploadFileMaster.TableName;
                            #endregion

                            using (ExcelPackage xlPackage = new ExcelPackage(newFile))
                            {
                                ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets[objUploadFileMaster.SheetName];

                                for (int intCol = 2; intCol < 20; intCol++)
                                {
                                    if (Convert.ToString(worksheet.Cell(1, intCol).Value) != "")
                                    {
                                        DataRow dr = dtColumnList.NewRow();
                                        dr["ColumnName"] = worksheet.Cell(1, intCol).Value;

                                        dtColumnList.Rows.Add(dr);
                                    }
                                    else
                                        break;
                                }

                                DataRow drNew;

                                for (int iRow = 2; iRow < 50; iRow++)
                                {
                                    if (Convert.ToString(worksheet.Cell(iRow, 1).Value) != "")
                                    {
                                        for (int intCol = 2; intCol < dtColumnList.Rows.Count + 2; intCol++)
                                        {
                                            drNew = dtData.NewRow();

                                            drNew["UNDERLYING_ID"] = objUnderlying.UnderlyingID;
                                            drNew["UNDERLYING_1"] = worksheet.Cell(iRow, 1).Value;
                                            drNew["UNDERLYING_2"] = dtColumnList.Rows[intCol - 2][0];

                                            if (worksheet.Cell(iRow, intCol).Value.Trim() != "")
                                                drNew["VALUE"] = worksheet.Cell(iRow, intCol).Value;
                                            else
                                                drNew["VALUE"] = "0";

                                            drNew["CREATED_BY"] = "1";
                                            drNew["CREATED_ON"] = DateTime.Now;

                                            dtData.Rows.Add(drNew);
                                        }
                                    }
                                    else
                                        break;
                                }
                                //}
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
                                ManageUploadFileInfo(objUnderlying.UnderlyingID, file.FileName, strFilePath, blnUploadStatus, blnUploadDataStatus);

                                ViewBag.Message = "Imported successfully";
                            }
                        }

                        return View(objUnderlying);
                    }
                    #endregion

                    #region BasketUpload
                    //else if (Command == "BasketUpload")
                    //{
                    //    Int32 intUnderlyingID = Convert.ToInt32(Session["UnderlyingID"]);//(collection["UnderlyingID"]);
                    //    ViewBag.UnderlyingID = intUnderlyingID;
                    //    blnUploadStatus = false;
                    //    blnUploadDataStatus = true;

                    //    if (file != null && file.ContentLength > 0)
                    //    {
                    //        //string strFilePath = GenerateUniqueFileName(file);
                    //        //file.SaveAs(strFilePath);

                    //        //blnUploadStatus = ReadDataFromExcel(strFilePath, intUnderlyingID);

                    //        //ObjectResult<BasketCorrelationResult> objBasketCorrelationResult = objSP_PRICINGEntities.SP_MANAGE_BASKET_CORRELATION();
                    //        //List<BasketCorrelationResult> BasketCorrelationResultList = objBasketCorrelationResult.ToList();

                    //        //if (BasketCorrelationResultList != null && BasketCorrelationResultList.Count > 0)
                    //        //    blnUploadDataStatus = false;

                    //        UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "BC"; });
                    //        string strFilePath = System.Web.HttpContext.Current.Server.MapPath("~/Uploads/");
                    //        string strFileName = Path.GetFileNameWithoutExtension(file.FileName) + DateTime.Now.ToString("dd_mmm_yyyy_hh_mm");
                    //        string strExtension = Path.GetExtension(file.FileName);
                    //        string Type = "";

                    //        Type = collection.Get("IV");

                    //        strFilePath += strFileName + strExtension;
                    //        file.SaveAs(strFilePath);

                    //        FileInfo newFile = new FileInfo(strFilePath);

                    //        #region Source and Destination Column
                    //        string strSourceColumn = objUploadFileMaster.SourceColumn;
                    //        string[] arrSourceColumn = null;
                    //        if (strSourceColumn != "")
                    //            arrSourceColumn = strSourceColumn.Split('|');

                    //        DataTable dtData = new DataTable();

                    //        for (int i = 0; i < arrSourceColumn.Length; i++)
                    //        {
                    //            dtData.Columns.Add(arrSourceColumn[i]);
                    //        }

                    //        DataTable dtColumnList = new DataTable();
                    //        dtColumnList.Columns.Add("ColumnName");

                    //        string strDestinationColumn = objUploadFileMaster.DestinationColumn;
                    //        string[] arrDestinationColumn = null;

                    //        if (strDestinationColumn != "")
                    //            arrDestinationColumn = strDestinationColumn.Split('|');

                    //        string strTableName = objUploadFileMaster.TableName;
                    //        #endregion

                    //        using (ExcelPackage xlPackage = new ExcelPackage(newFile))
                    //        {
                    //            ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets[objUploadFileMaster.SheetName];

                    //            for (int intCol = 2; intCol < 20; intCol++)
                    //            {
                    //                if (Convert.ToString(worksheet.Cell(1, intCol).Value) != "")
                    //                {
                    //                    DataRow dr = dtColumnList.NewRow();
                    //                    dr["ColumnName"] = worksheet.Cell(1, intCol).Value;

                    //                    dtColumnList.Rows.Add(dr);
                    //                }
                    //                else
                    //                    break;
                    //            }

                    //            DataRow drNew;

                    //            for (int iRow = 2; iRow < 50; iRow++)
                    //            {
                    //                if (Convert.ToString(worksheet.Cell(iRow, 1).Value) != "")
                    //                {
                    //                    for (int intCol = 2; intCol < dtColumnList.Rows.Count + 2; intCol++)
                    //                    {
                    //                        drNew = dtData.NewRow();

                    //                        drNew["UNDERLYING_ID"] = intUnderlyingID;
                    //                        drNew["UNDERLYING_1"] = worksheet.Cell(iRow, 1).Value;
                    //                        drNew["UNDERLYING_2"] = dtColumnList.Rows[intCol - 2][0];

                    //                        if (worksheet.Cell(iRow, intCol).Value.Trim() != "")
                    //                            drNew["VALUE"] = worksheet.Cell(iRow, intCol).Value;
                    //                        else
                    //                            drNew["VALUE"] = "0";

                    //                        drNew["CREATED_BY"] = "1";
                    //                        drNew["CREATED_ON"] = DateTime.Now;

                    //                        dtData.Rows.Add(drNew);
                    //                    }
                    //                }
                    //                else
                    //                    break;
                    //            }
                    //            //}
                    //        }

                    //      string strMyConnection = Convert.ToString(System.Configuration.ConfigurationManager.ConnectionStrings["SP_PRICINGConnectionString"]);

                    //        if (arrSourceColumn != null && arrDestinationColumn != null && arrSourceColumn.Length == arrDestinationColumn.Length)
                    //        {
                    //            using (SqlBulkCopy bulkCopy = new SqlBulkCopy(strMyConnection))
                    //            {
                    //                bulkCopy.DestinationTableName = strTableName;

                    //                for (int i = 0; i < arrSourceColumn.Length; i++)
                    //                {
                    //                    bulkCopy.ColumnMappings.Add(arrSourceColumn[i], arrDestinationColumn[i]);
                    //                }
                    //                bulkCopy.WriteToServer(dtData);
                    //            }
                    //            blnUploadStatus = true;
                    //            DataSet dsIV = new DataSet();
                    //            dsIV = General.ExecuteDataSet(objUploadFileMaster.ExtraProcedure);
                    //        }
                    //        else
                    //        {
                    //            blnUploadStatus = false;
                    //        }

                    //        if (blnUploadStatus)
                    //        {
                    //            ManageUploadFileInfo(intUnderlyingID, file.FileName, strFilePath, blnUploadStatus, blnUploadDataStatus);

                    //            ViewBag.Message = "Imported successfully";
                    //        }
                    //    }

                    //    objUnderlying = new Underlying();
                    //    objUnderlying = FetchDefaultDetails();

                    //    return View(objUnderlying);
                    //}
                    #endregion

                    #region IVUpload
                    else if (Command == "IVUpload")
                    {
                        Int32 intUnderlyingID = Convert.ToInt32(Session["UnderlyingID"]);//(collection["UnderlyingID"]);
                        if (intUnderlyingID != 0)
                        {
                            ViewBag.UnderlyingID = intUnderlyingID;
                            blnUploadStatus = false;
                            blnUploadDataStatus = true;

                            if (file1 != null && file1.ContentLength > 0)
                            {
                                UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "IV"; });
                                string strFilePath = System.Web.HttpContext.Current.Server.MapPath("~/Uploads/");
                                string strFileName = Path.GetFileNameWithoutExtension(file1.FileName) + DateTime.Now.ToString("dd_mmm_yyyy_hh_mm");
                                string strExtension = Path.GetExtension(file1.FileName);
                                string Type = "";

                                Type = collection.Get("IV");

                                int Version = 0;

                                strFilePath += strFileName + strExtension;
                                file1.SaveAs(strFilePath);

                                FileInfo newFile = new FileInfo(strFilePath);

                                #region Source and Destination Column
                                string strSourceColumn = objUploadFileMaster.SourceColumn;
                                string[] arrSourceColumn = null;
                                if (strSourceColumn != "")
                                    arrSourceColumn = strSourceColumn.Split('|');

                                DataTable dtData = new DataTable();

                                for (int i = 0; i < arrSourceColumn.Length; i++)
                                {
                                    dtData.Columns.Add(arrSourceColumn[i]);
                                }

                                DataTable dtColumnList = new DataTable();
                                dtColumnList.Columns.Add("ColumnName");

                                string strDestinationColumn = objUploadFileMaster.DestinationColumn;
                                string[] arrDestinationColumn = null;

                                if (strDestinationColumn != "")
                                    arrDestinationColumn = strDestinationColumn.Split('|');

                                string strTableName = objUploadFileMaster.TableName;
                                #endregion

                                using (ExcelPackage xlPackage = new ExcelPackage(newFile))
                                {
                                    ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets[objUploadFileMaster.SheetName];

                                    for (int intCol = 2; intCol < 20; intCol++)
                                    {
                                        if (Convert.ToString(worksheet.Cell(3, intCol).Value) != "")
                                        {
                                            DataRow dr = dtColumnList.NewRow();
                                            dr["ColumnName"] = worksheet.Cell(3, intCol).Value;

                                            dtColumnList.Rows.Add(dr);
                                        }
                                        else
                                            break;
                                    }

                                    DataRow drNew;

                                    //string strType = "Call";
                                    // output the data in column 2
                                    for (int iRow = 4; iRow < 50; iRow++)
                                    {
                                        if (Convert.ToString(worksheet.Cell(iRow, 1).Value) != "")
                                        {
                                            for (int intCol = 0; intCol < dtColumnList.Rows.Count; intCol++)
                                            {
                                                drNew = dtData.NewRow();

                                                //DataSet dsResult1 = new DataSet();
                                                //dsResult1 = General.ExecuteDataSet("CHECK_VERSION", Type, intUnderlyingID);

                                                //if (dsResult1 != null && dsResult1.Tables.Count > 0 && dsResult1.Tables[0].Rows.Count > 0 && dsResult1.Tables[0].Rows[0]["VERSION"].ToString() != string.Empty)
                                                //{
                                                //    Version = Convert.ToInt32(dsResult1.Tables[0].Rows[0].ItemArray[0]) + 1;
                                                //}
                                                //else
                                                //{
                                                //    Version = 1;
                                                //}

                                                drNew["VERSION"] = Version;
                                                drNew["TENURE"] = worksheet.Cell(iRow, 1).Value;
                                                drNew["MONEYNESS"] = Convert.ToDouble(dtColumnList.Rows[intCol][0]);
                                                drNew["VALUE"] = worksheet.Cell(iRow, intCol + 2).Value;
                                                drNew["CREATED_DATE"] = DateTime.Now;
                                                drNew["TYPE"] = Type;
                                                drNew["UNDERLYINGID"] = intUnderlyingID;

                                                dtData.Rows.Add(drNew);
                                            }
                                        }
                                        else
                                            break;
                                    }
                                    //}
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
                                    dsIV = General.ExecuteDataSet(objUploadFileMaster.ExtraProcedure, true);
                                }
                                else
                                {
                                    blnUploadStatus = false;
                                }

                                if (blnUploadStatus)
                                {
                                    ManageUploadFileInfo(intUnderlyingID, file1.FileName, strFilePath, blnUploadStatus, blnUploadDataStatus);

                                    ViewBag.Message = "Imported successfully";
                                }

                                objUnderlying = new Underlying();
                                objUnderlying = FetchDefaultDetails();

                                return View(objUnderlying);
                            }
                            else
                            {
                                ViewBag.Message = "No File Uploaded";
                                objUnderlying = new Underlying();
                                objUnderlying = FetchDefaultDetails();

                                return View(objUnderlying);
                            }
                        }
                        else
                        {
                            ViewBag.Message = "UnderLying Creation Not Created";
                            objUnderlying = new Underlying();
                            objUnderlying = FetchDefaultDetails();

                            return View(objUnderlying);
                        }
                    }
                    #endregion

                    #region IVDownload
                    else if (Command == "IVDownload")
                    {
                        Int32 intUnderlyingID = Convert.ToInt32(collection["UnderlyingID"]);

                        UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "IV"; });
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
                        else
                            return View();
                    }
                    #endregion

                    #region Roll Cost
                    else if (Command == "RCUpload")
                    {
                        Int32 intUnderlyingID = Convert.ToInt32(Session["UnderlyingID"]);//(collection["UnderlyingID"]);
                        if (intUnderlyingID != 0)
                        {
                            ViewBag.UnderlyingID = intUnderlyingID;
                            blnUploadStatus = false;
                            blnUploadDataStatus = true;
                            if (RCfile != null && RCfile.ContentLength > 0)
                            {
                                UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "RC"; });
                                string strFilePath = System.Web.HttpContext.Current.Server.MapPath("~/Uploads/");
                                string strFileName = Path.GetFileNameWithoutExtension(RCfile.FileName) + DateTime.Now.ToString("dd_mmm_yyyy_hh_mm");
                                string strExtension = Path.GetExtension(RCfile.FileName);
                                string Type = "";

                                Type = collection.Get("RC");

                                int Version = 0;

                                strFilePath += strFileName + strExtension;
                                RCfile.SaveAs(strFilePath);

                                FileInfo newFile = new FileInfo(strFilePath);

                                #region Source and Destination Column
                                string strSourceColumn = objUploadFileMaster.SourceColumn;
                                string[] arrSourceColumn = null;
                                if (strSourceColumn != "")
                                    arrSourceColumn = strSourceColumn.Split('|');

                                DataTable dtData = new DataTable();

                                for (int i = 0; i < arrSourceColumn.Length; i++)
                                {
                                    dtData.Columns.Add(arrSourceColumn[i]);
                                }

                                DataTable dtColumnList = new DataTable();
                                dtColumnList.Columns.Add("ColumnName");

                                string strDestinationColumn = objUploadFileMaster.DestinationColumn;
                                string[] arrDestinationColumn = null;

                                if (strDestinationColumn != "")
                                    arrDestinationColumn = strDestinationColumn.Split('|');

                                string strTableName = objUploadFileMaster.TableName;
                                #endregion

                                using (ExcelPackage xlPackage = new ExcelPackage(newFile))
                                {
                                    ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets[objUploadFileMaster.SheetName];

                                    for (int intCol = 2; intCol < 20; intCol++)
                                    {
                                        if (Convert.ToString(worksheet.Cell(3, intCol).Value) != "")
                                        {
                                            DataRow dr = dtColumnList.NewRow();
                                            dr["ColumnName"] = worksheet.Cell(3, intCol).Value;

                                            dtColumnList.Rows.Add(dr);
                                        }
                                        else
                                            break;
                                    }

                                    DataRow drNew;

                                    //string strType = "Call";
                                    // output the data in column 2
                                    for (int iRow = 4; iRow < 50; iRow++)
                                    {
                                        if (Convert.ToString(worksheet.Cell(iRow, 1).Value) != "")
                                        {
                                            for (int intCol = 0; intCol < dtColumnList.Rows.Count; intCol++)
                                            {
                                                drNew = dtData.NewRow();

                                                drNew["VERSION"] = Version;
                                                drNew["FREQUENCY"] = worksheet.Cell(iRow, 1).Value;
                                                drNew["TENURE"] = Convert.ToString(dtColumnList.Rows[intCol][0]);
                                                drNew["VALUE"] = worksheet.Cell(iRow, intCol + 2).Value;
                                                drNew["CREATED_DATE"] = DateTime.Now;
                                                drNew["TYPE"] = Type;
                                                drNew["UNDERLYINGID"] = intUnderlyingID;

                                                dtData.Rows.Add(drNew);
                                            }
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
                                    dsIV = General.ExecuteDataSet(objUploadFileMaster.ExtraProcedure, true);
                                }
                                else
                                {
                                    blnUploadStatus = false;
                                }

                                if (blnUploadStatus)
                                {
                                    ManageUploadFileInfo(intUnderlyingID, RCfile.FileName, strFilePath, blnUploadStatus, blnUploadDataStatus);

                                    ViewBag.Message = "Imported successfully";
                                }

                                objUnderlying = new Underlying();
                                objUnderlying = FetchDefaultDetails();

                                return View(objUnderlying);
                            }
                            else
                            {
                                ViewBag.Message = "No File Uploaded";
                                objUnderlying = new Underlying();
                                objUnderlying = FetchDefaultDetails();

                                return View(objUnderlying);
                            }
                        }
                        else
                        {
                            ViewBag.Message = "UnderLying Creation Not Created";
                            objUnderlying = new Underlying();
                            objUnderlying = FetchDefaultDetails();

                            return View(objUnderlying);
                        }
                    }
                    #endregion

                    #region RCDownload
                    else if (Command == "RCDownload")
                    {
                        Int32 intUnderlyingID = Convert.ToInt32(collection["UnderlyingID"]);

                        UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "RC"; });
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
                        else
                            return View();
                    }
                    #endregion

                    #region LVSUpload
                    else if (Command == "LVSUpload")
                    {
                        Int32 intUnderlyingID = Convert.ToInt32(Session["UnderlyingID"]);//(collection["UnderlyingID"]);

                        if (intUnderlyingID != 0)
                        {
                            ViewBag.UnderlyingID = intUnderlyingID;
                            blnUploadStatus = false;
                            blnUploadDataStatus = true;
                            if (LVSfile != null && LVSfile.ContentLength > 0)
                            {
                                UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "LVS"; });
                                string strFilePath = System.Web.HttpContext.Current.Server.MapPath("~/Uploads/");
                                string strFileName = Path.GetFileNameWithoutExtension(LVSfile.FileName) + DateTime.Now.ToString("dd_mmm_yyyy_hh_mm");
                                string strExtension = Path.GetExtension(LVSfile.FileName);
                                string Type = "";

                                Type = collection.Get("LVS");
                                int Version = 0;

                                strFilePath += strFileName + strExtension;
                                LVSfile.SaveAs(strFilePath);

                                FileInfo newFile = new FileInfo(strFilePath);

                                #region Source and Destination Column
                                string strSourceColumn = objUploadFileMaster.SourceColumn;
                                string[] arrSourceColumn = null;
                                if (strSourceColumn != "")
                                    arrSourceColumn = strSourceColumn.Split('|');

                                DataTable dtData = new DataTable();

                                for (int i = 0; i < arrSourceColumn.Length; i++)
                                {
                                    dtData.Columns.Add(arrSourceColumn[i]);
                                }

                                DataTable dtColumnList = new DataTable();
                                dtColumnList.Columns.Add("ColumnName");

                                string strDestinationColumn = objUploadFileMaster.DestinationColumn;
                                string[] arrDestinationColumn = null;

                                if (strDestinationColumn != "")
                                    arrDestinationColumn = strDestinationColumn.Split('|');

                                string strTableName = objUploadFileMaster.TableName;
                                #endregion

                                using (ExcelPackage xlPackage = new ExcelPackage(newFile))
                                {
                                    ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets[objUploadFileMaster.SheetName];

                                    for (int intCol = 2; intCol < 20; intCol++)
                                    {
                                        if (Convert.ToString(worksheet.Cell(3, intCol).Value) != "")
                                        {
                                            DataRow dr = dtColumnList.NewRow();
                                            dr["ColumnName"] = worksheet.Cell(3, intCol).Value;

                                            dtColumnList.Rows.Add(dr);
                                        }
                                        else
                                            break;
                                    }

                                    DataRow drNew;

                                    //string strType = "Call";
                                    // output the data in column 2
                                    for (int iRow = 4; iRow < 50; iRow++)
                                    {
                                        if (Convert.ToString(worksheet.Cell(iRow, 1).Value) != "")
                                        {
                                            for (int intCol = 0; intCol < dtColumnList.Rows.Count; intCol++)
                                            {
                                                drNew = dtData.NewRow();

                                                //DataSet dsResult1 = new DataSet();
                                                //dsResult1 = General.ExecuteDataSet("CHECK_VERSION", Type, intUnderlyingID);

                                                //if (dsResult1 != null && dsResult1.Tables.Count > 0 && dsResult1.Tables[2].Rows.Count > 0 && dsResult1.Tables[2].Rows[0]["VERSION"].ToString() != string.Empty)
                                                //{
                                                //    Version = Convert.ToInt32(dsResult1.Tables[2].Rows[0].ItemArray[0]) + 1;
                                                //}
                                                //else
                                                //{
                                                //    Version = 1;
                                                //}

                                                drNew["VERSION"] = Version;
                                                drNew["NO_OF_DURATION"] = worksheet.Cell(iRow, 1).Value;
                                                drNew["MONEYNESS"] = Convert.ToDouble(dtColumnList.Rows[intCol][0]);
                                                drNew["VALUE"] = worksheet.Cell(iRow, intCol + 2).Value;
                                                drNew["CREATED_DATE"] = DateTime.Now;
                                                drNew["TYPE"] = Type;
                                                drNew["UNDERLYINGID"] = intUnderlyingID;

                                                dtData.Rows.Add(drNew);
                                            }
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
                                    dsIV = General.ExecuteDataSet(objUploadFileMaster.ExtraProcedure, true);
                                }
                                else
                                {
                                    blnUploadStatus = false;
                                }

                                if (blnUploadStatus)
                                {
                                    ManageUploadFileInfo(intUnderlyingID, LVSfile.FileName, strFilePath, blnUploadStatus, blnUploadDataStatus);

                                    ViewBag.Message = "Imported successfully";
                                }


                                objUnderlying = new Underlying();
                                objUnderlying = FetchDefaultDetails();

                                return View(objUnderlying);
                            }
                            else
                            {
                                ViewBag.Message = "No File Uploaded";
                                objUnderlying = new Underlying();
                                objUnderlying = FetchDefaultDetails();

                                return View(objUnderlying);
                            }
                        }
                        else
                        {
                            ViewBag.Message = "UnderLying Creation Not Created";
                            objUnderlying = new Underlying();
                            objUnderlying = FetchDefaultDetails();

                            return View(objUnderlying);
                        }
                    }
                    #endregion

                    #region LVSDownload
                    else if (Command == "LVSDownload")
                    {
                        Int32 intUnderlyingID = Convert.ToInt32(collection["UnderlyingID"]);

                        UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "LVS"; });
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
                        else
                            return View();
                    }
                    #endregion

                    #region CSUpload
                    else if (Command == "CSUpload")
                    {
                        Int32 intUnderlyingID = Convert.ToInt32(Session["UnderlyingID"]);//(collection["UnderlyingID"]);

                        if (intUnderlyingID != 0)
                        {
                            ViewBag.UnderlyingID = intUnderlyingID;
                            blnUploadStatus = false;
                            blnUploadDataStatus = true;
                            if (CSUfile != null && CSUfile.ContentLength > 0)
                            {

                                string strFilePath = System.Web.HttpContext.Current.Server.MapPath("~/Uploads/");
                                string strFileName = Path.GetFileNameWithoutExtension(CSUfile.FileName) + DateTime.Now.ToString("dd_mmm_yyyy_hh_mm");
                                string strExtension = Path.GetExtension(CSUfile.FileName);
                                string LookupID = "";



                                //var value = ((System.Collections.Specialized.NameValueCollection)(collection)).AllKeys[4];
                                LookupID = collection.Get("CSU");
                                #region Spread Adjustment Surface
                                if (LookupID == "20")
                                {
                                    UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "CAS"; });
                                    int Version = 0;

                                    strFilePath += strFileName + strExtension;
                                    CSUfile.SaveAs(strFilePath);

                                    FileInfo newFile = new FileInfo(strFilePath);

                                    #region Source and Destination Column
                                    string strSourceColumn = objUploadFileMaster.SourceColumn;
                                    string[] arrSourceColumn = null;
                                    if (strSourceColumn != "")
                                        arrSourceColumn = strSourceColumn.Split('|');

                                    DataTable dtData = new DataTable();

                                    for (int i = 0; i < arrSourceColumn.Length; i++)
                                    {
                                        dtData.Columns.Add(arrSourceColumn[i]);
                                    }

                                    DataTable dtColumnList = new DataTable();
                                    dtColumnList.Columns.Add("ColumnName");

                                    string strDestinationColumn = objUploadFileMaster.DestinationColumn;
                                    string[] arrDestinationColumn = null;

                                    if (strDestinationColumn != "")
                                        arrDestinationColumn = strDestinationColumn.Split('|');

                                    string strTableName = objUploadFileMaster.TableName;
                                    #endregion

                                    using (ExcelPackage xlPackage = new ExcelPackage(newFile))
                                    {
                                        ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets[objUploadFileMaster.SheetName];

                                        for (int intCol = 2; intCol < 20; intCol++)
                                        {
                                            if (Convert.ToString(worksheet.Cell(3, intCol).Value) != "")
                                            {
                                                DataRow dr = dtColumnList.NewRow();
                                                dr["ColumnName"] = worksheet.Cell(3, intCol).Value;

                                                dtColumnList.Rows.Add(dr);
                                            }
                                            else
                                                break;
                                        }

                                        DataRow drNew;

                                        //string strType = "Call";
                                        // output the data in column 2
                                        for (int iRow = 4; iRow < 50; iRow++)
                                        {
                                            if (Convert.ToString(worksheet.Cell(iRow, 1).Value) != "")
                                            {
                                                for (int intCol = 0; intCol < dtColumnList.Rows.Count; intCol++)
                                                {
                                                    drNew = dtData.NewRow();

                                                    //DataSet dsResult1 = new DataSet();
                                                    //dsResult1 = General.ExecuteDataSet("CHECK_VERSION", LookupID, intUnderlyingID);

                                                    //if (dsResult1 != null && dsResult1.Tables.Count > 0 && dsResult1.Tables[3].Rows.Count > 0 && dsResult1.Tables[3].Rows[0]["VERSION"].ToString() != string.Empty)
                                                    //{
                                                    //    Version = Convert.ToInt32(dsResult1.Tables[3].Rows[0].ItemArray[0]) + 1;
                                                    //}
                                                    //else
                                                    //{
                                                    //    Version = 1;
                                                    //}

                                                    drNew["VERSION"] = Version;
                                                    drNew["TENURE"] = worksheet.Cell(iRow, 1).Value;
                                                    drNew["MONEYNESS"] = Convert.ToDouble(dtColumnList.Rows[intCol][0]);
                                                    drNew["VALUE"] = worksheet.Cell(iRow, intCol + 2).Value;
                                                    drNew["CREATED_DATE"] = DateTime.Now;
                                                    drNew["TYPE"] = LookupID;
                                                    drNew["UNDERLYINGID"] = intUnderlyingID;

                                                    dtData.Rows.Add(drNew);
                                                }
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
                                        dsIV = General.ExecuteDataSet(objUploadFileMaster.ExtraProcedure, true);
                                    }
                                    else
                                    {
                                        blnUploadStatus = false;
                                    }

                                    if (blnUploadStatus)
                                    {
                                        ManageUploadFileInfo(intUnderlyingID, CSUfile.FileName, strFilePath, blnUploadStatus, blnUploadDataStatus);

                                        ViewBag.Message = "Imported successfully";
                                    }
                                }
                                #endregion

                                #region Threshhold Strike
                                else if (LookupID == "21")
                                {
                                    UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "CT"; });
                                    int Version = 0;

                                    strFilePath += strFileName + strExtension;
                                    CSUfile.SaveAs(strFilePath);

                                    FileInfo newFile = new FileInfo(strFilePath);

                                    #region Source and Destination Column
                                    string strSourceColumn = objUploadFileMaster.SourceColumn;
                                    string[] arrSourceColumn = null;
                                    if (strSourceColumn != "")
                                        arrSourceColumn = strSourceColumn.Split('|');

                                    DataTable dtData = new DataTable();

                                    for (int i = 0; i < arrSourceColumn.Length; i++)
                                    {
                                        dtData.Columns.Add(arrSourceColumn[i]);
                                    }

                                    DataTable dtColumnList = new DataTable();
                                    dtColumnList.Columns.Add("ColumnName");

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

                                        //string strType = "Call";
                                        // output the data in column 2
                                        for (int iRow = 4; iRow < 36; iRow++)
                                        {
                                            if (Convert.ToString(worksheet.Cell(iRow, 1).Value) != "")
                                            {
                                                drNew = dtData.NewRow();

                                                //DataSet dsResult1 = new DataSet();
                                                //dsResult1 = General.ExecuteDataSet("CHECK_VERSION", LookupID, intUnderlyingID);

                                                //if (dsResult1 != null && dsResult1.Tables.Count > 0 && dsResult1.Tables[4].Rows.Count > 0 && dsResult1.Tables[4].Rows[0]["VERSION"].ToString() != string.Empty)
                                                //{
                                                //    Version = Convert.ToInt32(dsResult1.Tables[4].Rows[0].ItemArray[0]) + 1;
                                                //}
                                                //else
                                                //{
                                                //    Version = 1;
                                                //}

                                                drNew["VERSION"] = Version;
                                                drNew["TENURE"] = worksheet.Cell(iRow, 1).Value;
                                                drNew["STRIKE_CUT_OFF"] = worksheet.Cell(iRow, 2).Value;
                                                drNew["CREATED_DATE"] = DateTime.Now;
                                                drNew["UNDERLYINGID"] = intUnderlyingID;

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
                                        dsIV = General.ExecuteDataSet(objUploadFileMaster.ExtraProcedure, true);
                                    }
                                    else
                                    {
                                        blnUploadStatus = false;
                                    }

                                    if (blnUploadStatus)
                                    {
                                        ManageUploadFileInfo(intUnderlyingID, CSUfile.FileName, strFilePath, blnUploadStatus, blnUploadDataStatus);

                                        ViewBag.Message = "Imported successfully";
                                    }
                                }
                                #endregion

                                #region Spread Minimum IV
                                else if (LookupID == "22")
                                {
                                    UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "CIV"; });
                                    int Version = 0;

                                    strFilePath += strFileName + strExtension;
                                    CSUfile.SaveAs(strFilePath);

                                    FileInfo newFile = new FileInfo(strFilePath);

                                    #region Source and Destination Column
                                    string strSourceColumn = objUploadFileMaster.SourceColumn;
                                    string[] arrSourceColumn = null;
                                    if (strSourceColumn != "")
                                        arrSourceColumn = strSourceColumn.Split('|');

                                    DataTable dtData = new DataTable();

                                    for (int i = 0; i < arrSourceColumn.Length; i++)
                                    {
                                        dtData.Columns.Add(arrSourceColumn[i]);
                                    }

                                    DataTable dtColumnList = new DataTable();
                                    dtColumnList.Columns.Add("ColumnName");

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

                                        //string strType = "Call";
                                        // output the data in column 2
                                        for (int iRow = 4; iRow < 36; iRow++)
                                        {
                                            if (Convert.ToString(worksheet.Cell(iRow, 1).Value) != "")
                                            {
                                                drNew = dtData.NewRow();

                                                //DataSet dsResult1 = new DataSet();
                                                //dsResult1 = General.ExecuteDataSet("CHECK_VERSION", LookupID, intUnderlyingID);

                                                //if (dsResult1 != null && dsResult1.Tables.Count > 0 && dsResult1.Tables[5].Rows.Count > 0 && dsResult1.Tables[4].Rows[0]["VERSION"].ToString() != string.Empty)
                                                //{
                                                //    Version = Convert.ToInt32(dsResult1.Tables[4].Rows[0].ItemArray[0]) + 1;
                                                //}
                                                //else
                                                //{
                                                //    Version = 1;
                                                //}

                                                drNew["VERSION"] = Version;
                                                drNew["TENURE"] = worksheet.Cell(iRow, 1).Value;
                                                drNew["STRIKE_DIFFERENCE"] = worksheet.Cell(iRow, 2).Value;
                                                drNew["MINIMUM_VOL_DIFFERENCE"] = worksheet.Cell(iRow, 3).Value;
                                                drNew["CREATED_DATE"] = DateTime.Now;
                                                drNew["UNDERLYINGID"] = intUnderlyingID;

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
                                        dsIV = General.ExecuteDataSet(objUploadFileMaster.ExtraProcedure, true);
                                    }
                                    else
                                    {
                                        blnUploadStatus = false;
                                    }

                                    if (blnUploadStatus)
                                    {
                                        ManageUploadFileInfo(intUnderlyingID, CSUfile.FileName, strFilePath, blnUploadStatus, blnUploadDataStatus);

                                        ViewBag.Message = "Imported successfully";
                                    }
                                }
                                #endregion

                                objUnderlying = new Underlying();
                                objUnderlying = FetchDefaultDetails();

                                return View(objUnderlying);
                            }
                            else
                            {
                                ViewBag.Message = "No File Uploaded";
                                objUnderlying = new Underlying();
                                objUnderlying = FetchDefaultDetails();

                                return View(objUnderlying);
                            }
                        }
                        else
                        {
                            ViewBag.Message = "UnderLying Creation Not Created";
                            objUnderlying = new Underlying();
                            objUnderlying = FetchDefaultDetails();

                            return View(objUnderlying);
                        }
                    }
                    #endregion

                    #region CSDownload
                    else if (Command == "CSDownload")
                    {
                        Int32 intUnderlyingID = Convert.ToInt32(collection["UnderlyingID"]);
                        string LookupID = "";
                        LookupID = collection.Get("CSU");

                        #region Spread Adjustment Surface
                        if (LookupID == "20")
                        {
                            UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "CAS"; });
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
                            else
                                return View();
                        }
                        #endregion

                        #region Call Strike Threshold
                        else if (LookupID == "21")
                        {
                            UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "CT"; });
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
                            else
                                return View();
                        }
                        #endregion

                        #region Call Spread Minimum
                        else if (LookupID == "22")
                        {
                            UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "CIV"; });
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
                            else
                                return View();
                        }
                        #endregion
                    }
                    #endregion

                    #region PSUpload
                    else if (Command == "PSUpload")
                    {
                        Int32 intUnderlyingID = Convert.ToInt32(Session["UnderlyingID"]);//(collection["UnderlyingID"]);

                        if (intUnderlyingID != 0)
                        {
                            ViewBag.UnderlyingID = intUnderlyingID;
                            blnUploadStatus = false;
                            blnUploadDataStatus = true;
                            if (PSUfile != null && PSUfile.ContentLength > 0)
                            {

                                string strFilePath = System.Web.HttpContext.Current.Server.MapPath("~/Uploads/");
                                string strFileName = Path.GetFileNameWithoutExtension(PSUfile.FileName) + DateTime.Now.ToString("dd_mmm_yyyy_hh_mm");
                                string strExtension = Path.GetExtension(PSUfile.FileName);
                                string LookupID = "";

                                LookupID = collection.Get("PSU");

                                #region Spread Adjustment Surface
                                if (LookupID == "24")
                                {
                                    UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "PAS"; });
                                    int Version = 0;

                                    strFilePath += strFileName + strExtension;
                                    PSUfile.SaveAs(strFilePath);

                                    FileInfo newFile = new FileInfo(strFilePath);

                                    #region Source and Destination Column
                                    string strSourceColumn = objUploadFileMaster.SourceColumn;
                                    string[] arrSourceColumn = null;
                                    if (strSourceColumn != "")
                                        arrSourceColumn = strSourceColumn.Split('|');

                                    DataTable dtData = new DataTable();

                                    for (int i = 0; i < arrSourceColumn.Length; i++)
                                    {
                                        dtData.Columns.Add(arrSourceColumn[i]);
                                    }

                                    DataTable dtColumnList = new DataTable();
                                    dtColumnList.Columns.Add("ColumnName");

                                    string strDestinationColumn = objUploadFileMaster.DestinationColumn;
                                    string[] arrDestinationColumn = null;

                                    if (strDestinationColumn != "")
                                        arrDestinationColumn = strDestinationColumn.Split('|');

                                    string strTableName = objUploadFileMaster.TableName;
                                    #endregion

                                    using (ExcelPackage xlPackage = new ExcelPackage(newFile))
                                    {
                                        ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets[objUploadFileMaster.SheetName];

                                        for (int intCol = 2; intCol < 20; intCol++)
                                        {
                                            if (Convert.ToString(worksheet.Cell(3, intCol).Value) != "")
                                            {
                                                DataRow dr = dtColumnList.NewRow();
                                                dr["ColumnName"] = worksheet.Cell(3, intCol).Value;

                                                dtColumnList.Rows.Add(dr);
                                            }
                                            else
                                                break;
                                        }

                                        DataRow drNew;

                                        //string strType = "Call";
                                        // output the data in column 2
                                        for (int iRow = 4; iRow < 50; iRow++)
                                        {
                                            if (Convert.ToString(worksheet.Cell(iRow, 1).Value) != "")
                                            {
                                                for (int intCol = 0; intCol < dtColumnList.Rows.Count; intCol++)
                                                {
                                                    drNew = dtData.NewRow();

                                                    //DataSet dsResult1 = new DataSet();
                                                    //dsResult1 = General.ExecuteDataSet("CHECK_VERSION", LookupID, intUnderlyingID);

                                                    //if (dsResult1 != null && dsResult1.Tables.Count > 0 && dsResult1.Tables[3].Rows.Count > 0 && dsResult1.Tables[3].Rows[0]["VERSION"].ToString() != string.Empty)
                                                    //{
                                                    //    Version = Convert.ToInt32(dsResult1.Tables[3].Rows[0].ItemArray[0]) + 1;
                                                    //}
                                                    //else
                                                    //{
                                                    //    Version = 1;
                                                    //}

                                                    drNew["VERSION"] = Version;
                                                    drNew["TENURE"] = worksheet.Cell(iRow, 1).Value;
                                                    drNew["MONEYNESS"] = Convert.ToDouble(dtColumnList.Rows[intCol][0]);
                                                    drNew["VALUE"] = worksheet.Cell(iRow, intCol + 2).Value;
                                                    drNew["CREATED_DATE"] = DateTime.Now;
                                                    drNew["TYPE"] = LookupID;
                                                    drNew["UNDERLYINGID"] = intUnderlyingID;

                                                    dtData.Rows.Add(drNew);
                                                }
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
                                        dsIV = General.ExecuteDataSet(objUploadFileMaster.ExtraProcedure, true);
                                    }
                                    else
                                    {
                                        blnUploadStatus = false;
                                    }

                                    if (blnUploadStatus)
                                    {
                                        ManageUploadFileInfo(intUnderlyingID, PSUfile.FileName, strFilePath, blnUploadStatus, blnUploadDataStatus);

                                        ViewBag.Message = "Imported successfully";
                                    }
                                }

                                #endregion

                                #region Spread Skew Adjustment
                                else if (LookupID == "25")
                                {
                                    UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "PS"; });
                                    int Version = 0;

                                    strFilePath += strFileName + strExtension;
                                    PSUfile.SaveAs(strFilePath);

                                    FileInfo newFile = new FileInfo(strFilePath);

                                    #region Source and Destination Column
                                    string strSourceColumn = objUploadFileMaster.SourceColumn;
                                    string[] arrSourceColumn = null;
                                    if (strSourceColumn != "")
                                        arrSourceColumn = strSourceColumn.Split('|');

                                    DataTable dtData = new DataTable();

                                    for (int i = 0; i < arrSourceColumn.Length; i++)
                                    {
                                        dtData.Columns.Add(arrSourceColumn[i]);
                                    }

                                    DataTable dtColumnList = new DataTable();
                                    dtColumnList.Columns.Add("ColumnName");

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

                                        //string strType = "Call";
                                        // output the data in column 2
                                        for (int iRow = 4; iRow < 50; iRow++)
                                        {
                                            if (Convert.ToString(worksheet.Cell(iRow, 1).Value) != "")
                                            {
                                                drNew = dtData.NewRow();

                                                //DataSet dsResult1 = new DataSet();
                                                //dsResult1 = General.ExecuteDataSet("CHECK_VERSION", LookupID, intUnderlyingID);

                                                //if (dsResult1 != null && dsResult1.Tables.Count > 0 && dsResult1.Tables[6].Rows.Count > 0 && dsResult1.Tables[6].Rows[0]["VERSION"].ToString() != string.Empty)
                                                //{
                                                //    Version = Convert.ToInt32(dsResult1.Tables[6].Rows[0].ItemArray[0]) + 1;
                                                //}
                                                //else
                                                //{
                                                //    Version = 1;
                                                //}

                                                drNew["VERSION"] = Version;
                                                drNew["TENURE"] = worksheet.Cell(iRow, 1).Value;
                                                drNew["STRIKE_DIFFERENCE"] = worksheet.Cell(iRow, 2).Value;
                                                drNew["MINIMUM_VOL_DIFFERENCE"] = worksheet.Cell(iRow, 3).Value;
                                                drNew["CREATED_DATE"] = DateTime.Now;
                                                drNew["UNDERLYINGID"] = intUnderlyingID;

                                                dtData.Rows.Add(drNew);
                                            }
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
                                        dsIV = General.ExecuteDataSet(objUploadFileMaster.ExtraProcedure, true);
                                    }
                                    else
                                    {
                                        blnUploadStatus = false;
                                    }

                                    if (blnUploadStatus)
                                    {
                                        ManageUploadFileInfo(intUnderlyingID, PSUfile.FileName, strFilePath, blnUploadStatus, blnUploadDataStatus);

                                        ViewBag.Message = "Imported successfully";
                                    }
                                }
                                #endregion

                                #region Spread Minimum
                                else if (LookupID == "26")
                                {
                                    UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "PIV"; });
                                    int Version = 0;

                                    strFilePath += strFileName + strExtension;
                                    PSUfile.SaveAs(strFilePath);

                                    FileInfo newFile = new FileInfo(strFilePath);

                                    #region Source and Destination Column
                                    string strSourceColumn = objUploadFileMaster.SourceColumn;
                                    string[] arrSourceColumn = null;
                                    if (strSourceColumn != "")
                                        arrSourceColumn = strSourceColumn.Split('|');

                                    DataTable dtData = new DataTable();

                                    for (int i = 0; i < arrSourceColumn.Length; i++)
                                    {
                                        dtData.Columns.Add(arrSourceColumn[i]);
                                    }

                                    DataTable dtColumnList = new DataTable();
                                    dtColumnList.Columns.Add("ColumnName");

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

                                        //string strType = "Call";
                                        // output the data in column 2
                                        for (int iRow = 4; iRow < 36; iRow++)
                                        {
                                            if (Convert.ToString(worksheet.Cell(iRow, 1).Value) != "")
                                            {
                                                drNew = dtData.NewRow();

                                                //DataSet dsResult1 = new DataSet();
                                                //dsResult1 = General.ExecuteDataSet("CHECK_VERSION", LookupID, intUnderlyingID);

                                                //if (dsResult1 != null && dsResult1.Tables.Count > 0 && dsResult1.Tables[7].Rows.Count > 0 && dsResult1.Tables[7].Rows[0]["VERSION"].ToString() != string.Empty)
                                                //{
                                                //    Version = Convert.ToInt32(dsResult1.Tables[7].Rows[0].ItemArray[0]) + 1;
                                                //}
                                                //else
                                                //{
                                                //    Version = 1;
                                                //}

                                                drNew["VERSION"] = Version;
                                                drNew["TENURE"] = worksheet.Cell(iRow, 1).Value;
                                                drNew["STRIKE_DIFFERENCE"] = worksheet.Cell(iRow, 2).Value;
                                                drNew["MINIMUM_VOL_DIFFERENCE"] = worksheet.Cell(iRow, 3).Value;
                                                drNew["CREATED_DATE"] = DateTime.Now;
                                                drNew["UNDERLYINGID"] = intUnderlyingID;

                                                dtData.Rows.Add(drNew);

                                            }
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
                                        dsIV = General.ExecuteDataSet(objUploadFileMaster.ExtraProcedure, true);
                                    }
                                    else
                                    {
                                        blnUploadStatus = false;
                                    }

                                    if (blnUploadStatus)
                                    {
                                        ManageUploadFileInfo(intUnderlyingID, PSUfile.FileName, strFilePath, blnUploadStatus, blnUploadDataStatus);

                                        ViewBag.Message = "Imported successfully";
                                    }
                                }
                                #endregion

                                #region Spread Adjustment Exception
                                if (LookupID == "111")
                                {
                                    UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "PAE"; });
                                    int Version = 0;

                                    strFilePath += strFileName + strExtension;
                                    PSUfile.SaveAs(strFilePath);

                                    FileInfo newFile = new FileInfo(strFilePath);

                                    #region Source and Destination Column
                                    string strSourceColumn = objUploadFileMaster.SourceColumn;
                                    string[] arrSourceColumn = null;
                                    if (strSourceColumn != "")
                                        arrSourceColumn = strSourceColumn.Split('|');

                                    DataTable dtData = new DataTable();

                                    for (int i = 0; i < arrSourceColumn.Length; i++)
                                    {
                                        dtData.Columns.Add(arrSourceColumn[i]);
                                    }

                                    DataTable dtColumnList = new DataTable();
                                    dtColumnList.Columns.Add("ColumnName");

                                    string strDestinationColumn = objUploadFileMaster.DestinationColumn;
                                    string[] arrDestinationColumn = null;

                                    if (strDestinationColumn != "")
                                        arrDestinationColumn = strDestinationColumn.Split('|');

                                    string strTableName = objUploadFileMaster.TableName;
                                    #endregion

                                    using (ExcelPackage xlPackage = new ExcelPackage(newFile))
                                    {
                                        ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets[objUploadFileMaster.SheetName];

                                        for (int intCol = 3; intCol < 20; intCol++)
                                        {
                                            if (Convert.ToString(worksheet.Cell(3, intCol).Value) != "")
                                            {
                                                DataRow dr = dtColumnList.NewRow();
                                                dr["ColumnName"] = worksheet.Cell(3, intCol).Value;

                                                dtColumnList.Rows.Add(dr);
                                            }
                                            else
                                                break;
                                        }

                                        DataRow drNew;

                                        //string strType = "Call";
                                        // output the data in column 2
                                        for (int iRow = 5; iRow < 50; iRow++)
                                        {
                                            if (Convert.ToString(worksheet.Cell(iRow, 1).Value) != "")
                                            {
                                                for (int intCol = 0; intCol < dtColumnList.Rows.Count; intCol++)
                                                {
                                                    drNew = dtData.NewRow();

                                                    drNew["VERSION"] = Version;
                                                    drNew["TENURE"] = Convert.ToDouble(dtColumnList.Rows[intCol][0]);
                                                    drNew["STRIKE"] = worksheet.Cell(iRow, 1).Value;
                                                    drNew["GAP"] = worksheet.Cell(iRow, 2).Value;
                                                    drNew["AVERAGING"] = worksheet.Cell(iRow, 3).Value;
                                                    drNew["VALUE"] = worksheet.Cell(iRow, intCol + 3).Value;
                                                    drNew["UNDERLYING_ID"] = intUnderlyingID;
                                                    drNew["CREATED_BY"] = objUserMaster.UserID;
                                                    drNew["CREATED_ON"] = DateTime.Now;


                                                    dtData.Rows.Add(drNew);
                                                }
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
                                        dsIV = General.ExecuteDataSet(objUploadFileMaster.ExtraProcedure, objUserMaster.UserID, true);
                                    }
                                    else
                                    {
                                        blnUploadStatus = false;
                                    }

                                    if (blnUploadStatus)
                                    {
                                        ManageUploadFileInfo(intUnderlyingID, PSUfile.FileName, strFilePath, blnUploadStatus, blnUploadDataStatus);

                                        ViewBag.Message = "Imported successfully";
                                    }
                                }

                                #endregion


                                objUnderlying = new Underlying();
                                objUnderlying = FetchDefaultDetails();

                                return View(objUnderlying);
                            }
                            else
                            {
                                ViewBag.Message = "No File Uploaded";
                                objUnderlying = new Underlying();
                                objUnderlying = FetchDefaultDetails();

                                return View(objUnderlying);
                            }
                        }
                        else
                        {
                            ViewBag.Message = "UnderLying Creation Not Created";
                            objUnderlying = new Underlying();
                            objUnderlying = FetchDefaultDetails();

                            return View(objUnderlying);
                        }
                    }
                    #endregion

                    #region PSDownload
                    else if (Command == "PSDownload")
                    {
                        Int32 intUnderlyingID = Convert.ToInt32(collection["UnderlyingID"]);
                        string LookupID = "";
                        LookupID = collection.Get("PSU");

                        #region Spread Adjustment Surface
                        if (LookupID == "24")
                        {
                            UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "PAS"; });
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
                            else
                                return View();
                        }
                        #endregion

                        #region Spread Skew Adjustment
                        else if (LookupID == "25")
                        {
                            UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "PS"; });
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
                            else
                                return View();
                        }
                        #endregion

                        #region Spread Minimum
                        else if (LookupID == "26")
                        {
                            UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "PIV"; });
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
                            else
                                return View();
                        }
                        #endregion

                        #region Spread Adjustment Surface
                        if (LookupID == "111")
                        {
                            UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "PAE"; });
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
                            else
                                return View();
                        }
                        #endregion
                    }
                    #endregion


                    objUnderlying = new Underlying();
                    objUnderlying = FetchDefaultDetails();
                    return View(objUnderlying);
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
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "UnderlyingCreation Post", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        [HttpPost]
        public ActionResult UnderlyingCreationEdit(Underlying objUnderlying, string Command, FormCollection collection, HttpPostedFileBase file, HttpPostedFileBase file1, HttpPostedFileBase RCfile,
            HttpPostedFileBase LVSfile, HttpPostedFileBase CSUfile, HttpPostedFileBase PSUfile)
        {
            LoginController objLoginController = new LoginController();
            try
            {
                if (ValidateSession())
                {

                    List<UploadFileMaster> UploadFileMasterList = (List<UploadFileMaster>)Session["UploadFileMasterList"];
                    //var httpPostedFile = HttpContext.Request.Files[0];
                    bool blnUploadStatus = false;
                    bool blnUploadDataStatus = true;

                    UserMaster objUserMaster = new UserMaster();
                    objUserMaster = (UserMaster)Session["LoggedInUser"];


                    #region Submit
                    if (Command == "Submit")
                    {
                        Int32 intUnderlyingID = Convert.ToInt32(collection["UnderlyingID"]);

                        ObjectResult<UnderlyingResult> objUnderlyingResult = objSP_PRICINGEntities.SP_MANAGE_UNDERLYING_DETAILS(objUnderlying.UnderlyingID, objUnderlying.UnderlyingName, objUnderlying.UnderlyingShortName, objUnderlying.FilterUnderlyingType, objUnderlying.FilterStandard, objUnderlying.FilterSubType, objUnderlying.Tickers, objUnderlying.AutocallIV, objUnderlying.AutocallRFR, objUserMaster.UserID);
                        List<UnderlyingResult> UnderlyingResultList = objUnderlyingResult.ToList();

                        objUnderlying = new Underlying();
                        objUnderlying = FetchDefaultDetails();

                        if (UnderlyingResultList != null && UnderlyingResultList.Count > 0)
                            objUnderlying.UnderlyingID = Convert.ToInt32(UnderlyingResultList[0].UNDERLYING_ID);

                        //return View(objUnderlying);
                        return RedirectToAction("UnderlyingCreationEdit", new { underlyingID = intUnderlyingID });
                    }
                    #endregion

                    #region BasketUpload
                    else if (Command == "BasketUpload")
                    {
                        Int32 intUnderlyingID = Convert.ToInt32(Session["UnderlyingID"]);//(collection["UnderlyingID"]);
                        ViewBag.UnderlyingID = intUnderlyingID;
                        blnUploadStatus = false;
                        blnUploadDataStatus = true;

                        if (file != null && file.ContentLength > 0)
                        {
                            //string strFilePath = GenerateUniqueFileName(file);
                            //file.SaveAs(strFilePath);

                            //blnUploadStatus = ReadDataFromExcel(strFilePath, intUnderlyingID);

                            //ObjectResult<BasketCorrelationResult> objBasketCorrelationResult = objSP_PRICINGEntities.SP_MANAGE_BASKET_CORRELATION();
                            //List<BasketCorrelationResult> BasketCorrelationResultList = objBasketCorrelationResult.ToList();

                            //if (BasketCorrelationResultList != null && BasketCorrelationResultList.Count > 0)
                            //    blnUploadDataStatus = false;

                            UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "BC"; });
                            string strFilePath = System.Web.HttpContext.Current.Server.MapPath("~/Uploads/");
                            string strFileName = Path.GetFileNameWithoutExtension(file.FileName) + DateTime.Now.ToString("dd_mmm_yyyy_hh_mm");
                            string strExtension = Path.GetExtension(file.FileName);
                            string Type = "";

                            Type = collection.Get("IV");

                            strFilePath += strFileName + strExtension;
                            file.SaveAs(strFilePath);

                            FileInfo newFile = new FileInfo(strFilePath);

                            #region Source and Destination Column
                            string strSourceColumn = objUploadFileMaster.SourceColumn;
                            string[] arrSourceColumn = null;
                            if (strSourceColumn != "")
                                arrSourceColumn = strSourceColumn.Split('|');

                            DataTable dtData = new DataTable();

                            for (int i = 0; i < arrSourceColumn.Length; i++)
                            {
                                dtData.Columns.Add(arrSourceColumn[i]);
                            }

                            DataTable dtColumnList = new DataTable();
                            dtColumnList.Columns.Add("ColumnName");

                            string strDestinationColumn = objUploadFileMaster.DestinationColumn;
                            string[] arrDestinationColumn = null;

                            if (strDestinationColumn != "")
                                arrDestinationColumn = strDestinationColumn.Split('|');

                            string strTableName = objUploadFileMaster.TableName;
                            #endregion

                            using (ExcelPackage xlPackage = new ExcelPackage(newFile))
                            {
                                ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets[objUploadFileMaster.SheetName];

                                for (int intCol = 2; intCol < 20; intCol++)
                                {
                                    if (Convert.ToString(worksheet.Cell(1, intCol).Value) != "")
                                    {
                                        DataRow dr = dtColumnList.NewRow();
                                        dr["ColumnName"] = worksheet.Cell(1, intCol).Value;

                                        dtColumnList.Rows.Add(dr);
                                    }
                                    else
                                        break;
                                }

                                DataRow drNew;

                                for (int iRow = 2; iRow < 50; iRow++)
                                {
                                    if (Convert.ToString(worksheet.Cell(iRow, 1).Value) != "")
                                    {
                                        for (int intCol = 2; intCol < dtColumnList.Rows.Count + 2; intCol++)
                                        {
                                            drNew = dtData.NewRow();

                                            drNew["UNDERLYING_ID"] = intUnderlyingID;
                                            drNew["UNDERLYING_1"] = worksheet.Cell(iRow, 1).Value;
                                            drNew["UNDERLYING_2"] = dtColumnList.Rows[intCol - 2][0];

                                            if (worksheet.Cell(iRow, intCol).Value.Trim() != "")
                                                drNew["VALUE"] = worksheet.Cell(iRow, intCol).Value;
                                            else
                                                drNew["VALUE"] = "0";

                                            drNew["CREATED_BY"] = "1";
                                            drNew["CREATED_ON"] = DateTime.Now;

                                            dtData.Rows.Add(drNew);
                                        }
                                    }
                                    else
                                        break;
                                }
                                //}
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
                                ManageUploadFileInfo(intUnderlyingID, file.FileName, strFilePath, blnUploadStatus, blnUploadDataStatus);

                                ViewBag.Message = "Imported successfully";
                            }
                        }

                        objUnderlying = new Underlying();
                        objUnderlying = FetchDefaultDetails();

                        return View(objUnderlying);
                    }
                    #endregion

                    #region IVUpload
                    else if (Command == "IVUpload")
                    {
                        Int32 intUnderlyingID = Convert.ToInt32(collection["UnderlyingID"]);
                        if (intUnderlyingID != 0)
                        {
                            blnUploadStatus = false;
                            blnUploadDataStatus = true;

                            if (file1 != null && file1.ContentLength > 0)
                            {
                                UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "IV"; });
                                string strFilePath = System.Web.HttpContext.Current.Server.MapPath("~/Uploads/");
                                string strFileName = Path.GetFileNameWithoutExtension(file1.FileName) + DateTime.Now.ToString("dd_mmm_yyyy_hh_mm");
                                string strExtension = Path.GetExtension(file1.FileName);
                                string Type = "";

                                Type = collection.Get("IV");

                                int Version = 0;

                                strFilePath += strFileName + strExtension;
                                file1.SaveAs(strFilePath);

                                FileInfo newFile = new FileInfo(strFilePath);

                                #region Source and Destination Column
                                string strSourceColumn = objUploadFileMaster.SourceColumn;
                                string[] arrSourceColumn = null;
                                if (strSourceColumn != "")
                                    arrSourceColumn = strSourceColumn.Split('|');

                                DataTable dtData = new DataTable();

                                for (int i = 0; i < arrSourceColumn.Length; i++)
                                {
                                    dtData.Columns.Add(arrSourceColumn[i]);
                                }

                                DataTable dtColumnList = new DataTable();
                                dtColumnList.Columns.Add("ColumnName");

                                string strDestinationColumn = objUploadFileMaster.DestinationColumn;
                                string[] arrDestinationColumn = null;

                                if (strDestinationColumn != "")
                                    arrDestinationColumn = strDestinationColumn.Split('|');

                                string strTableName = objUploadFileMaster.TableName;
                                #endregion

                                using (ExcelPackage xlPackage = new ExcelPackage(newFile))
                                {
                                    ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets[objUploadFileMaster.SheetName];

                                    for (int intCol = 2; intCol < 20; intCol++)
                                    {
                                        if (Convert.ToString(worksheet.Cell(3, intCol).Value) != "")
                                        {
                                            DataRow dr = dtColumnList.NewRow();
                                            dr["ColumnName"] = worksheet.Cell(3, intCol).Value;

                                            dtColumnList.Rows.Add(dr);
                                        }
                                        else
                                            break;
                                    }

                                    DataRow drNew;

                                    //string strType = "Call";
                                    // output the data in column 2
                                    for (int iRow = 4; iRow < 50; iRow++)
                                    {
                                        if (Convert.ToString(worksheet.Cell(iRow, 1).Value) != "")
                                        {
                                            for (int intCol = 0; intCol < dtColumnList.Rows.Count; intCol++)
                                            {
                                                drNew = dtData.NewRow();

                                                //DataSet dsResult1 = new DataSet();
                                                //dsResult1 = General.ExecuteDataSet("CHECK_VERSION", Type, intUnderlyingID);

                                                //if (dsResult1 != null && dsResult1.Tables.Count > 0 && dsResult1.Tables[0].Rows.Count > 0 && dsResult1.Tables[0].Rows[0]["VERSION"].ToString() != string.Empty)
                                                //{
                                                //    Version = Convert.ToInt32(dsResult1.Tables[0].Rows[0].ItemArray[0]) + 1;
                                                //}
                                                //else
                                                //{
                                                //    Version = 1;
                                                //}

                                                drNew["VERSION"] = Version;
                                                drNew["TENURE"] = worksheet.Cell(iRow, 1).Value;
                                                drNew["MONEYNESS"] = Convert.ToDouble(dtColumnList.Rows[intCol][0]);
                                                drNew["VALUE"] = worksheet.Cell(iRow, intCol + 2).Value;
                                                drNew["CREATED_DATE"] = DateTime.Now;
                                                drNew["TYPE"] = Type;
                                                drNew["UNDERLYINGID"] = intUnderlyingID;

                                                dtData.Rows.Add(drNew);
                                            }
                                        }
                                        else
                                            break;
                                    }
                                    //}
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
                                    dsIV = General.ExecuteDataSet(objUploadFileMaster.ExtraProcedure, true);
                                }
                                else
                                {
                                    blnUploadStatus = false;
                                }

                                if (blnUploadStatus)
                                {
                                    ManageUploadFileInfo(intUnderlyingID, file1.FileName, strFilePath, blnUploadStatus, blnUploadDataStatus);

                                    ViewBag.Message = "Imported successfully";
                                }

                                objUnderlying = new Underlying();
                                objUnderlying = FetchDefaultDetails();

                                return RedirectToAction("UnderlyingCreationEdit", new { underlyingID = intUnderlyingID, blnUploadStatus = blnUploadStatus });
                            }
                            else
                            {
                                ViewBag.Message = "No File Uploaded";
                                objUnderlying = new Underlying();
                                objUnderlying = FetchDefaultDetails();

                                //return View(objUnderlying);
                                return RedirectToAction("UnderlyingCreationEdit", new { underlyingID = intUnderlyingID });
                            }
                        }
                        else
                        {
                            ViewBag.Message = "UnderLying Creation Not Created";
                            objUnderlying = new Underlying();
                            objUnderlying = FetchDefaultDetails();

                            return RedirectToAction("UnderlyingCreationEdit", new { underlyingID = intUnderlyingID });
                        }
                    }
                    #endregion

                    #region IVDownload
                    else if (Command == "IVDownload")
                    {
                        Int32 intUnderlyingID = Convert.ToInt32(collection["UnderlyingID"]);

                        UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "IV"; });
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

                            objUnderlying = new Underlying();
                            objUnderlying = FetchDefaultDetails();

                            return RedirectToAction("UnderlyingCreationEdit", new { underlyingID = intUnderlyingID });
                        }
                        else
                            return View();
                    }
                    #endregion

                    #region RCUpload
                    else if (Command == "RCUpload")
                    {
                        Int32 intUnderlyingID = Convert.ToInt32(collection["UnderlyingID"]);
                        if (intUnderlyingID != 0)
                        {
                            blnUploadStatus = false;
                            blnUploadDataStatus = true;
                            if (RCfile != null && RCfile.ContentLength > 0)
                            {
                                UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "RC"; });
                                string strFilePath = System.Web.HttpContext.Current.Server.MapPath("~/Uploads/");
                                string strFileName = Path.GetFileNameWithoutExtension(RCfile.FileName) + DateTime.Now.ToString("dd_mmm_yyyy_hh_mm");
                                string strExtension = Path.GetExtension(RCfile.FileName);
                                string Type = "";

                                Type = collection.Get("RC");

                                int Version = 0;

                                strFilePath += strFileName + strExtension;
                                RCfile.SaveAs(strFilePath);

                                FileInfo newFile = new FileInfo(strFilePath);

                                #region Source and Destination Column
                                string strSourceColumn = objUploadFileMaster.SourceColumn;
                                string[] arrSourceColumn = null;
                                if (strSourceColumn != "")
                                    arrSourceColumn = strSourceColumn.Split('|');

                                DataTable dtData = new DataTable();

                                for (int i = 0; i < arrSourceColumn.Length; i++)
                                {
                                    dtData.Columns.Add(arrSourceColumn[i]);
                                }

                                DataTable dtColumnList = new DataTable();
                                dtColumnList.Columns.Add("ColumnName");

                                string strDestinationColumn = objUploadFileMaster.DestinationColumn;
                                string[] arrDestinationColumn = null;

                                if (strDestinationColumn != "")
                                    arrDestinationColumn = strDestinationColumn.Split('|');

                                string strTableName = objUploadFileMaster.TableName;
                                #endregion

                                using (ExcelPackage xlPackage = new ExcelPackage(newFile))
                                {
                                    ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets[objUploadFileMaster.SheetName];

                                    for (int intCol = 2; intCol < 20; intCol++)
                                    {
                                        if (Convert.ToString(worksheet.Cell(3, intCol).Value) != "")
                                        {
                                            DataRow dr = dtColumnList.NewRow();
                                            dr["ColumnName"] = worksheet.Cell(3, intCol).Value;

                                            dtColumnList.Rows.Add(dr);
                                        }
                                        else
                                            break;
                                    }

                                    DataRow drNew;

                                    //string strType = "Call";
                                    // output the data in column 2
                                    for (int iRow = 4; iRow < 50; iRow++)
                                    {
                                        if (Convert.ToString(worksheet.Cell(iRow, 1).Value) != "")
                                        {
                                            for (int intCol = 0; intCol < dtColumnList.Rows.Count; intCol++)
                                            {
                                                drNew = dtData.NewRow();

                                                //DataSet dsResult1 = new DataSet();
                                                //dsResult1 = General.ExecuteDataSet("CHECK_VERSION", Type, intUnderlyingID);

                                                //if (dsResult1 != null && dsResult1.Tables.Count > 0 && dsResult1.Tables[1].Rows.Count > 0 && dsResult1.Tables[1].Rows[0]["VERSION"].ToString() != string.Empty)
                                                //{
                                                //    Version = Convert.ToInt32(dsResult1.Tables[1].Rows[0].ItemArray[0]) + 1;
                                                //}
                                                //else
                                                //{
                                                //    Version = 1;
                                                //}

                                                drNew["VERSION"] = Version;
                                                drNew["FREQUENCY"] = worksheet.Cell(iRow, 1).Value;
                                                drNew["TENURE"] = Convert.ToString(dtColumnList.Rows[intCol][0]);
                                                drNew["VALUE"] = worksheet.Cell(iRow, intCol + 2).Value;
                                                drNew["CREATED_DATE"] = DateTime.Now;
                                                drNew["TYPE"] = Type;
                                                drNew["UNDERLYINGID"] = intUnderlyingID;

                                                dtData.Rows.Add(drNew);
                                            }
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
                                    dsIV = General.ExecuteDataSet(objUploadFileMaster.ExtraProcedure, true);
                                }
                                else
                                {
                                    blnUploadStatus = false;
                                }

                                if (blnUploadStatus)
                                {
                                    ManageUploadFileInfo(intUnderlyingID, RCfile.FileName, strFilePath, blnUploadStatus, blnUploadDataStatus);

                                    ViewBag.Message = "Imported successfully";
                                }

                                objUnderlying = new Underlying();
                                objUnderlying = FetchDefaultDetails();

                                return RedirectToAction("UnderlyingCreationEdit", new { underlyingID = intUnderlyingID, blnUploadStatus = blnUploadStatus });
                            }
                            else
                            {
                                ViewBag.Message = "No File Uploaded";
                                objUnderlying = new Underlying();
                                objUnderlying = FetchDefaultDetails();

                                return RedirectToAction("UnderlyingCreationEdit", new { underlyingID = intUnderlyingID });
                            }
                        }
                        else
                        {
                            ViewBag.Message = "UnderLying Creation Not Created";
                            objUnderlying = new Underlying();
                            objUnderlying = FetchDefaultDetails();

                            return RedirectToAction("UnderlyingCreationEdit", new { underlyingID = intUnderlyingID });
                        }
                    }
                    #endregion

                    #region RCDownload
                    else if (Command == "RCDownload")
                    {
                        Int32 intUnderlyingID = Convert.ToInt32(collection["UnderlyingID"]);

                        UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "RC"; });
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

                            objUnderlying = new Underlying();
                            objUnderlying = FetchDefaultDetails();

                            return RedirectToAction("UnderlyingCreationEdit", new { underlyingID = intUnderlyingID });
                        }
                        else
                            return View();
                    }
                    #endregion

                    #region LVSUpload
                    else if (Command == "LVSUpload")
                    {
                        Int32 intUnderlyingID = Convert.ToInt32(collection["UnderlyingID"]);
                        if (intUnderlyingID != 0)
                        {
                            blnUploadStatus = false;
                            blnUploadDataStatus = true;
                            if (LVSfile != null && LVSfile.ContentLength > 0)
                            {
                                UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "LVS"; });
                                string strFilePath = System.Web.HttpContext.Current.Server.MapPath("~/Uploads/");
                                string strFileName = Path.GetFileNameWithoutExtension(LVSfile.FileName) + DateTime.Now.ToString("dd_mmm_yyyy_hh_mm");
                                string strExtension = Path.GetExtension(LVSfile.FileName);
                                string Type = "";

                                Type = collection.Get("LVS");
                                int Version = 0;

                                strFilePath += strFileName + strExtension;
                                LVSfile.SaveAs(strFilePath);

                                FileInfo newFile = new FileInfo(strFilePath);

                                #region Source and Destination Column
                                string strSourceColumn = objUploadFileMaster.SourceColumn;
                                string[] arrSourceColumn = null;
                                if (strSourceColumn != "")
                                    arrSourceColumn = strSourceColumn.Split('|');

                                DataTable dtData = new DataTable();

                                for (int i = 0; i < arrSourceColumn.Length; i++)
                                {
                                    dtData.Columns.Add(arrSourceColumn[i]);
                                }

                                DataTable dtColumnList = new DataTable();
                                dtColumnList.Columns.Add("ColumnName");

                                string strDestinationColumn = objUploadFileMaster.DestinationColumn;
                                string[] arrDestinationColumn = null;

                                if (strDestinationColumn != "")
                                    arrDestinationColumn = strDestinationColumn.Split('|');

                                string strTableName = objUploadFileMaster.TableName;
                                #endregion

                                using (ExcelPackage xlPackage = new ExcelPackage(newFile))
                                {
                                    ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets[objUploadFileMaster.SheetName];

                                    for (int intCol = 2; intCol < 20; intCol++)
                                    {
                                        if (Convert.ToString(worksheet.Cell(3, intCol).Value) != "")
                                        {
                                            DataRow dr = dtColumnList.NewRow();
                                            dr["ColumnName"] = worksheet.Cell(3, intCol).Value;

                                            dtColumnList.Rows.Add(dr);
                                        }
                                        else
                                            break;
                                    }

                                    DataRow drNew;

                                    //string strType = "Call";
                                    // output the data in column 2
                                    for (int iRow = 4; iRow < 50; iRow++)
                                    {
                                        if (Convert.ToString(worksheet.Cell(iRow, 1).Value) != "")
                                        {
                                            for (int intCol = 0; intCol < dtColumnList.Rows.Count; intCol++)
                                            {
                                                drNew = dtData.NewRow();

                                                //DataSet dsResult1 = new DataSet();
                                                //dsResult1 = General.ExecuteDataSet("CHECK_VERSION", Type, intUnderlyingID);

                                                //if (dsResult1 != null && dsResult1.Tables.Count > 0 && dsResult1.Tables[2].Rows.Count > 0 && dsResult1.Tables[2].Rows[0]["VERSION"].ToString() != string.Empty)
                                                //{
                                                //    Version = Convert.ToInt32(dsResult1.Tables[2].Rows[0].ItemArray[0]) + 1;
                                                //}
                                                //else
                                                //{
                                                //    Version = 1;
                                                //}

                                                drNew["VERSION"] = Version;
                                                drNew["NO_OF_DURATION"] = worksheet.Cell(iRow, 1).Value;
                                                drNew["MONEYNESS"] = Convert.ToDouble(dtColumnList.Rows[intCol][0]);
                                                drNew["VALUE"] = worksheet.Cell(iRow, intCol + 2).Value;
                                                drNew["CREATED_DATE"] = DateTime.Now;
                                                drNew["TYPE"] = Type;
                                                drNew["UNDERLYINGID"] = intUnderlyingID;

                                                dtData.Rows.Add(drNew);
                                            }
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
                                    dsIV = General.ExecuteDataSet(objUploadFileMaster.ExtraProcedure, true);
                                }
                                else
                                {
                                    blnUploadStatus = false;
                                }

                                if (blnUploadStatus)
                                {
                                    ManageUploadFileInfo(intUnderlyingID, LVSfile.FileName, strFilePath, blnUploadStatus, blnUploadDataStatus);

                                    ViewBag.Message = "Imported successfully";
                                }


                                objUnderlying = new Underlying();
                                objUnderlying = FetchDefaultDetails();

                                return RedirectToAction("UnderlyingCreationEdit", new { underlyingID = intUnderlyingID, blnUploadStatus = blnUploadStatus });
                            }
                            else
                            {
                                ViewBag.Message = "No File Uploaded";
                                objUnderlying = new Underlying();
                                objUnderlying = FetchDefaultDetails();

                                return RedirectToAction("UnderlyingCreationEdit", new { underlyingID = intUnderlyingID });
                            }
                        }
                        else
                        {
                            ViewBag.Message = "UnderLying Creation Not Created";
                            objUnderlying = new Underlying();
                            objUnderlying = FetchDefaultDetails();

                            return RedirectToAction("UnderlyingCreationEdit", new { underlyingID = intUnderlyingID });
                        }
                    }
                    #endregion

                    #region LVSDownload
                    else if (Command == "LVSDownload")
                    {
                        Int32 intUnderlyingID = Convert.ToInt32(collection["UnderlyingID"]);

                        UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "LVS"; });
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

                            objUnderlying = new Underlying();
                            objUnderlying = FetchDefaultDetails();

                            return RedirectToAction("UnderlyingCreationEdit", new { underlyingID = intUnderlyingID });
                        }
                        else
                            return View();
                    }
                    #endregion

                    #region CSUpload
                    else if (Command == "CSUpload")
                    {
                        Int32 intUnderlyingID = Convert.ToInt32(collection["UnderlyingID"]);
                        if (intUnderlyingID != 0)
                        {
                            blnUploadStatus = false;
                            blnUploadDataStatus = true;
                            if (CSUfile != null && CSUfile.ContentLength > 0)
                            {

                                string strFilePath = System.Web.HttpContext.Current.Server.MapPath("~/Uploads/");
                                string strFileName = Path.GetFileNameWithoutExtension(CSUfile.FileName) + DateTime.Now.ToString("dd_mmm_yyyy_hh_mm");
                                string strExtension = Path.GetExtension(CSUfile.FileName);
                                string LookupID = "";

                                //var value = ((System.Collections.Specialized.NameValueCollection)(collection)).AllKeys[4];
                                LookupID = collection.Get("CSU");

                                #region Spread Adjustment Surface
                                if (LookupID == "20")
                                {
                                    //Type = collection.Get(value);
                                    UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "CAS"; });
                                    int Version = 0;

                                    strFilePath += strFileName + strExtension;
                                    CSUfile.SaveAs(strFilePath);

                                    FileInfo newFile = new FileInfo(strFilePath);

                                    #region Source and Destination Column
                                    string strSourceColumn = objUploadFileMaster.SourceColumn;
                                    string[] arrSourceColumn = null;
                                    if (strSourceColumn != "")
                                        arrSourceColumn = strSourceColumn.Split('|');

                                    DataTable dtData = new DataTable();

                                    for (int i = 0; i < arrSourceColumn.Length; i++)
                                    {
                                        dtData.Columns.Add(arrSourceColumn[i]);
                                    }

                                    DataTable dtColumnList = new DataTable();
                                    dtColumnList.Columns.Add("ColumnName");

                                    string strDestinationColumn = objUploadFileMaster.DestinationColumn;
                                    string[] arrDestinationColumn = null;

                                    if (strDestinationColumn != "")
                                        arrDestinationColumn = strDestinationColumn.Split('|');

                                    string strTableName = objUploadFileMaster.TableName;
                                    #endregion

                                    using (ExcelPackage xlPackage = new ExcelPackage(newFile))
                                    {
                                        ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets[objUploadFileMaster.SheetName];

                                        for (int intCol = 2; intCol < 20; intCol++)
                                        {
                                            if (Convert.ToString(worksheet.Cell(3, intCol).Value) != "")
                                            {
                                                DataRow dr = dtColumnList.NewRow();
                                                dr["ColumnName"] = worksheet.Cell(3, intCol).Value;

                                                dtColumnList.Rows.Add(dr);
                                            }
                                            else
                                                break;
                                        }

                                        DataRow drNew;

                                        //string strType = "Call";
                                        // output the data in column 2
                                        for (int iRow = 4; iRow < 50; iRow++)
                                        {
                                            if (Convert.ToString(worksheet.Cell(iRow, 1).Value) != "")
                                            {
                                                for (int intCol = 0; intCol < dtColumnList.Rows.Count; intCol++)
                                                {
                                                    drNew = dtData.NewRow();

                                                    //DataSet dsResult1 = new DataSet();
                                                    //dsResult1 = General.ExecuteDataSet("CHECK_VERSION", LookupID, intUnderlyingID);

                                                    //if (dsResult1 != null && dsResult1.Tables.Count > 0 && dsResult1.Tables[3].Rows.Count > 0 && dsResult1.Tables[3].Rows[0]["VERSION"].ToString() != string.Empty)
                                                    //{
                                                    //    Version = Convert.ToInt32(dsResult1.Tables[3].Rows[0].ItemArray[0]) + 1;
                                                    //}
                                                    //else
                                                    //{
                                                    //    Version = 1;
                                                    //}

                                                    drNew["VERSION"] = Version;
                                                    drNew["TENURE"] = worksheet.Cell(iRow, 1).Value;
                                                    drNew["MONEYNESS"] = Convert.ToDouble(dtColumnList.Rows[intCol][0]);
                                                    drNew["VALUE"] = worksheet.Cell(iRow, intCol + 2).Value;
                                                    drNew["CREATED_DATE"] = DateTime.Now;
                                                    drNew["TYPE"] = LookupID;
                                                    drNew["UNDERLYINGID"] = intUnderlyingID;

                                                    dtData.Rows.Add(drNew);
                                                }
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
                                        dsIV = General.ExecuteDataSet(objUploadFileMaster.ExtraProcedure, true);
                                    }
                                    else
                                    {
                                        blnUploadStatus = false;
                                    }

                                    if (blnUploadStatus)
                                    {
                                        ManageUploadFileInfo(intUnderlyingID, CSUfile.FileName, strFilePath, blnUploadStatus, blnUploadDataStatus);

                                        ViewBag.Message = "Imported successfully";
                                    }
                                }
                                #endregion

                                #region Threshhold Strike
                                else if (LookupID == "21")
                                {
                                    UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "CT"; });
                                    int Version = 0;

                                    strFilePath += strFileName + strExtension;
                                    CSUfile.SaveAs(strFilePath);

                                    FileInfo newFile = new FileInfo(strFilePath);

                                    #region Source and Destination Column
                                    string strSourceColumn = objUploadFileMaster.SourceColumn;
                                    string[] arrSourceColumn = null;
                                    if (strSourceColumn != "")
                                        arrSourceColumn = strSourceColumn.Split('|');

                                    DataTable dtData = new DataTable();

                                    for (int i = 0; i < arrSourceColumn.Length; i++)
                                    {
                                        dtData.Columns.Add(arrSourceColumn[i]);
                                    }

                                    DataTable dtColumnList = new DataTable();
                                    dtColumnList.Columns.Add("ColumnName");

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

                                        //string strType = "Call";
                                        // output the data in column 2
                                        for (int iRow = 4; iRow < 36; iRow++)
                                        {
                                            if (Convert.ToString(worksheet.Cell(iRow, 1).Value) != "")
                                            {
                                                drNew = dtData.NewRow();

                                                //DataSet dsResult1 = new DataSet();
                                                //dsResult1 = General.ExecuteDataSet("CHECK_VERSION", LookupID, intUnderlyingID);

                                                //if (dsResult1 != null && dsResult1.Tables.Count > 0 && dsResult1.Tables[4].Rows.Count > 0 && dsResult1.Tables[4].Rows[0]["VERSION"].ToString() != string.Empty)
                                                //{
                                                //    Version = Convert.ToInt32(dsResult1.Tables[4].Rows[0].ItemArray[0]) + 1;
                                                //}
                                                //else
                                                //{
                                                //    Version = 1;
                                                //}

                                                drNew["VERSION"] = Version;
                                                drNew["TENURE"] = worksheet.Cell(iRow, 1).Value;
                                                drNew["STRIKE_CUT_OFF"] = worksheet.Cell(iRow, 2).Value;
                                                drNew["CREATED_DATE"] = DateTime.Now;
                                                drNew["UNDERLYINGID"] = intUnderlyingID;

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
                                        dsIV = General.ExecuteDataSet(objUploadFileMaster.ExtraProcedure, true);
                                    }
                                    else
                                    {
                                        blnUploadStatus = false;
                                    }

                                    if (blnUploadStatus)
                                    {
                                        ManageUploadFileInfo(intUnderlyingID, CSUfile.FileName, strFilePath, blnUploadStatus, blnUploadDataStatus);

                                        ViewBag.Message = "Imported successfully";
                                    }
                                }
                                #endregion

                                #region Spread Minimum IV
                                else if (LookupID == "22")
                                {
                                    UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "CIV"; });
                                    int Version = 0;

                                    strFilePath += strFileName + strExtension;
                                    CSUfile.SaveAs(strFilePath);

                                    FileInfo newFile = new FileInfo(strFilePath);

                                    #region Source and Destination Column
                                    string strSourceColumn = objUploadFileMaster.SourceColumn;
                                    string[] arrSourceColumn = null;
                                    if (strSourceColumn != "")
                                        arrSourceColumn = strSourceColumn.Split('|');

                                    DataTable dtData = new DataTable();

                                    for (int i = 0; i < arrSourceColumn.Length; i++)
                                    {
                                        dtData.Columns.Add(arrSourceColumn[i]);
                                    }

                                    DataTable dtColumnList = new DataTable();
                                    dtColumnList.Columns.Add("ColumnName");

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

                                        //string strType = "Call";
                                        // output the data in column 2
                                        for (int iRow = 4; iRow < 36; iRow++)
                                        {
                                            if (Convert.ToString(worksheet.Cell(iRow, 1).Value) != "")
                                            {
                                                drNew = dtData.NewRow();

                                                //DataSet dsResult1 = new DataSet();
                                                //dsResult1 = General.ExecuteDataSet("CHECK_VERSION", LookupID, intUnderlyingID);

                                                //if (dsResult1 != null && dsResult1.Tables.Count > 0 && dsResult1.Tables[5].Rows.Count > 0 && dsResult1.Tables[4].Rows[0]["VERSION"].ToString() != string.Empty)
                                                //{
                                                //    Version = Convert.ToInt32(dsResult1.Tables[4].Rows[0].ItemArray[0]) + 1;
                                                //}
                                                //else
                                                //{
                                                //    Version = 1;
                                                //}

                                                drNew["VERSION"] = Version;
                                                drNew["TENURE"] = worksheet.Cell(iRow, 1).Value;
                                                drNew["STRIKE_DIFFERENCE"] = worksheet.Cell(iRow, 2).Value;
                                                drNew["MINIMUM_VOL_DIFFERENCE"] = worksheet.Cell(iRow, 3).Value;
                                                drNew["CREATED_DATE"] = DateTime.Now;
                                                drNew["UNDERLYINGID"] = intUnderlyingID;

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
                                        dsIV = General.ExecuteDataSet(objUploadFileMaster.ExtraProcedure, true);
                                    }
                                    else
                                    {
                                        blnUploadStatus = false;
                                    }

                                    if (blnUploadStatus)
                                    {
                                        ManageUploadFileInfo(intUnderlyingID, CSUfile.FileName, strFilePath, blnUploadStatus, blnUploadDataStatus);

                                        ViewBag.Message = "Imported successfully";
                                    }
                                }
                                #endregion

                                objUnderlying = new Underlying();
                                objUnderlying = FetchDefaultDetails();

                                return RedirectToAction("UnderlyingCreationEdit", new { underlyingID = intUnderlyingID, blnUploadStatus = blnUploadStatus });
                            }
                            else
                            {
                                ViewBag.Message = "No File Uploaded";
                                objUnderlying = new Underlying();
                                objUnderlying = FetchDefaultDetails();

                                return RedirectToAction("UnderlyingCreationEdit", new { underlyingID = intUnderlyingID });
                            }
                        }
                        else
                        {
                            ViewBag.Message = "UnderLying Creation Not Created";
                            objUnderlying = new Underlying();
                            objUnderlying = FetchDefaultDetails();

                            return RedirectToAction("UnderlyingCreationEdit", new { underlyingID = intUnderlyingID });
                        }
                    }
                    #endregion

                    #region CSDownload
                    else if (Command == "CSDownload")
                    {
                        Int32 intUnderlyingID = Convert.ToInt32(collection["UnderlyingID"]);
                        string LookupID = "";
                        LookupID = collection.Get("CSU");

                        #region Spread Adjustment Surface
                        if (LookupID == "20")
                        {
                            UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "CAS"; });
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

                                objUnderlying = new Underlying();
                                objUnderlying = FetchDefaultDetails();

                                return RedirectToAction("UnderlyingCreationEdit", new { underlyingID = intUnderlyingID });
                            }
                            else
                                return View();
                        }
                        #endregion

                        #region Call Strike Threshold
                        else if (LookupID == "21")
                        {
                            UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "CT"; });
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

                                objUnderlying = new Underlying();
                                objUnderlying = FetchDefaultDetails();

                                return RedirectToAction("UnderlyingCreationEdit", new { underlyingID = intUnderlyingID });
                            }
                            else
                                return View();
                        }
                        #endregion

                        #region Call Spread Minimum
                        else if (LookupID == "22")
                        {
                            UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "CIV"; });
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

                                objUnderlying = new Underlying();
                                objUnderlying = FetchDefaultDetails();

                                return RedirectToAction("UnderlyingCreationEdit", new { underlyingID = intUnderlyingID });
                            }
                            else
                                return View();
                        }
                        #endregion
                    }
                    #endregion

                    #region PSUpload
                    else if (Command == "PSUpload")
                    {
                        Int32 intUnderlyingID = Convert.ToInt32(collection["UnderlyingID"]);
                        if (intUnderlyingID != 0)
                        {
                            blnUploadStatus = false;
                            blnUploadDataStatus = true;
                            if (PSUfile != null && PSUfile.ContentLength > 0)
                            {

                                string strFilePath = System.Web.HttpContext.Current.Server.MapPath("~/Uploads/");
                                string strFileName = Path.GetFileNameWithoutExtension(PSUfile.FileName) + DateTime.Now.ToString("dd_mmm_yyyy_hh_mm");
                                string strExtension = Path.GetExtension(PSUfile.FileName);
                                string LookupID = "";



                                //var value = ((System.Collections.Specialized.NameValueCollection)(collection)).AllKeys[4];
                                LookupID = collection.Get("PSU");

                                #region Spread Adjustment Surface
                                if (LookupID == "24")
                                {
                                    UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "PAS"; });
                                    int Version = 0;

                                    strFilePath += strFileName + strExtension;
                                    PSUfile.SaveAs(strFilePath);

                                    FileInfo newFile = new FileInfo(strFilePath);

                                    #region Source and Destination Column
                                    string strSourceColumn = objUploadFileMaster.SourceColumn;
                                    string[] arrSourceColumn = null;
                                    if (strSourceColumn != "")
                                        arrSourceColumn = strSourceColumn.Split('|');

                                    DataTable dtData = new DataTable();

                                    for (int i = 0; i < arrSourceColumn.Length; i++)
                                    {
                                        dtData.Columns.Add(arrSourceColumn[i]);
                                    }

                                    DataTable dtColumnList = new DataTable();
                                    dtColumnList.Columns.Add("ColumnName");

                                    string strDestinationColumn = objUploadFileMaster.DestinationColumn;
                                    string[] arrDestinationColumn = null;

                                    if (strDestinationColumn != "")
                                        arrDestinationColumn = strDestinationColumn.Split('|');

                                    string strTableName = objUploadFileMaster.TableName;
                                    #endregion

                                    using (ExcelPackage xlPackage = new ExcelPackage(newFile))
                                    {
                                        ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets[objUploadFileMaster.SheetName];

                                        for (int intCol = 2; intCol < 20; intCol++)
                                        {
                                            if (Convert.ToString(worksheet.Cell(3, intCol).Value) != "")
                                            {
                                                DataRow dr = dtColumnList.NewRow();
                                                dr["ColumnName"] = worksheet.Cell(3, intCol).Value;

                                                dtColumnList.Rows.Add(dr);
                                            }
                                            else
                                                break;
                                        }

                                        DataRow drNew;

                                        //string strType = "Call";
                                        // output the data in column 2
                                        for (int iRow = 4; iRow < 50; iRow++)
                                        {
                                            if (Convert.ToString(worksheet.Cell(iRow, 1).Value) != "")
                                            {
                                                for (int intCol = 0; intCol < dtColumnList.Rows.Count; intCol++)
                                                {
                                                    drNew = dtData.NewRow();

                                                    //DataSet dsResult1 = new DataSet();
                                                    //dsResult1 = General.ExecuteDataSet("CHECK_VERSION", LookupID, intUnderlyingID);

                                                    //if (dsResult1 != null && dsResult1.Tables.Count > 0 && dsResult1.Tables[3].Rows.Count > 0 && dsResult1.Tables[3].Rows[0]["VERSION"].ToString() != string.Empty)
                                                    //{
                                                    //    Version = Convert.ToInt32(dsResult1.Tables[3].Rows[0].ItemArray[0]) + 1;
                                                    //}
                                                    //else
                                                    //{
                                                    //    Version = 1;
                                                    //}

                                                    drNew["VERSION"] = Version;
                                                    drNew["TENURE"] = worksheet.Cell(iRow, 1).Value;
                                                    drNew["MONEYNESS"] = Convert.ToDouble(dtColumnList.Rows[intCol][0]);
                                                    drNew["VALUE"] = worksheet.Cell(iRow, intCol + 2).Value;
                                                    drNew["CREATED_DATE"] = DateTime.Now;
                                                    drNew["TYPE"] = LookupID;
                                                    drNew["UNDERLYINGID"] = intUnderlyingID;

                                                    dtData.Rows.Add(drNew);
                                                }
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
                                        dsIV = General.ExecuteDataSet(objUploadFileMaster.ExtraProcedure, true);
                                    }
                                    else
                                    {
                                        blnUploadStatus = false;
                                    }

                                    if (blnUploadStatus)
                                    {
                                        ManageUploadFileInfo(intUnderlyingID, PSUfile.FileName, strFilePath, blnUploadStatus, blnUploadDataStatus);

                                        ViewBag.Message = "Imported successfully";
                                    }
                                }

                                #endregion

                                #region Spread Skew Adjustment
                                else if (LookupID == "25")
                                {
                                    UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "PS"; });
                                    int Version = 0;

                                    strFilePath += strFileName + strExtension;
                                    PSUfile.SaveAs(strFilePath);

                                    FileInfo newFile = new FileInfo(strFilePath);

                                    #region Source and Destination Column
                                    string strSourceColumn = objUploadFileMaster.SourceColumn;
                                    string[] arrSourceColumn = null;
                                    if (strSourceColumn != "")
                                        arrSourceColumn = strSourceColumn.Split('|');

                                    DataTable dtData = new DataTable();

                                    for (int i = 0; i < arrSourceColumn.Length; i++)
                                    {
                                        dtData.Columns.Add(arrSourceColumn[i]);
                                    }

                                    DataTable dtColumnList = new DataTable();
                                    dtColumnList.Columns.Add("ColumnName");

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

                                        //string strType = "Call";
                                        // output the data in column 2
                                        for (int iRow = 4; iRow < 50; iRow++)
                                        {
                                            if (Convert.ToString(worksheet.Cell(iRow, 1).Value) != "")
                                            {
                                                drNew = dtData.NewRow();

                                                //DataSet dsResult1 = new DataSet();
                                                //dsResult1 = General.ExecuteDataSet("CHECK_VERSION", LookupID, intUnderlyingID);

                                                //if (dsResult1 != null && dsResult1.Tables.Count > 0 && dsResult1.Tables[6].Rows.Count > 0 && dsResult1.Tables[6].Rows[0]["VERSION"].ToString() != string.Empty)
                                                //{
                                                //    Version = Convert.ToInt32(dsResult1.Tables[6].Rows[0].ItemArray[0]) + 1;
                                                //}
                                                //else
                                                //{
                                                //    Version = 1;
                                                //}

                                                drNew["VERSION"] = Version;
                                                drNew["TENURE"] = worksheet.Cell(iRow, 1).Value;
                                                drNew["STRIKE_DIFFERENCE"] = worksheet.Cell(iRow, 2).Value;
                                                drNew["MINIMUM_VOL_DIFFERENCE"] = worksheet.Cell(iRow, 3).Value;
                                                drNew["CREATED_DATE"] = DateTime.Now;
                                                drNew["UNDERLYINGID"] = intUnderlyingID;

                                                dtData.Rows.Add(drNew);
                                            }
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
                                        dsIV = General.ExecuteDataSet(objUploadFileMaster.ExtraProcedure, true);
                                    }
                                    else
                                    {
                                        blnUploadStatus = false;
                                    }

                                    if (blnUploadStatus)
                                    {
                                        ManageUploadFileInfo(intUnderlyingID, PSUfile.FileName, strFilePath, blnUploadStatus, blnUploadDataStatus);

                                        ViewBag.Message = "Imported successfully";
                                    }
                                }
                                #endregion

                                #region Spread Minimum
                                else if (LookupID == "26")
                                {
                                    UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "PIV"; });
                                    int Version = 0;

                                    strFilePath += strFileName + strExtension;
                                    PSUfile.SaveAs(strFilePath);

                                    FileInfo newFile = new FileInfo(strFilePath);

                                    #region Source and Destination Column
                                    string strSourceColumn = objUploadFileMaster.SourceColumn;
                                    string[] arrSourceColumn = null;
                                    if (strSourceColumn != "")
                                        arrSourceColumn = strSourceColumn.Split('|');

                                    DataTable dtData = new DataTable();

                                    for (int i = 0; i < arrSourceColumn.Length; i++)
                                    {
                                        dtData.Columns.Add(arrSourceColumn[i]);
                                    }

                                    DataTable dtColumnList = new DataTable();
                                    dtColumnList.Columns.Add("ColumnName");

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

                                        //string strType = "Call";
                                        // output the data in column 2
                                        for (int iRow = 4; iRow < 36; iRow++)
                                        {
                                            if (Convert.ToString(worksheet.Cell(iRow, 1).Value) != "")
                                            {
                                                drNew = dtData.NewRow();

                                                //DataSet dsResult1 = new DataSet();
                                                //dsResult1 = General.ExecuteDataSet("CHECK_VERSION", LookupID, intUnderlyingID);

                                                //if (dsResult1 != null && dsResult1.Tables.Count > 0 && dsResult1.Tables[7].Rows.Count > 0 && dsResult1.Tables[7].Rows[0]["VERSION"].ToString() != string.Empty)
                                                //{
                                                //    Version = Convert.ToInt32(dsResult1.Tables[7].Rows[0].ItemArray[0]) + 1;
                                                //}
                                                //else
                                                //{
                                                //    Version = 1;
                                                //}

                                                drNew["VERSION"] = Version;
                                                drNew["TENURE"] = worksheet.Cell(iRow, 1).Value;
                                                drNew["STRIKE_DIFFERENCE"] = worksheet.Cell(iRow, 2).Value;
                                                drNew["MINIMUM_VOL_DIFFERENCE"] = worksheet.Cell(iRow, 3).Value;
                                                drNew["CREATED_DATE"] = DateTime.Now;
                                                drNew["UNDERLYINGID"] = intUnderlyingID;

                                                dtData.Rows.Add(drNew);

                                            }
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
                                        dsIV = General.ExecuteDataSet(objUploadFileMaster.ExtraProcedure, true);
                                    }
                                    else
                                    {
                                        blnUploadStatus = false;
                                    }

                                    if (blnUploadStatus)
                                    {
                                        ManageUploadFileInfo(intUnderlyingID, PSUfile.FileName, strFilePath, blnUploadStatus, blnUploadDataStatus);

                                        ViewBag.Message = "Imported successfully";
                                    }
                                }
                                #endregion

                                #region Spread Adjustment Exception
                                if (LookupID == "111")
                                {
                                    UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "PAE"; });
                                    int Version = 0;

                                    strFilePath += strFileName + strExtension;
                                    PSUfile.SaveAs(strFilePath);

                                    FileInfo newFile = new FileInfo(strFilePath);

                                    #region Source and Destination Column
                                    string strSourceColumn = objUploadFileMaster.SourceColumn;
                                    string[] arrSourceColumn = null;
                                    if (strSourceColumn != "")
                                        arrSourceColumn = strSourceColumn.Split('|');

                                    DataTable dtData = new DataTable();

                                    for (int i = 0; i < arrSourceColumn.Length; i++)
                                    {
                                        dtData.Columns.Add(arrSourceColumn[i]);
                                    }

                                    DataTable dtColumnList = new DataTable();
                                    dtColumnList.Columns.Add("ColumnName");

                                    string strDestinationColumn = objUploadFileMaster.DestinationColumn;
                                    string[] arrDestinationColumn = null;

                                    if (strDestinationColumn != "")
                                        arrDestinationColumn = strDestinationColumn.Split('|');

                                    string strTableName = objUploadFileMaster.TableName;
                                    #endregion

                                    using (ExcelPackage xlPackage = new ExcelPackage(newFile))
                                    {
                                        ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets[objUploadFileMaster.SheetName];

                                        for (int intCol = 4; intCol < 20; intCol++)
                                        {
                                            if (Convert.ToString(worksheet.Cell(3, intCol).Value) != "")
                                            {
                                                DataRow dr = dtColumnList.NewRow();
                                                dr["ColumnName"] = worksheet.Cell(3, intCol).Value;

                                                dtColumnList.Rows.Add(dr);
                                            }
                                            else
                                                break;
                                        }

                                        DataRow drNew;

                                        //string strType = "Call";
                                        // output the data in column 2
                                        for (int iRow = 5; iRow < 50; iRow++)
                                        {
                                            if (Convert.ToString(worksheet.Cell(iRow, 1).Value) != "")
                                            {
                                                for (int intCol = 0; intCol < dtColumnList.Rows.Count; intCol++)
                                                {
                                                    drNew = dtData.NewRow();

                                                    drNew["VERSION"] = Version;
                                                    drNew["TENURE"] = Convert.ToDouble(dtColumnList.Rows[intCol][0]);
                                                    drNew["STRIKE"] = worksheet.Cell(iRow, 1).Value;
                                                    drNew["GAP"] = worksheet.Cell(iRow, 2).Value;
                                                    drNew["AVERAGING"] = worksheet.Cell(iRow, 3).Value;
                                                    drNew["VALUE"] = worksheet.Cell(iRow, intCol + 4).Value;
                                                    drNew["UNDERLYING_ID"] = intUnderlyingID;
                                                    drNew["CREATED_BY"] = objUserMaster.UserID;
                                                    drNew["CREATED_ON"] = DateTime.Now;


                                                    dtData.Rows.Add(drNew);
                                                }
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
                                        dsIV = General.ExecuteDataSet(objUploadFileMaster.ExtraProcedure, objUserMaster.UserID, true);
                                    }
                                    else
                                    {
                                        blnUploadStatus = false;
                                    }

                                    if (blnUploadStatus)
                                    {
                                        ManageUploadFileInfo(intUnderlyingID, PSUfile.FileName, strFilePath, blnUploadStatus, blnUploadDataStatus);

                                        ViewBag.Message = "Imported successfully";
                                    }
                                }

                                #endregion


                                objUnderlying = new Underlying();
                                objUnderlying = FetchDefaultDetails();

                                return RedirectToAction("UnderlyingCreationEdit", new { underlyingID = intUnderlyingID, blnUploadStatus = blnUploadStatus });
                            }
                            else
                            {
                                ViewBag.Message = "No File Uploaded";
                                objUnderlying = new Underlying();
                                objUnderlying = FetchDefaultDetails();

                                return RedirectToAction("UnderlyingCreationEdit", new { underlyingID = intUnderlyingID });
                            }
                        }
                        else
                        {
                            ViewBag.Message = "UnderLying Creation Not Created";
                            objUnderlying = new Underlying();
                            objUnderlying = FetchDefaultDetails();

                            return RedirectToAction("UnderlyingCreationEdit", new { underlyingID = intUnderlyingID });
                        }
                    }
                    #endregion

                    #region PSDownload
                    else if (Command == "PSDownload")
                    {
                        Int32 intUnderlyingID = Convert.ToInt32(collection["UnderlyingID"]);
                        string LookupID = "";
                        LookupID = collection.Get("PSU");

                        #region Spread Adjustment Surface
                        if (LookupID == "24")
                        {
                            UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "PAS"; });
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

                                objUnderlying = new Underlying();
                                objUnderlying = FetchDefaultDetails();

                                return RedirectToAction("UnderlyingCreationEdit", new { underlyingID = intUnderlyingID });
                            }
                            else
                                return View();
                        }
                        #endregion

                        #region Spread Skew Adjustment
                        else if (LookupID == "25")
                        {
                            UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "PS"; });
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

                                objUnderlying = new Underlying();
                                objUnderlying = FetchDefaultDetails();

                                return RedirectToAction("UnderlyingCreationEdit", new { underlyingID = intUnderlyingID });
                            }
                            else
                                return View();
                        }
                        #endregion

                        #region Spread Minimum
                        else if (LookupID == "26")
                        {
                            UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "PIV"; });
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

                                objUnderlying = new Underlying();
                                objUnderlying = FetchDefaultDetails();

                                return RedirectToAction("UnderlyingCreationEdit", new { underlyingID = intUnderlyingID });
                            }
                            else
                                return View();
                        }
                        #endregion

                        #region Spread Adjustment Surface
                        if (LookupID == "111")
                        {
                            UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "PAE"; });
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
                            else
                                return View();
                        }
                        #endregion
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
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "UnderlyingCreationEdit Post", objUserMaster.UserID);
                return RedirectToAction("ErrorDetails", "Login");
            }
        }

        #region IVRF
        [HttpGet]
        public ActionResult ImpliedVolatility(string IVRF, string Version)
        {

            LoginController objLoginController = new LoginController();
            ObjectResult<LookupResult> objLookupResult;
            List<LookupResult> LookupResultList;
            Underlying objUnderlying = new Underlying();
            int IVLookUpID = 0;
            int RCLookUpID = 0;
            int LVLookUpID = 0;
            int IVUnderlyingID = 0;
            int RCUnderlyingID = 0;
            int LVUnderlyingID = 0;

            try
            {
                if (ValidateSession())
                {
                    #region Menu Access By on Role

                    Int32 intResult = 0;
                    // bool PPorNonPP = false;

                    UserMaster objUserMaster = new UserMaster();
                    objUserMaster = (UserMaster)Session["LoggedInUser"];

                    var Result = objSP_PRICINGEntities.VALIDATE_MENU_BY_ROLE(objUserMaster.RoleID, "UCIV");
                    intResult = Convert.ToInt32(Result.SingleOrDefault());

                    if (intResult == 0)
                        return RedirectToAction("UserNotAuthorize", "Login");

                    #endregion

                    FetchUploadFileMasterList();

                    #region IV RF List
                    if (IVRF != null)
                    {
                        var List = IVRF.Split('/');
                        IVLookUpID = Convert.ToInt32(List[0]);
                        RCLookUpID = Convert.ToInt32(List[1]);
                        LVLookUpID = Convert.ToInt32(List[2]);
                        IVUnderlyingID = Convert.ToInt32(List[3]);
                        RCUnderlyingID = Convert.ToInt32(List[4]);
                        LVUnderlyingID = Convert.ToInt32(List[5]);
                    }

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

                        objUnderlying.IVRFCategoryList = IVRFList;
                        objUnderlying.RCTypeList = IVRFList;
                        objUnderlying.LVTypeList = IVRFList;

                        //--Set default underlying--Added by Shweta on 3rd May 2016------------START--------------------
                        string strDefaultUnderlyingType = System.Configuration.ConfigurationManager.AppSettings["DefaultUnderlyingType"].ToUpper();
                        LookupMaster objLookupMasterNew = objUnderlying.IVRFCategoryList.Find(delegate(LookupMaster oLookupMaster) { return oLookupMaster.LookupDescription.ToUpper() == strDefaultUnderlyingType; });
                        //--Set default underlying--Added by Shweta on 3rd May 2016------------END----------------------

                        if (IVLookUpID == 0 && RCLookUpID == 0 && LVLookUpID == 0)
                        {
                            objUnderlying.FilterIVRFCategory = objLookupMasterNew.LookupID;
                            objUnderlying.FilterRCType = objLookupMasterNew.LookupID;
                            objUnderlying.FilterLVType = objLookupMasterNew.LookupID;
                        }
                        else
                        {
                            objUnderlying.FilterIVRFCategory = Convert.ToInt32(IVLookUpID);
                            objUnderlying.FilterRCType = Convert.ToInt32(RCLookUpID);
                            objUnderlying.FilterLVType = Convert.ToInt32(LVLookUpID);

                        }
                    }

                    DataSet dsResult = new DataSet();
                    dsResult = General.ExecuteDataSet("GET_UNDERLYING_ID_LIST");

                    List<Underlying> UnderlyingList = new List<Underlying>();

                    if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                    {
                        foreach (DataRow dr in dsResult.Tables[0].Rows)
                        {
                            Underlying objRMMaster = new Underlying();

                            objRMMaster.UnderlyingID = Convert.ToInt32(dr["ID"]);
                            objRMMaster.UnderlyingShortName = Convert.ToString(dr["UNDERLYING_SHORTNAME"]);

                            UnderlyingList.Add(objRMMaster);
                        }

                        objUnderlying.UnderLyingList = UnderlyingList;
                        objUnderlying.RCUnderLyingList = UnderlyingList;
                        objUnderlying.LVUnderLyingList = UnderlyingList;

                        //--Set default underlying--Added by Shweta on 3rd May 2016------------START--------------------
                        string strDefaultUnderlying = System.Configuration.ConfigurationManager.AppSettings["DefaultUnderlying"].ToUpper();
                        Underlying objDefaulyUnderlying = objUnderlying.UnderLyingList.Find(delegate(Underlying oUnderlying) { return oUnderlying.UnderlyingShortName == strDefaultUnderlying; });
                        //--Set default underlying--Added by Shweta on 3rd May 2016------------END----------------------

                        if (IVUnderlyingID == 0 && RCUnderlyingID == 0 && LVUnderlyingID == 0)
                        {
                            objUnderlying.FilterUnderLyingCategory = objDefaulyUnderlying.UnderlyingID;
                            objUnderlying.FilterRCUnderLyingCategory = objDefaulyUnderlying.UnderlyingID;
                            objUnderlying.FilterLVUnderLyingCategory = objDefaulyUnderlying.UnderlyingID;

                        }
                        else
                        {
                            objUnderlying.FilterUnderLyingCategory = Convert.ToInt32(IVUnderlyingID);
                            objUnderlying.FilterRCUnderLyingCategory = Convert.ToInt32(RCUnderlyingID);
                            objUnderlying.FilterLVUnderLyingCategory = Convert.ToInt32(LVUnderlyingID);

                        }
                    }
                    else
                    {
                        TempData["CreateUnderlying"] = "Create Underlying";

                        return RedirectToAction("UnderlyingCreation", "UnderlyingCreation");
                    }


                    #endregion

                    if (IVRF != null)
                        if (Version == null || Version == "0")
                        {
                            ViewBag.Message = "Version is not Mentioned in Excel";
                            return View(objUnderlying);
                        }
                        else if (Version == "-1")
                        {
                            ViewBag.VersionAlready = "Version Not Found";
                            return View(objUnderlying);
                        }
                        else
                        {
                            ViewBag.Successfull = "Successfully Updated";
                            return View(objUnderlying);
                        }

                    return View(objUnderlying);
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
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "ImpliedVolatility Get", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        [HttpPost]
        public ActionResult ImpliedVolatility(Underlying objUnderlying, string Command, FormCollection collection, HttpPostedFileBase file, HttpPostedFileBase RCfile, HttpPostedFileBase LVfile)
        {
            LoginController objLoginController = new LoginController();
            List<UploadFileMaster> UploadFileMasterList = (List<UploadFileMaster>)Session["UploadFileMasterList"];
            // Underlying objUnderlying = new Underlying();

            try
            {
                if (ValidateSession())
                {
                    bool blnUploadStatus = false;
                    bool blnUploadDataStatus = true;

                    #region Implied Volatility
                    if (Command == "IVUpload")
                    {
                        if (file != null && file.ContentLength > 0)
                        {
                            UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "IV"; });
                            string strFilePath = System.Web.HttpContext.Current.Server.MapPath("~/Uploads/");
                            string strFileName = Path.GetFileNameWithoutExtension(file.FileName) + DateTime.Now.ToString("dd_mmm_yyyy_hh_mm");
                            string strExtension = Path.GetExtension(file.FileName);

                            int intIVUnderlyingID = Convert.ToInt32(collection.Get("FilterUnderLyingCategory"));
                            string IVType = collection.Get("FilterIVRFCategory");

                            int intRCUnderlyingID = Convert.ToInt32(collection.Get("FilterRCUnderLyingCategory"));
                            string RCType = collection.Get("FilterRCType");

                            int intLVUnderlyingID = Convert.ToInt32(collection.Get("FilterLVUnderLyingCategory"));
                            string LVType = collection.Get("FilterLVType");

                            strFilePath += strFileName + strExtension;
                            file.SaveAs(strFilePath);

                            FileInfo newFile = new FileInfo(strFilePath);

                            #region Source and Destination Column
                            string strSourceColumn = objUploadFileMaster.SourceColumn;
                            string[] arrSourceColumn = null;
                            if (strSourceColumn != "")
                                arrSourceColumn = strSourceColumn.Split('|');

                            DataTable dtData = new DataTable();

                            for (int i = 0; i < arrSourceColumn.Length; i++)
                            {
                                dtData.Columns.Add(arrSourceColumn[i]);
                            }

                            DataTable dtColumnList = new DataTable();
                            dtColumnList.Columns.Add("ColumnName");

                            string strDestinationColumn = objUploadFileMaster.DestinationColumn;
                            string[] arrDestinationColumn = null;

                            if (strDestinationColumn != "")
                                arrDestinationColumn = strDestinationColumn.Split('|');

                            string strTableName = objUploadFileMaster.TableName;
                            #endregion

                            using (ExcelPackage xlPackage = new ExcelPackage(newFile))
                            {
                                ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets[objUploadFileMaster.SheetName];

                                string strVersion = worksheet.Cell(1, 2).Value.Trim();

                                if (strVersion == "" || strVersion == "0")
                                {
                                    var IVRF1 = IVType + "/" + RCType + "/" + LVType + "/" + intIVUnderlyingID + "/" + intRCUnderlyingID + "/" + intLVUnderlyingID;

                                    return RedirectToAction("ImpliedVolatility", new { IVRF = IVRF1, Version = strVersion });

                                }
                                else
                                {
                                    for (int intCol = 2; intCol < 50; intCol++)
                                    {
                                        if (Convert.ToString(worksheet.Cell(3, intCol).Value) != "")
                                        {
                                            DataRow dr = dtColumnList.NewRow();
                                            dr["ColumnName"] = worksheet.Cell(3, intCol).Value;

                                            dtColumnList.Rows.Add(dr);
                                        }
                                        else
                                            break;
                                    }

                                    DataRow drNew;

                                    //string strType = "Call";
                                    // output the data in column 2
                                    for (int iRow = 4; iRow < 50; iRow++)
                                    {
                                        if (Convert.ToString(worksheet.Cell(iRow, 1).Value) != "")
                                        {
                                            for (int intCol = 0; intCol < dtColumnList.Rows.Count; intCol++)
                                            {
                                                drNew = dtData.NewRow();

                                                drNew["VERSION"] = strVersion;
                                                drNew["TENURE"] = worksheet.Cell(iRow, 1).Value;
                                                drNew["MONEYNESS"] = Convert.ToDouble(dtColumnList.Rows[intCol][0]);
                                                drNew["VALUE"] = worksheet.Cell(iRow, intCol + 2).Value;
                                                drNew["CREATED_DATE"] = DateTime.Now;
                                                drNew["TYPE"] = IVType;
                                                drNew["UNDERLYINGID"] = intIVUnderlyingID;

                                                dtData.Rows.Add(drNew);
                                            }
                                        }
                                        else
                                            break;
                                    }
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
                                dsIV = General.ExecuteDataSet(objUploadFileMaster.ExtraProcedure, false);

                                var result = Convert.ToInt32(dsIV.Tables[0].Rows[0]["Result"]);
                                if (result == -1)
                                {
                                    var IVRF1 = IVType + "/" + RCType + "/" + LVType + "/" + intIVUnderlyingID + "/" + intRCUnderlyingID + "/" + intLVUnderlyingID;

                                    return RedirectToAction("ImpliedVolatility", new { IVRF = IVRF1, Version = -1 });
                                }
                            }
                            else
                            {
                                blnUploadStatus = false;
                            }

                            if (blnUploadStatus)
                            {
                                ManageUploadFileInfo(intIVUnderlyingID, file.FileName, strFilePath, blnUploadStatus, blnUploadDataStatus);

                                ViewBag.Message = "Imported successfully";
                            }

                            objUnderlying = new Underlying();
                            objUnderlying = FetchDefaultDetails();

                            var IVRF = IVType + "/" + RCType + "/" + LVType + "/" + intIVUnderlyingID + "/" + intRCUnderlyingID + "/" + intLVUnderlyingID;

                            return RedirectToAction("ImpliedVolatility", new { IVRF = IVRF, Version = 1 });
                        }
                    }
                    #endregion

                    #region ROll Cost
                    else if (Command == "RCUpload")
                    {
                        if (RCfile != null && RCfile.ContentLength > 0)
                        {
                            UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "RC"; });
                            string strFilePath = System.Web.HttpContext.Current.Server.MapPath("~/Uploads/");
                            string strFileName = Path.GetFileNameWithoutExtension(RCfile.FileName) + DateTime.Now.ToString("dd_mmm_yyyy_hh_mm");
                            string strExtension = Path.GetExtension(RCfile.FileName);

                            int intIVUnderlyingID = Convert.ToInt32(collection.Get("FilterUnderLyingCategory"));
                            string IVType = collection.Get("FilterIVRFCategory");

                            int intRCUnderlyingID = Convert.ToInt32(collection.Get("FilterRCUnderLyingCategory"));
                            string RCType = collection.Get("FilterRCType");

                            int intLVUnderlyingID = Convert.ToInt32(collection.Get("FilterLVUnderLyingCategory"));
                            string LVType = collection.Get("FilterLVType");


                            strFilePath += strFileName + strExtension;
                            RCfile.SaveAs(strFilePath);

                            FileInfo newFile = new FileInfo(strFilePath);

                            #region Source and Destination Column
                            string strSourceColumn = objUploadFileMaster.SourceColumn;
                            string[] arrSourceColumn = null;
                            if (strSourceColumn != "")
                                arrSourceColumn = strSourceColumn.Split('|');

                            DataTable dtData = new DataTable();

                            for (int i = 0; i < arrSourceColumn.Length; i++)
                            {
                                dtData.Columns.Add(arrSourceColumn[i]);
                            }

                            DataTable dtColumnList = new DataTable();
                            dtColumnList.Columns.Add("ColumnName");

                            string strDestinationColumn = objUploadFileMaster.DestinationColumn;
                            string[] arrDestinationColumn = null;

                            if (strDestinationColumn != "")
                                arrDestinationColumn = strDestinationColumn.Split('|');

                            string strTableName = objUploadFileMaster.TableName;
                            #endregion

                            using (ExcelPackage xlPackage = new ExcelPackage(newFile))
                            {
                                ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets[objUploadFileMaster.SheetName];

                                string strVersion = worksheet.Cell(1, 2).Value.Trim();

                                if (strVersion == "" || strVersion == "0")
                                {
                                    var IVRF1 = IVType + "/" + RCType + "/" + LVType + "/" + intIVUnderlyingID + "/" + intRCUnderlyingID + "/" + intLVUnderlyingID;

                                    return RedirectToAction("ImpliedVolatility", new { IVRF = IVRF1, Version = strVersion });

                                }
                                else
                                {
                                    for (int intCol = 2; intCol < 20; intCol++)
                                    {
                                        if (Convert.ToString(worksheet.Cell(3, intCol).Value) != "")
                                        {
                                            DataRow dr = dtColumnList.NewRow();
                                            dr["ColumnName"] = worksheet.Cell(3, intCol).Value;

                                            dtColumnList.Rows.Add(dr);
                                        }
                                        else
                                            break;
                                    }

                                    DataRow drNew;

                                    //string strType = "Call";
                                    // output the data in column 2
                                    for (int iRow = 4; iRow < 36; iRow++)
                                    {
                                        if (Convert.ToString(worksheet.Cell(iRow, 3).Value) != "")
                                        {
                                            for (int intCol = 0; intCol < dtColumnList.Rows.Count; intCol++)
                                            {
                                                drNew = dtData.NewRow();

                                                drNew["VERSION"] = strVersion;
                                                drNew["FREQUENCY"] = worksheet.Cell(iRow, 1).Value;
                                                drNew["TENURE"] = Convert.ToString(dtColumnList.Rows[intCol][0]);
                                                drNew["VALUE"] = worksheet.Cell(iRow, intCol + 2).Value;
                                                drNew["CREATED_DATE"] = DateTime.Now;
                                                drNew["TYPE"] = RCType;
                                                drNew["UNDERLYINGID"] = intRCUnderlyingID;


                                                dtData.Rows.Add(drNew);
                                            }
                                        }
                                        else
                                            break;
                                    }
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
                                dsIV = General.ExecuteDataSet(objUploadFileMaster.ExtraProcedure, false);

                                var result = Convert.ToInt32(dsIV.Tables[0].Rows[0]["Result"]);
                                if (result == -1)
                                {
                                    var IVRF1 = IVType + "/" + RCType + "/" + LVType + "/" + intIVUnderlyingID + "/" + intRCUnderlyingID + "/" + intLVUnderlyingID;

                                    return RedirectToAction("ImpliedVolatility", new { IVRF = IVRF1, Version = -1 });
                                }
                            }
                            else
                            {
                                blnUploadStatus = false;
                            }

                            if (blnUploadStatus)
                            {
                                ManageUploadFileInfo(intRCUnderlyingID, RCfile.FileName, strFilePath, blnUploadStatus, blnUploadDataStatus);

                                ViewBag.Message = "Imported successfully";
                            }

                            objUnderlying = new Underlying();
                            objUnderlying = FetchDefaultDetails();

                            var IVRF = IVType + "/" + RCType + "/" + LVType + "/" + intIVUnderlyingID + "/" + intRCUnderlyingID + "/" + intLVUnderlyingID;

                            return RedirectToAction("ImpliedVolatility", new { IVRF = IVRF, Version = 1 });
                        }
                    }
                    #endregion

                    #region Locale Volatility
                    else if (Command == "LVUpload")
                    {
                        if (LVfile != null && LVfile.ContentLength > 0)
                        {
                            UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "LVS"; });
                            string strFilePath = System.Web.HttpContext.Current.Server.MapPath("~/Uploads/");
                            string strFileName = Path.GetFileNameWithoutExtension(LVfile.FileName) + DateTime.Now.ToString("dd_mmm_yyyy_hh_mm");
                            string strExtension = Path.GetExtension(LVfile.FileName);


                            int intIVUnderlyingID = Convert.ToInt32(collection.Get("FilterUnderLyingCategory"));
                            string IVType = collection.Get("FilterIVRFCategory");

                            int intRCUnderlyingID = Convert.ToInt32(collection.Get("FilterRCUnderLyingCategory"));
                            string RCType = collection.Get("FilterRCType");

                            int intLVUnderlyingID = Convert.ToInt32(collection.Get("FilterLVUnderLyingCategory"));
                            string LVType = collection.Get("FilterLVType");

                            strFilePath += strFileName + strExtension;
                            LVfile.SaveAs(strFilePath);

                            FileInfo newFile = new FileInfo(strFilePath);

                            #region Source and Destination Column
                            string strSourceColumn = objUploadFileMaster.SourceColumn;
                            string[] arrSourceColumn = null;
                            if (strSourceColumn != "")
                                arrSourceColumn = strSourceColumn.Split('|');

                            DataTable dtData = new DataTable();

                            for (int i = 0; i < arrSourceColumn.Length; i++)
                            {
                                dtData.Columns.Add(arrSourceColumn[i]);
                            }

                            DataTable dtColumnList = new DataTable();
                            dtColumnList.Columns.Add("ColumnName");

                            string strDestinationColumn = objUploadFileMaster.DestinationColumn;
                            string[] arrDestinationColumn = null;

                            if (strDestinationColumn != "")
                                arrDestinationColumn = strDestinationColumn.Split('|');

                            string strTableName = objUploadFileMaster.TableName;
                            #endregion

                            using (ExcelPackage xlPackage = new ExcelPackage(newFile))
                            {
                                ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets[objUploadFileMaster.SheetName];

                                string strVersion = worksheet.Cell(1, 2).Value.Trim();

                                if (strVersion == "" || strVersion == "0")
                                {
                                    var IVRF1 = IVType + "/" + RCType + "/" + LVType + "/" + intIVUnderlyingID + "/" + intRCUnderlyingID + "/" + intLVUnderlyingID;

                                    return RedirectToAction("ImpliedVolatility", new { IVRF = IVRF1, Version = strVersion });

                                }
                                else
                                {
                                    for (int intCol = 2; intCol < 20; intCol++)
                                    {
                                        if (Convert.ToString(worksheet.Cell(3, intCol).Value) != "")
                                        {
                                            DataRow dr = dtColumnList.NewRow();
                                            dr["ColumnName"] = worksheet.Cell(3, intCol).Value;

                                            dtColumnList.Rows.Add(dr);
                                        }
                                        else
                                            break;
                                    }

                                    DataRow drNew;

                                    //string strType = "Call";
                                    // output the data in column 2
                                    for (int iRow = 4; iRow < 36; iRow++)
                                    {
                                        if (Convert.ToString(worksheet.Cell(iRow, 1).Value) != "")
                                        {
                                            for (int intCol = 0; intCol < dtColumnList.Rows.Count; intCol++)
                                            {
                                                drNew = dtData.NewRow();



                                                drNew["VERSION"] = strVersion;
                                                drNew["NO_OF_DURATION"] = worksheet.Cell(iRow, 1).Value;
                                                drNew["MONEYNESS"] = Convert.ToDouble(dtColumnList.Rows[intCol][0]);
                                                drNew["VALUE"] = worksheet.Cell(iRow, intCol + 2).Value;
                                                drNew["CREATED_DATE"] = DateTime.Now;
                                                drNew["TYPE"] = LVType;
                                                drNew["UNDERLYINGID"] = intLVUnderlyingID;




                                                dtData.Rows.Add(drNew);
                                            }
                                        }
                                        else
                                            break;
                                    }
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
                                dsIV = General.ExecuteDataSet(objUploadFileMaster.ExtraProcedure, false);

                                var result = Convert.ToInt32(dsIV.Tables[0].Rows[0]["Result"]);
                                if (result == -1)
                                {
                                    var IVRF1 = IVType + "/" + RCType + "/" + LVType + "/" + intIVUnderlyingID + "/" + intRCUnderlyingID + "/" + intLVUnderlyingID;

                                    return RedirectToAction("ImpliedVolatility", new { IVRF = IVRF1, Version = -1 });
                                }
                            }
                            else
                            {
                                blnUploadStatus = false;
                            }

                            if (blnUploadStatus)
                            {
                                ManageUploadFileInfo(intLVUnderlyingID, LVfile.FileName, strFilePath, blnUploadStatus, blnUploadDataStatus);

                                ViewBag.Message = "Imported successfully";
                            }

                            objUnderlying = new Underlying();
                            objUnderlying = FetchDefaultDetails();

                            var IVRF = IVType + "/" + RCType + "/" + LVType + "/" + intIVUnderlyingID + "/" + intRCUnderlyingID + "/" + intLVUnderlyingID;

                            return RedirectToAction("ImpliedVolatility", new { IVRF = IVRF, Version = 1 });
                        }
                    }
                    #endregion

                    #region IVDownload
                    else if (Command == "IVDownload")
                    {
                        UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "IV"; });
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

                        int intIVUnderlyingID = Convert.ToInt32(collection.Get("FilterUnderLyingCategory"));
                        string IVType = collection.Get("FilterIVRFCategory");

                        int intRCUnderlyingID = Convert.ToInt32(collection.Get("FilterRCUnderLyingCategory"));
                        string RCType = collection.Get("FilterRCType");

                        int intLVUnderlyingID = Convert.ToInt32(collection.Get("FilterLVUnderLyingCategory"));
                        string LVType = collection.Get("FilterLVType");

                        var IVRF = IVType + "/" + RCType + "/" + LVType + "/" + intIVUnderlyingID + "/" + intRCUnderlyingID + "/" + intLVUnderlyingID;

                        return RedirectToAction("ImpliedVolatility", new { IVRF = IVRF, Version = 1 });

                    }
                    #endregion

                    #region RCDownload
                    else if (Command == "RCDownload")
                    {
                        UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "RC"; });
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

                        int intIVUnderlyingID = Convert.ToInt32(collection.Get("FilterUnderLyingCategory"));
                        string IVType = collection.Get("FilterIVRFCategory");

                        int intRCUnderlyingID = Convert.ToInt32(collection.Get("FilterRCUnderLyingCategory"));
                        string RCType = collection.Get("FilterRCType");

                        int intLVUnderlyingID = Convert.ToInt32(collection.Get("FilterLVUnderLyingCategory"));
                        string LVType = collection.Get("FilterLVType");

                        var IVRF = IVType + "/" + RCType + "/" + LVType + "/" + intIVUnderlyingID + "/" + intRCUnderlyingID + "/" + intLVUnderlyingID;

                        return RedirectToAction("ImpliedVolatility", new { IVRF = IVRF, Version = 1 });
                    }
                    #endregion

                    #region LVSDownload
                    else if (Command == "LVSDownload")
                    {
                        UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "LVS"; });
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

                        int intIVUnderlyingID = Convert.ToInt32(collection.Get("FilterUnderLyingCategory"));
                        string IVType = collection.Get("FilterIVRFCategory");

                        int intRCUnderlyingID = Convert.ToInt32(collection.Get("FilterRCUnderLyingCategory"));
                        string RCType = collection.Get("FilterRCType");

                        int intLVUnderlyingID = Convert.ToInt32(collection.Get("FilterLVUnderLyingCategory"));
                        string LVType = collection.Get("FilterLVType");

                        var IVRF = IVType + "/" + RCType + "/" + LVType + "/" + intIVUnderlyingID + "/" + intRCUnderlyingID + "/" + intLVUnderlyingID;

                        return RedirectToAction("ImpliedVolatility", new { IVRF = IVRF, Version = 1 });
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
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "ImpliedVolatility Post", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        //public JsonResult FetchImpliedVolatility(int LookUpId, int UnderlyingId)
        //{
        //    try
        //    {
        //        List<ImpliedVolatility> ImpliedVolatilityList = new List<ImpliedVolatility>();


        //        DataSet dsResult1 = new DataSet();
        //        dsResult1 = General.ExecuteDataSet("FETCH_IV", LookUpId, UnderlyingId);

        //        DataSet dsResult = new DataSet();
        //        dsResult = General.ExecuteDataSet("FETCH_IMPLIED_VOLATILITY_BY_LOOKUPID", LookUpId, UnderlyingId);
        //        //List<KeyValuePair<int, string>> lstData = new List<KeyValuePair<int, string>>();
        //        //if (dsResult1 != null && dsResult1.Tables.Count > 0 && dsResult1.Tables[0].Rows.Count > 0)
        //        //{
        //        //    for (int i = 0; i < dsResult1.Tables[0].Rows.Count; i++)
        //        //    {
        //        //        for (int j = 0; j < dsResult1.Tables[0].Columns.Count; j++)
        //        //        {
        //        //            lstData.Add(new KeyValuePair<int, string>(i, dsResult1.Tables[0].Rows[i][j].ToString()));
        //        //            //obj.Version = Convert.ToInt32(dr["VERSION"]);
        //        //            //obj.Tenure = Convert.ToString(dr["TENURE"]);
        //        //            //obj.Moneyness = Convert.ToDouble(dr["MONEYNESS"]);
        //        //            //obj.Value = Convert.ToDouble(dr["VALUE"]);
        //        //            //obj.CreatedDate = Convert.ToDateTime(dr["CREATED_DATE"]);
        //        //            //obj.LookUpCode = Convert.ToString(dr["LOOKUP_DESCRIPTION"]);
        //        //            //obj.UnderlyingShortName = Convert.ToString(dr["UNDERLYING_SHORTNAME"]);
        //        //            //ImpliedVolatilityList.Add(obj);
        //        //        }
        //        //    }
        //        //}



        //        if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
        //        {
        //            foreach (DataRow dr in dsResult.Tables[0].Rows)
        //            {
        //                ImpliedVolatility obj = new ImpliedVolatility();

        //                obj.Version = Convert.ToInt32(dr["VERSION"]);
        //                obj.Tenure = Convert.ToString(dr["TENURE"]);
        //                obj.Moneyness = Convert.ToDouble(dr["MONEYNESS"]);
        //                obj.Value = Convert.ToDouble(dr["VALUE"]);
        //                obj.CreatedDate = Convert.ToDateTime(dr["CREATED_DATE"]);
        //                obj.LookUpCode = Convert.ToString(dr["LOOKUP_DESCRIPTION"]);
        //                obj.UnderlyingShortName = Convert.ToString(dr["UNDERLYING_SHORTNAME"]);
        //                ImpliedVolatilityList.Add(obj);

        //            }
        //        }

        //        var ImpliedVolatilityListData = ImpliedVolatilityList.ToList();
        //        return Json(ImpliedVolatilityListData, JsonRequestBehavior.AllowGet);
        //    }
        //    catch (Exception ex)
        //    {

        //        return Json("");
        //    }
        //}

        public JsonResult FetchImpliedVolatility(int LookUpId, int UnderlyingId)
        {
            try
            {
                List<ImpliedVolatility> ImpliedVolatilityList = new List<ImpliedVolatility>();


                DataSet dsResult1 = new DataSet();
                dsResult1 = General.ExecuteDataSet("FETCH_IV", LookUpId, UnderlyingId);

                string strHTML = "<table border='1'>";
                strHTML += "<tr style='background-color:#BCBCBC;color:Black;'>";

                for (int intCol = 0; intCol < dsResult1.Tables[0].Columns.Count; intCol++)
                {
                    strHTML += "<td>";
                    strHTML += dsResult1.Tables[0].Columns[intCol].Caption;
                    strHTML += "</td>";
                }
                strHTML += "</tr>";

                for (int intRow = 0; intRow < dsResult1.Tables[0].Rows.Count; intRow++)
                {
                    strHTML += "<tr>";
                    for (int intCol = 0; intCol < dsResult1.Tables[0].Columns.Count; intCol++)
                    {
                        strHTML += "<td>";
                        strHTML += dsResult1.Tables[0].Rows[intRow][intCol];
                        strHTML += "</td>";
                    }
                    strHTML += "</tr>";
                }
                strHTML += "</table>";

                if (dsResult1.Tables[0].Rows.Count == 0)
                    strHTML = "No Records Found";

                return Json(strHTML, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "FetchImpliedVolatility", objUserMaster.UserID);
                return Json("");
            }
        }

        public JsonResult FetchRollCost(int LookUpId, int UnderlyingId)
        {
            try
            {
                List<RollCost> RollCostList = new List<RollCost>();


                DataSet dsResult1 = new DataSet();
                dsResult1 = General.ExecuteDataSet("FETCH_ROLL_COST_BY_LOOKUPID", LookUpId, UnderlyingId);

                string strHTML = "<table border='1'>";
                strHTML += "<tr style='background-color:#BCBCBC;color:Black;'>";

                for (int intCol = 0; intCol < dsResult1.Tables[0].Columns.Count; intCol++)
                {
                    strHTML += "<td>";
                    strHTML += dsResult1.Tables[0].Columns[intCol].Caption;
                    strHTML += "</td>";
                }
                strHTML += "</tr>";

                for (int intRow = 0; intRow < dsResult1.Tables[0].Rows.Count; intRow++)
                {
                    strHTML += "<tr>";
                    for (int intCol = 0; intCol < dsResult1.Tables[0].Columns.Count; intCol++)
                    {
                        strHTML += "<td>";
                        strHTML += dsResult1.Tables[0].Rows[intRow][intCol];
                        strHTML += "</td>";
                    }
                    strHTML += "</tr>";
                }
                strHTML += "</table>";

                if (dsResult1.Tables[0].Rows.Count == 0)
                    strHTML = "No Records Found";

                return Json(strHTML, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "FetchRollCost", objUserMaster.UserID);
                return Json("");
            }
        }


        public JsonResult FetchLVSurface(int LookUpId, int UnderlyingId)
        {
            try
            {
                List<LocaleVolatilitySurface> LocaleVolatilityList = new List<LocaleVolatilitySurface>();


                //DataSet dsResult = new DataSet();
                //dsResult = General.ExecuteDataSet("FETCH_LOCALE_VOLATILITY_BY_LOOKUPID", LookUpId, UnderlyingId);

                //if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                //{
                //    foreach (DataRow dr in dsResult.Tables[0].Rows)
                //    {
                //        LocaleVolatilitySurface obj = new LocaleVolatilitySurface();

                //        obj.Version = Convert.ToInt32(dr["VERSION"]);
                //        obj.Tenure = Convert.ToString(dr["NO_OF_DURATION"]);
                //        obj.Moneyness = Convert.ToDouble(dr["MONEYNESS"]);
                //        obj.Value = Convert.ToDouble(dr["VALUE"]);
                //        obj.CreatedDate = Convert.ToDateTime(dr["CREATED_DATE"]);
                //        obj.LookUpCode = Convert.ToString(dr["LOOKUP_DESCRIPTION"]);
                //        obj.UnderlyingShortName = Convert.ToString(dr["UNDERLYING_SHORTNAME"]);

                //        LocaleVolatilityList.Add(obj);
                //    }
                //}

                //var LocaleVolatilityListData = LocaleVolatilityList.ToList();

                // return Json(LocaleVolatilityListData, JsonRequestBehavior.AllowGet);

                DataSet dsResult1 = new DataSet();
                dsResult1 = General.ExecuteDataSet("FETCH_LV", LookUpId, UnderlyingId);

                string strHTML = "";

                if (dsResult1 != null && dsResult1.Tables.Count > 0)
                {
                    if (dsResult1.Tables[0].Rows.Count > 0)
                    {
                        strHTML = "<table border='1'>";
                        strHTML += "<tr style='background-color:#BCBCBC;color:Black;'>";

                        for (int intCol = 0; intCol < dsResult1.Tables[0].Columns.Count; intCol++)
                        {
                            strHTML += "<td>";
                            strHTML += dsResult1.Tables[0].Columns[intCol].Caption;
                            strHTML += "</td>";
                        }
                        strHTML += "</tr>";

                        for (int intRow = 0; intRow < dsResult1.Tables[0].Rows.Count; intRow++)
                        {
                            strHTML += "<tr>";
                            for (int intCol = 0; intCol < dsResult1.Tables[0].Columns.Count; intCol++)
                            {
                                strHTML += "<td>";
                                strHTML += dsResult1.Tables[0].Rows[intRow][intCol];
                                strHTML += "</td>";
                            }
                            strHTML += "</tr>";
                        }
                        strHTML += "</table>";
                    }
                    else if (dsResult1.Tables[0].Rows.Count == 0)
                        strHTML = "No Records Found";
                }

                return Json(strHTML, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "FetchLVSurface", objUserMaster.UserID);
                return Json("");
            }
        }
        #endregion

        #region Call Spread


        [HttpGet]
        public ActionResult AdjustmentSurface(string IVRF, string Version)
        {
            LoginController objLoginController = new LoginController();
            try
            {
                if (ValidateSession())
                {

                    Underlying objUnderlying = new Underlying();

                    #region Menu Access By on Role

                    Int32 intResult = 0;
                    // bool PPorNonPP = false;

                    UserMaster objUserMaster = new UserMaster();
                    objUserMaster = (UserMaster)Session["LoggedInUser"];

                    var Result = objSP_PRICINGEntities.VALIDATE_MENU_BY_ROLE(objUserMaster.RoleID, "UCAS");
                    intResult = Convert.ToInt32(Result.SingleOrDefault());

                    if (intResult == 0)
                        return RedirectToAction("UserNotAuthorize", "Login");

                    #endregion

                    int CSUAUnderlyingID = 0;
                    int CSUTUnderlyingID = 0;
                    int CSUMUnderlyingID = 0;

                    FetchUploadFileMasterList();

                    #region IVRF
                    if (IVRF != null)
                    {
                        var List = IVRF.Split('/');

                        CSUAUnderlyingID = Convert.ToInt32(List[0]);
                        CSUTUnderlyingID = Convert.ToInt32(List[1]);
                        CSUMUnderlyingID = Convert.ToInt32(List[2]);
                    }

                    DataSet dsResult = new DataSet();
                    dsResult = General.ExecuteDataSet("GET_UNDERLYING_ID_LIST");

                    List<Underlying> UnderlyingList = new List<Underlying>();

                    if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                    {
                        foreach (DataRow dr in dsResult.Tables[0].Rows)
                        {
                            Underlying obj = new Underlying();

                            obj.UnderlyingID = Convert.ToInt32(dr["ID"]);
                            obj.UnderlyingShortName = Convert.ToString(dr["UNDERLYING_SHORTNAME"]);

                            UnderlyingList.Add(obj);
                        }

                        objUnderlying.CSUAUnderLyingList = UnderlyingList;
                        objUnderlying.CSUTUnderLyingList = UnderlyingList;
                        objUnderlying.CSUMinimumUnderLyingList = UnderlyingList;

                        //--Set default underlying--Added by Shweta on 3rd May 2016------------START--------------------
                        string strDefaultUnderlying = System.Configuration.ConfigurationManager.AppSettings["DefaultUnderlying"].ToUpper();
                        Underlying objDefaulyUnderlying = objUnderlying.CSUAUnderLyingList.Find(delegate(Underlying oUnderlying) { return oUnderlying.UnderlyingShortName == strDefaultUnderlying; });
                        //--Set default underlying--Added by Shweta on 3rd May 2016------------END----------------------

                        if (CSUAUnderlyingID == 0 && CSUTUnderlyingID == 0 && CSUMUnderlyingID == 0)
                        {
                            objUnderlying.FilterCSUAdjustment = objDefaulyUnderlying.UnderlyingID;
                            objUnderlying.FilterCSUThreshold = objDefaulyUnderlying.UnderlyingID;
                            objUnderlying.FilterCSUMinimum = objDefaulyUnderlying.UnderlyingID;

                        }
                        else
                        {
                            objUnderlying.FilterCSUAdjustment = Convert.ToInt32(CSUAUnderlyingID);
                            objUnderlying.FilterCSUThreshold = Convert.ToInt32(CSUTUnderlyingID);
                            objUnderlying.FilterCSUMinimum = Convert.ToInt32(CSUMUnderlyingID);

                        }
                    }
                    else
                    {
                        TempData["CreateUnderlying"] = "Create Underlying";

                        return RedirectToAction("UnderlyingCreation", "UnderlyingCreation");
                    }

                    #endregion

                    if (IVRF != null)
                        if (Version == null || Version == "0")
                        {
                            ViewBag.Message = "Version is not Mentioned in Excel";
                            return View(objUnderlying);
                        }
                        else if (Version == "-1")
                        {
                            ViewBag.VersionAlready = "Version Not Found";
                            return View(objUnderlying);
                        }
                        else
                        {
                            ViewBag.Successfull = "Successfully Updated";
                            return View(objUnderlying);
                        }

                    return View(objUnderlying);
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
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "AdjustmentSurface Get", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        [HttpPost]
        public ActionResult AdjustmentSurface(Underlying objUnderlying, string Command, FormCollection collection, HttpPostedFileBase file, HttpPostedFileBase Threshold, HttpPostedFileBase CallMinimum)
        {

            List<UploadFileMaster> UploadFileMasterList = (List<UploadFileMaster>)Session["UploadFileMasterList"];
            // Underlying objUnderlying = new Underlying();
            LoginController objLoginController = new LoginController();
            try
            {
                if (ValidateSession())
                {
                    bool blnUploadStatus = false;
                    bool blnUploadDataStatus = true;

                    #region Call Adjustment
                    if (Command == "AdjustmentUpload")
                    {
                        if (file != null && file.ContentLength > 0)
                        {
                            UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "CAS"; });
                            string strFilePath = System.Web.HttpContext.Current.Server.MapPath("~/Uploads/");
                            string strFileName = Path.GetFileNameWithoutExtension(file.FileName) + DateTime.Now.ToString("dd_mmm_yyyy_hh_mm");
                            string strExtension = Path.GetExtension(file.FileName);
                            string Type = "";

                            int intAdjustmentUnderlyingID = Convert.ToInt32(collection.Get("FilterCSUAdjustment"));
                            Type = "20";

                            int intThresholdUnderlyingID = Convert.ToInt32(collection.Get("FilterCSUThreshold"));

                            int intMinimumUnderlyingID = Convert.ToInt32(collection.Get("FilterCSUMinimum"));

                            strFilePath += strFileName + strExtension;
                            file.SaveAs(strFilePath);

                            FileInfo newFile = new FileInfo(strFilePath);

                            #region Source and Destination Column
                            string strSourceColumn = objUploadFileMaster.SourceColumn;
                            string[] arrSourceColumn = null;
                            if (strSourceColumn != "")
                                arrSourceColumn = strSourceColumn.Split('|');

                            DataTable dtData = new DataTable();

                            for (int i = 0; i < arrSourceColumn.Length; i++)
                            {
                                dtData.Columns.Add(arrSourceColumn[i]);
                            }

                            DataTable dtColumnList = new DataTable();
                            dtColumnList.Columns.Add("ColumnName");

                            string strDestinationColumn = objUploadFileMaster.DestinationColumn;
                            string[] arrDestinationColumn = null;

                            if (strDestinationColumn != "")
                                arrDestinationColumn = strDestinationColumn.Split('|');

                            string strTableName = objUploadFileMaster.TableName;
                            #endregion

                            using (ExcelPackage xlPackage = new ExcelPackage(newFile))
                            {
                                ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets[objUploadFileMaster.SheetName];

                                string strVersion = worksheet.Cell(1, 2).Value.Trim();

                                if (strVersion == "" || strVersion == "0")
                                {
                                    var IVRF1 = intAdjustmentUnderlyingID + "/" + intThresholdUnderlyingID + "/" + intMinimumUnderlyingID;

                                    return RedirectToAction("AdjustmentSurface", new { IVRF = IVRF1, Version = strVersion });
                                }
                                else
                                {
                                    for (int intCol = 2; intCol < 20; intCol++)
                                    {
                                        if (Convert.ToString(worksheet.Cell(3, intCol).Value) != "")
                                        {
                                            DataRow dr = dtColumnList.NewRow();
                                            dr["ColumnName"] = worksheet.Cell(3, intCol).Value;

                                            dtColumnList.Rows.Add(dr);
                                        }
                                        else
                                            break;
                                    }

                                    DataRow drNew;

                                    //string strType = "Call";
                                    // output the data in column 2
                                    for (int iRow = 4; iRow < 50; iRow++)
                                    {
                                        if (Convert.ToString(worksheet.Cell(iRow, 1).Value) != "")
                                        {
                                            for (int intCol = 0; intCol < dtColumnList.Rows.Count; intCol++)
                                            {
                                                drNew = dtData.NewRow();



                                                drNew["VERSION"] = strVersion;
                                                drNew["TENURE"] = worksheet.Cell(iRow, 1).Value;
                                                drNew["MONEYNESS"] = Convert.ToDouble(dtColumnList.Rows[intCol][0]);
                                                drNew["VALUE"] = worksheet.Cell(iRow, intCol + 2).Value;
                                                drNew["CREATED_DATE"] = DateTime.Now;
                                                drNew["TYPE"] = Type;
                                                drNew["UNDERLYINGID"] = intAdjustmentUnderlyingID;




                                                dtData.Rows.Add(drNew);
                                            }
                                        }
                                        else
                                            break;
                                    }
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
                                dsIV = General.ExecuteDataSet(objUploadFileMaster.ExtraProcedure, false);

                                var result = Convert.ToInt32(dsIV.Tables[0].Rows[0]["Result"]);
                                if (result == -1)
                                {
                                    var IVRF1 = intAdjustmentUnderlyingID + "/" + intThresholdUnderlyingID + "/" + intMinimumUnderlyingID;

                                    return RedirectToAction("AdjustmentSurface", new { IVRF = IVRF1, Version = -1 });
                                }
                            }
                            else
                            {
                                blnUploadStatus = false;
                            }

                            if (blnUploadStatus)
                            {
                                ManageUploadFileInfo(intAdjustmentUnderlyingID, file.FileName, strFilePath, blnUploadStatus, blnUploadDataStatus);

                                ViewBag.Message = "Imported successfully";
                            }

                            objUnderlying = new Underlying();
                            objUnderlying = FetchDefaultDetails();

                            var IVRF = intAdjustmentUnderlyingID + "/" + intThresholdUnderlyingID + "/" + intMinimumUnderlyingID;

                            return RedirectToAction("AdjustmentSurface", new { IVRF = IVRF, Version = 1 });
                        }
                    }
                    #endregion

                    #region Call Threshold
                    else if (Command == "ThresholdUpload")
                    {
                        if (Threshold != null && Threshold.ContentLength > 0)
                        {
                            UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "CT"; });
                            string strFilePath = System.Web.HttpContext.Current.Server.MapPath("~/Uploads/");
                            string strFileName = Path.GetFileNameWithoutExtension(Threshold.FileName) + DateTime.Now.ToString("dd_mmm_yyyy_hh_mm");
                            string strExtension = Path.GetExtension(Threshold.FileName);

                            int intAdjustmentUnderlyingID = Convert.ToInt32(collection.Get("FilterCSUAdjustment"));

                            int intThresholdUnderlyingID = Convert.ToInt32(collection.Get("FilterCSUThreshold"));

                            int intMinimumUnderlyingID = Convert.ToInt32(collection.Get("FilterCSUMinimum"));

                            strFilePath += strFileName + strExtension;
                            Threshold.SaveAs(strFilePath);

                            FileInfo newFile = new FileInfo(strFilePath);

                            #region Source and Destination Column
                            string strSourceColumn = objUploadFileMaster.SourceColumn;
                            string[] arrSourceColumn = null;
                            if (strSourceColumn != "")
                                arrSourceColumn = strSourceColumn.Split('|');

                            DataTable dtData = new DataTable();

                            for (int i = 0; i < arrSourceColumn.Length; i++)
                            {
                                dtData.Columns.Add(arrSourceColumn[i]);
                            }

                            DataTable dtColumnList = new DataTable();
                            dtColumnList.Columns.Add("ColumnName");

                            string strDestinationColumn = objUploadFileMaster.DestinationColumn;
                            string[] arrDestinationColumn = null;

                            if (strDestinationColumn != "")
                                arrDestinationColumn = strDestinationColumn.Split('|');

                            string strTableName = objUploadFileMaster.TableName;
                            #endregion


                            using (ExcelPackage xlPackage = new ExcelPackage(newFile))
                            {
                                ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets[objUploadFileMaster.SheetName];

                                string strVersion = worksheet.Cell(1, 2).Value.Trim();

                                if (strVersion == "" || strVersion == "0")
                                {
                                    var IVRF1 = intAdjustmentUnderlyingID + "/" + intThresholdUnderlyingID + "/" + intMinimumUnderlyingID;

                                    return RedirectToAction("AdjustmentSurface", new { IVRF = IVRF1, Version = strVersion });
                                }
                                else
                                {
                                    DataRow drNew;

                                    //string strType = "Call";
                                    // output the data in column 2
                                    for (int iRow = 4; iRow < 50; iRow++)
                                    {
                                        if (Convert.ToString(worksheet.Cell(iRow, 1).Value) != "")
                                        {
                                            drNew = dtData.NewRow();

                                            drNew["VERSION"] = strVersion;
                                            drNew["TENURE"] = worksheet.Cell(iRow, 1).Value;
                                            drNew["STRIKE_CUT_OFF"] = worksheet.Cell(iRow, 2).Value;
                                            drNew["CREATED_DATE"] = DateTime.Now;
                                            drNew["UNDERLYINGID"] = intThresholdUnderlyingID;

                                            dtData.Rows.Add(drNew);

                                        }
                                        else
                                            break;
                                    }
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
                                dsIV = General.ExecuteDataSet(objUploadFileMaster.ExtraProcedure, false);

                                var result = Convert.ToInt32(dsIV.Tables[0].Rows[0]["Result"]);
                                if (result == -1)
                                {
                                    var IVRF1 = intAdjustmentUnderlyingID + "/" + intThresholdUnderlyingID + "/" + intMinimumUnderlyingID;

                                    return RedirectToAction("AdjustmentSurface", new { IVRF = IVRF1, Version = -1 });
                                }
                            }
                            else
                            {
                                blnUploadStatus = false;
                            }

                            if (blnUploadStatus)
                            {
                                ManageUploadFileInfo(intThresholdUnderlyingID, Threshold.FileName, strFilePath, blnUploadStatus, blnUploadDataStatus);

                                ViewBag.Message = "Imported successfully";
                            }
                            objUnderlying = new Underlying();
                            objUnderlying = FetchDefaultDetails();

                            var IVRF = intAdjustmentUnderlyingID + "/" + intThresholdUnderlyingID + "/" + intMinimumUnderlyingID;

                            return RedirectToAction("AdjustmentSurface", new { IVRF = IVRF, Version = 1 });
                        }
                    }
                    #endregion

                    #region Call Minimum
                    else if (Command == "CallMinimumUpload")
                    {
                        if (CallMinimum != null && CallMinimum.ContentLength > 0)
                        {
                            UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "CIV"; });
                            string strFilePath = System.Web.HttpContext.Current.Server.MapPath("~/Uploads/");
                            string strFileName = Path.GetFileNameWithoutExtension(CallMinimum.FileName) + DateTime.Now.ToString("dd_mmm_yyyy_hh_mm");
                            string strExtension = Path.GetExtension(CallMinimum.FileName);

                            int intAdjustmentUnderlyingID = Convert.ToInt32(collection.Get("FilterCSUAdjustment"));

                            int intThresholdUnderlyingID = Convert.ToInt32(collection.Get("FilterCSUThreshold"));

                            int intMinimumUnderlyingID = Convert.ToInt32(collection.Get("FilterCSUMinimum"));

                            strFilePath += strFileName + strExtension;
                            CallMinimum.SaveAs(strFilePath);

                            FileInfo newFile = new FileInfo(strFilePath);

                            #region Source and Destination Column
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
                            #endregion


                            using (ExcelPackage xlPackage = new ExcelPackage(newFile))
                            {
                                ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets[objUploadFileMaster.SheetName];

                                string strVersion = worksheet.Cell(1, 2).Value.Trim();

                                if (strVersion == "" || strVersion == "0")
                                {
                                    var IVRF1 = intAdjustmentUnderlyingID + "/" + intThresholdUnderlyingID + "/" + intMinimumUnderlyingID;

                                    return RedirectToAction("AdjustmentSurface", new { IVRF = IVRF1, Version = strVersion });
                                }
                                else
                                {
                                    DataRow drNew;

                                    //string strType = "Call";
                                    // output the data in column 2
                                    for (int iRow = 4; iRow < 36; iRow++)
                                    {
                                        if (Convert.ToString(worksheet.Cell(iRow, 1).Value) != "")
                                        {
                                            drNew = dtData.NewRow();

                                            drNew["VERSION"] = strVersion;
                                            drNew["TENURE"] = worksheet.Cell(iRow, 1).Value;
                                            drNew["STRIKE_DIFFERENCE"] = worksheet.Cell(iRow, 2).Value;
                                            drNew["MINIMUM_VOL_DIFFERENCE"] = worksheet.Cell(iRow, 3).Value;
                                            drNew["CREATED_DATE"] = DateTime.Now;
                                            drNew["UNDERLYINGID"] = intMinimumUnderlyingID;

                                            dtData.Rows.Add(drNew);

                                        }
                                    }
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
                                dsIV = General.ExecuteDataSet(objUploadFileMaster.ExtraProcedure, false);

                                var result = Convert.ToInt32(dsIV.Tables[0].Rows[0]["Result"]);
                                if (result == -1)
                                {
                                    var IVRF1 = intAdjustmentUnderlyingID + "/" + intThresholdUnderlyingID + "/" + intMinimumUnderlyingID;

                                    return RedirectToAction("AdjustmentSurface", new { IVRF = IVRF1, Version = -1 });
                                }
                            }
                            else
                            {
                                blnUploadStatus = false;
                            }

                            if (blnUploadStatus)
                            {
                                ManageUploadFileInfo(intMinimumUnderlyingID, CallMinimum.FileName, strFilePath, blnUploadStatus, blnUploadDataStatus);

                                ViewBag.Message = "Imported successfully";
                            }
                            objUnderlying = new Underlying();
                            objUnderlying = FetchDefaultDetails();

                            var IVRF = intAdjustmentUnderlyingID + "/" + intThresholdUnderlyingID + "/" + intMinimumUnderlyingID;

                            return RedirectToAction("AdjustmentSurface", new { IVRF = IVRF, Version = 1 });
                        }
                    }
                    #endregion

                    #region Spread Adjustment Surface
                    else if (Command == "ASDownload")
                    {
                        UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "CAS"; });
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

                        int intAdjustmentUnderlyingID = Convert.ToInt32(collection.Get("FilterCSUAdjustment"));
                        int intThresholdUnderlyingID = Convert.ToInt32(collection.Get("FilterCSUThreshold"));
                        int intMinimumUnderlyingID = Convert.ToInt32(collection.Get("FilterCSUMinimum"));
                        var IVRF = intAdjustmentUnderlyingID + "/" + intThresholdUnderlyingID + "/" + intMinimumUnderlyingID;

                        return RedirectToAction("AdjustmentSurface", new { IVRF = IVRF, Version = 1 });
                    }
                    #endregion

                    #region Call Strike Threshold
                    else if (Command == "CTDownload")
                    {
                        UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "CT"; });
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

                        int intAdjustmentUnderlyingID = Convert.ToInt32(collection.Get("FilterCSUAdjustment"));
                        int intThresholdUnderlyingID = Convert.ToInt32(collection.Get("FilterCSUThreshold"));
                        int intMinimumUnderlyingID = Convert.ToInt32(collection.Get("FilterCSUMinimum"));
                        var IVRF = intAdjustmentUnderlyingID + "/" + intThresholdUnderlyingID + "/" + intMinimumUnderlyingID;

                        return RedirectToAction("AdjustmentSurface", new { IVRF = IVRF, Version = 1 });
                    }
                    #endregion

                    #region Call Spread Minimum
                    else if (Command == "CMDownload")
                    {
                        UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "CIV"; });
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

                        int intAdjustmentUnderlyingID = Convert.ToInt32(collection.Get("FilterCSUAdjustment"));
                        int intThresholdUnderlyingID = Convert.ToInt32(collection.Get("FilterCSUThreshold"));
                        int intMinimumUnderlyingID = Convert.ToInt32(collection.Get("FilterCSUMinimum"));
                        var IVRF = intAdjustmentUnderlyingID + "/" + intThresholdUnderlyingID + "/" + intMinimumUnderlyingID;

                        return RedirectToAction("AdjustmentSurface", new { IVRF = IVRF, Version = 1 });
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
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "AdjustmentSurface Post", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        public JsonResult FetchCallAdjusment(int UnderlyingId)
        {
            try
            {
                if (UnderlyingId == 0)
                {
                    UnderlyingId = Convert.ToInt32(Session["UnderlyingID"]);
                }
                List<Call_Put_AdjustmentSurface> AdjustmentSurfaceList = new List<Call_Put_AdjustmentSurface>();

                var lookupCode = "CSA";
                DataSet dsResult1 = new DataSet();
                dsResult1 = General.ExecuteDataSet("FETCH_CALL_ADJUSMENT_SURFACE", lookupCode, UnderlyingId);


                string strHTML = "<table border='1'>";
                strHTML += "<tr style='background-color:#BCBCBC;color:Black;'>";

                for (int intCol = 0; intCol < dsResult1.Tables[0].Columns.Count; intCol++)
                {
                    strHTML += "<td>";
                    strHTML += dsResult1.Tables[0].Columns[intCol].Caption;
                    strHTML += "</td>";
                }
                strHTML += "</tr>";

                for (int intRow = 0; intRow < dsResult1.Tables[0].Rows.Count; intRow++)
                {
                    strHTML += "<tr>";
                    for (int intCol = 0; intCol < dsResult1.Tables[0].Columns.Count; intCol++)
                    {
                        strHTML += "<td>";
                        strHTML += dsResult1.Tables[0].Rows[intRow][intCol];
                        strHTML += "</td>";
                    }
                    strHTML += "</tr>";
                }
                strHTML += "</table>";

                if (dsResult1.Tables[0].Rows.Count == 0)
                    strHTML = "No Records Found";

                return Json(strHTML, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "FetchCallAdjusment", objUserMaster.UserID);
                return Json("");
            }
        }


        public JsonResult FetchCallThresholdStrike(int UnderlyingId)
        {
            try
            {
                if (UnderlyingId == 0)
                {
                    UnderlyingId = Convert.ToInt32(Session["UnderlyingID"]);
                }
                List<CallThresholdStrike> ThresholdStrikeList = new List<CallThresholdStrike>();


                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("FETCH_CALL_THRESHOLD_STRIKE", UnderlyingId);

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in dsResult.Tables[0].Rows)
                    {
                        CallThresholdStrike obj = new CallThresholdStrike();

                        obj.Version = Convert.ToInt32(dr["VERSION"]);
                        obj.Tenure = Convert.ToString(dr["TENURE"]);
                        obj.Strikecutoff = Convert.ToDouble(dr["STRIKE_CUT_OFF"]);
                        obj.CreatedDate = Convert.ToDateTime(dr["CREATED_DATE"]);
                        obj.UnderlyingShortName = Convert.ToString(dr["UNDERLYING_SHORTNAME"]);

                        ThresholdStrikeList.Add(obj);
                    }
                }

                var ThresholdStrikeListData = ThresholdStrikeList.ToList();
                return Json(ThresholdStrikeListData, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "FetchCallThresholdStrike", objUserMaster.UserID);
                return Json("");
            }
        }

        public JsonResult FetchCallMinimum(int UnderlyingId)
        {
            try
            {
                if (UnderlyingId == 0)
                {
                    UnderlyingId = Convert.ToInt32(Session["UnderlyingID"]);
                }
                List<CallMinimumGap> CallMinimumGapList = new List<CallMinimumGap>();


                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("FETCH_CALL_MINIMUM_GAP", UnderlyingId);

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in dsResult.Tables[0].Rows)
                    {
                        CallMinimumGap obj = new CallMinimumGap();

                        obj.Version = Convert.ToInt32(dr["VERSION"]);
                        obj.Tenure = Convert.ToString(dr["TENURE"]);
                        obj.StrikeDifference = Convert.ToDouble(dr["STRIKE_DIFFERENCE"]);
                        obj.MinimumVolDifference = Convert.ToDouble(dr["MINIMUM_VOL_DIFFERENCE"]);
                        obj.CreatedDate = Convert.ToDateTime(dr["CREATED_DATE"]);
                        obj.UnderlyingShortName = Convert.ToString(dr["UNDERLYING_SHORTNAME"]);

                        CallMinimumGapList.Add(obj);
                    }
                }

                var CallMinimumGapListData = CallMinimumGapList.ToList();
                return Json(CallMinimumGapListData, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "FetchCallMinimum", objUserMaster.UserID);
                return Json("");
            }
        }

        #endregion

        #region Put Spread

        [HttpGet]
        public ActionResult PutAdjustmentSurface(string IVRF, string Version)
        {
            try
            {
                if (ValidateSession())
                {
                    LoginController objLoginController = new LoginController();
                    Underlying objUnderlying = new Underlying();

                    #region Menu Access By on Role

                    Int32 intResult = 0;
                    // bool PPorNonPP = false;

                    UserMaster objUserMaster = new UserMaster();
                    objUserMaster = (UserMaster)Session["LoggedInUser"];

                    var Result = objSP_PRICINGEntities.VALIDATE_MENU_BY_ROLE(objUserMaster.RoleID, "UPAS");
                    intResult = Convert.ToInt32(Result.SingleOrDefault());

                    if (intResult == 0)
                        return RedirectToAction("UserNotAuthorize", "Login");

                    #endregion

                    int PSUAUnderlyingID = 0;
                    int PSUSkewUnderlyingID = 0;
                    int PSUMUnderlyingID = 0;
                    int PSUEUnderlyingID = 0;

                    FetchUploadFileMasterList();

                    #region IVRF
                    if (IVRF != null)
                    {
                        var List = IVRF.Split('/');

                        PSUAUnderlyingID = Convert.ToInt32(List[0]);
                        PSUSkewUnderlyingID = Convert.ToInt32(List[1]);
                        PSUMUnderlyingID = Convert.ToInt32(List[2]);
                        PSUEUnderlyingID = Convert.ToInt32(List[3]);
                    }

                    DataSet dsResult = new DataSet();
                    dsResult = General.ExecuteDataSet("GET_UNDERLYING_ID_LIST");

                    List<Underlying> UnderlyingList = new List<Underlying>();

                    if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                    {
                        foreach (DataRow dr in dsResult.Tables[0].Rows)
                        {
                            Underlying obj = new Underlying();

                            obj.UnderlyingID = Convert.ToInt32(dr["ID"]);
                            obj.UnderlyingShortName = Convert.ToString(dr["UNDERLYING_SHORTNAME"]);

                            UnderlyingList.Add(obj);
                        }
                        objUnderlying.PSUAUnderLyingList = UnderlyingList;
                        objUnderlying.PSUSkewUnderLyingList = UnderlyingList;
                        objUnderlying.PSUMinimumUnderLyingList = UnderlyingList;
                        objUnderlying.PSUExceptionUnderLyingList = UnderlyingList;

                        //--Set default underlying--Added by Shweta on 3rd May 2016------------START--------------------
                        string strDefaultUnderlying = System.Configuration.ConfigurationManager.AppSettings["DefaultUnderlying"].ToUpper();
                        Underlying objDefaulyUnderlying = objUnderlying.PSUAUnderLyingList.Find(delegate(Underlying oUnderlying) { return oUnderlying.UnderlyingShortName == strDefaultUnderlying; });
                        //--Set default underlying--Added by Shweta on 3rd May 2016------------END----------------------

                        if (PSUAUnderlyingID == 0 && PSUSkewUnderlyingID == 0 && PSUMUnderlyingID == 0 && PSUEUnderlyingID == 0)
                        {
                            objUnderlying.FilterPSUAdjustment = objDefaulyUnderlying.UnderlyingID;
                            objUnderlying.FilterPSUSkew = objDefaulyUnderlying.UnderlyingID;
                            objUnderlying.FilterPSUMinimum = objDefaulyUnderlying.UnderlyingID;
                            objUnderlying.FilterPSUException = objDefaulyUnderlying.UnderlyingID;
                        }
                        else
                        {
                            objUnderlying.FilterPSUAdjustment = Convert.ToInt32(PSUAUnderlyingID);
                            objUnderlying.FilterPSUSkew = Convert.ToInt32(PSUSkewUnderlyingID);
                            objUnderlying.FilterPSUMinimum = Convert.ToInt32(PSUMUnderlyingID);
                            objUnderlying.FilterPSUException = Convert.ToInt32(PSUEUnderlyingID);
                        }
                    }
                    else
                    {
                        TempData["CreateUnderlying"] = "Create Underlying";

                        return RedirectToAction("UnderlyingCreation", "UnderlyingCreation");
                    }

                    #endregion

                    if (IVRF != null)
                        if (Version == null || Version == "0")
                        {
                            ViewBag.Message = "Version is not Mentioned in Excel";
                            return View(objUnderlying);
                        }
                        else if (Version == "-1")
                        {
                            ViewBag.VersionAlready = "Version Not Found";
                            return View(objUnderlying);
                        }
                        else
                        {
                            ViewBag.Successfull = "Successfully Updated";
                            return View(objUnderlying);
                        }


                    return View(objUnderlying);
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
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "PutAdjustmentSurface Get", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        [HttpPost]
        public ActionResult PutAdjustmentSurface(Underlying objUnderlying, string Command, FormCollection collection, HttpPostedFileBase file, HttpPostedFileBase Skew, HttpPostedFileBase PutMinimum)
        {
            LoginController objLoginController = new LoginController();
            List<UploadFileMaster> UploadFileMasterList = (List<UploadFileMaster>)Session["UploadFileMasterList"];
            // Underlying objUnderlying = new Underlying();

            try
            {
                if (ValidateSession())
                {
                    bool blnUploadStatus = false;
                    bool blnUploadDataStatus = true;

                    #region Put AdjustmentUpload
                    if (Command == "AdjustmentUpload")
                    {
                        if (file != null && file.ContentLength > 0)
                        {
                            UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "PAS"; });
                            string strFilePath = System.Web.HttpContext.Current.Server.MapPath("~/Uploads/");
                            string strFileName = Path.GetFileNameWithoutExtension(file.FileName) + DateTime.Now.ToString("dd_mmm_yyyy_hh_mm");
                            string strExtension = Path.GetExtension(file.FileName);
                            string Type = "";

                            int intAdjustmentUnderlyingID = Convert.ToInt32(collection.Get("FilterPSUAdjustment"));
                            Type = "24";

                            int intSkewUnderlyingID = Convert.ToInt32(collection.Get("FilterPSUSkew"));

                            int intMinimumUnderlyingID = Convert.ToInt32(collection.Get("FilterPSUMinimum"));

                            strFilePath += strFileName + strExtension;
                            file.SaveAs(strFilePath);

                            FileInfo newFile = new FileInfo(strFilePath);

                            #region Source and Destination Column
                            string strSourceColumn = objUploadFileMaster.SourceColumn;
                            string[] arrSourceColumn = null;
                            if (strSourceColumn != "")
                                arrSourceColumn = strSourceColumn.Split('|');

                            DataTable dtData = new DataTable();

                            for (int i = 0; i < arrSourceColumn.Length; i++)
                            {
                                dtData.Columns.Add(arrSourceColumn[i]);
                            }

                            DataTable dtColumnList = new DataTable();
                            dtColumnList.Columns.Add("ColumnName");

                            string strDestinationColumn = objUploadFileMaster.DestinationColumn;
                            string[] arrDestinationColumn = null;

                            if (strDestinationColumn != "")
                                arrDestinationColumn = strDestinationColumn.Split('|');

                            string strTableName = objUploadFileMaster.TableName;
                            #endregion

                            using (ExcelPackage xlPackage = new ExcelPackage(newFile))
                            {
                                ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets[objUploadFileMaster.SheetName];

                                string strVersion = worksheet.Cell(1, 2).Value.Trim();

                                if (strVersion == "" || strVersion == "0")
                                {
                                    var IVRF1 = intAdjustmentUnderlyingID + "/" + intSkewUnderlyingID + "/" + intMinimumUnderlyingID;

                                    return RedirectToAction("PutAdjustmentSurface", new { IVRF = IVRF1, Version = strVersion });
                                }
                                else
                                {
                                    for (int intCol = 2; intCol < 20; intCol++)
                                    {
                                        if (Convert.ToString(worksheet.Cell(3, intCol).Value) != "")
                                        {
                                            DataRow dr = dtColumnList.NewRow();
                                            dr["ColumnName"] = worksheet.Cell(3, intCol).Value;

                                            dtColumnList.Rows.Add(dr);
                                        }
                                    }

                                    DataRow drNew;

                                    //string strType = "Call";
                                    // output the data in column 2
                                    for (int iRow = 4; iRow < 36; iRow++)
                                    {
                                        if (Convert.ToString(worksheet.Cell(iRow, 1).Value) != "")
                                        {
                                            for (int intCol = 0; intCol < dtColumnList.Rows.Count; intCol++)
                                            {
                                                drNew = dtData.NewRow();



                                                drNew["VERSION"] = strVersion;
                                                drNew["TENURE"] = worksheet.Cell(iRow, 1).Value;
                                                drNew["MONEYNESS"] = Convert.ToDouble(dtColumnList.Rows[intCol][0]);
                                                drNew["VALUE"] = worksheet.Cell(iRow, intCol + 2).Value;
                                                drNew["CREATED_DATE"] = DateTime.Now;
                                                drNew["TYPE"] = Type;
                                                drNew["UNDERLYINGID"] = intAdjustmentUnderlyingID;




                                                dtData.Rows.Add(drNew);
                                            }
                                        }
                                    }
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
                                dsIV = General.ExecuteDataSet(objUploadFileMaster.ExtraProcedure, false);

                                var result = Convert.ToInt32(dsIV.Tables[0].Rows[0]["Result"]);
                                if (result == -1)
                                {
                                    var IVRF1 = intAdjustmentUnderlyingID + "/" + intSkewUnderlyingID + "/" + intMinimumUnderlyingID;

                                    return RedirectToAction("PutAdjustmentSurface", new { IVRF = IVRF1, Version = -1 });
                                }
                            }
                            else
                            {
                                blnUploadStatus = false;
                            }

                            if (blnUploadStatus)
                            {
                                ManageUploadFileInfo(intAdjustmentUnderlyingID, file.FileName, strFilePath, blnUploadStatus, blnUploadDataStatus);

                                ViewBag.Message = "Imported successfully";
                            }

                            objUnderlying = new Underlying();
                            objUnderlying = FetchDefaultDetails();

                            var IVRF = intAdjustmentUnderlyingID + "/" + intSkewUnderlyingID + "/" + intMinimumUnderlyingID;

                            return RedirectToAction("PutAdjustmentSurface", new { IVRF = IVRF, Version = 1 });
                        }
                    }
                    #endregion

                    #region Skew
                    else if (Command == "SkewUpload")
                    {
                        if (Skew != null && Skew.ContentLength > 0)
                        {
                            UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "PS"; });
                            string strFilePath = System.Web.HttpContext.Current.Server.MapPath("~/Uploads/");
                            string strFileName = Path.GetFileNameWithoutExtension(Skew.FileName) + DateTime.Now.ToString("dd_mmm_yyyy_hh_mm");
                            string strExtension = Path.GetExtension(Skew.FileName);

                            int intAdjustmentUnderlyingID = Convert.ToInt32(collection.Get("FilterPSUAdjustment"));

                            int intSkewUnderlyingID = Convert.ToInt32(collection.Get("FilterPSUSkew"));

                            int intMinimumUnderlyingID = Convert.ToInt32(collection.Get("FilterPSUMinimum"));

                            strFilePath += strFileName + strExtension;
                            Skew.SaveAs(strFilePath);

                            FileInfo newFile = new FileInfo(strFilePath);

                            #region Source and Destination Column
                            string strSourceColumn = objUploadFileMaster.SourceColumn;
                            string[] arrSourceColumn = null;
                            if (strSourceColumn != "")
                                arrSourceColumn = strSourceColumn.Split('|');

                            DataTable dtData = new DataTable();

                            for (int i = 0; i < arrSourceColumn.Length; i++)
                            {
                                dtData.Columns.Add(arrSourceColumn[i]);
                            }

                            DataTable dtColumnList = new DataTable();
                            dtColumnList.Columns.Add("ColumnName");

                            string strDestinationColumn = objUploadFileMaster.DestinationColumn;
                            string[] arrDestinationColumn = null;

                            if (strDestinationColumn != "")
                                arrDestinationColumn = strDestinationColumn.Split('|');

                            string strTableName = objUploadFileMaster.TableName;
                            #endregion


                            using (ExcelPackage xlPackage = new ExcelPackage(newFile))
                            {
                                ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets[objUploadFileMaster.SheetName];

                                string strVersion = worksheet.Cell(1, 2).Value.Trim();

                                if (strVersion == "" || strVersion == "0")
                                {
                                    var IVRF1 = intAdjustmentUnderlyingID + "/" + intSkewUnderlyingID + "/" + intMinimumUnderlyingID;

                                    return RedirectToAction("PutAdjustmentSurface", new { IVRF = IVRF1, Version = strVersion });
                                }
                                else
                                {

                                    DataRow drNew;

                                    //string strType = "Call";
                                    // output the data in column 2
                                    for (int iRow = 4; iRow < 36; iRow++)
                                    {
                                        if (Convert.ToString(worksheet.Cell(iRow, 1).Value) != "")
                                        {
                                            drNew = dtData.NewRow();

                                            drNew["VERSION"] = strVersion;
                                            drNew["TENURE"] = worksheet.Cell(iRow, 1).Value;
                                            drNew["STRIKE_DIFFERENCE"] = worksheet.Cell(iRow, 2).Value;
                                            drNew["MINIMUM_VOL_DIFFERENCE"] = worksheet.Cell(iRow, 3).Value;
                                            drNew["CREATED_DATE"] = DateTime.Now;
                                            drNew["UNDERLYINGID"] = intSkewUnderlyingID;

                                            dtData.Rows.Add(drNew);

                                        }
                                    }
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
                                dsIV = General.ExecuteDataSet(objUploadFileMaster.ExtraProcedure, false);

                                var result = Convert.ToInt32(dsIV.Tables[0].Rows[0]["Result"]);
                                if (result == -1)
                                {
                                    var IVRF1 = intAdjustmentUnderlyingID + "/" + intSkewUnderlyingID + "/" + intMinimumUnderlyingID;

                                    return RedirectToAction("PutAdjustmentSurface", new { IVRF = IVRF1, Version = -1 });
                                }
                            }
                            else
                            {
                                blnUploadStatus = false;
                            }

                            if (blnUploadStatus)
                            {
                                ManageUploadFileInfo(intSkewUnderlyingID, Skew.FileName, strFilePath, blnUploadStatus, blnUploadDataStatus);

                                ViewBag.Message = "Imported successfully";
                            }
                            objUnderlying = new Underlying();
                            objUnderlying = FetchDefaultDetails();

                            var IVRF = intAdjustmentUnderlyingID + "/" + intSkewUnderlyingID + "/" + intMinimumUnderlyingID;

                            return RedirectToAction("PutAdjustmentSurface", new { IVRF = IVRF, Version = 1 });
                        }
                    }
                    #endregion

                    #region Minimum
                    else if (Command == "PutMinimumUpload")
                    {
                        if (PutMinimum != null && PutMinimum.ContentLength > 0)
                        {
                            UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "PIV"; });
                            string strFilePath = System.Web.HttpContext.Current.Server.MapPath("~/Uploads/");
                            string strFileName = Path.GetFileNameWithoutExtension(PutMinimum.FileName) + DateTime.Now.ToString("dd_mmm_yyyy_hh_mm");
                            string strExtension = Path.GetExtension(PutMinimum.FileName);

                            int intAdjustmentUnderlyingID = Convert.ToInt32(collection.Get("FilterPSUAdjustment"));

                            int intSkewUnderlyingID = Convert.ToInt32(collection.Get("FilterPSUSkew"));

                            int intMinimumUnderlyingID = Convert.ToInt32(collection.Get("FilterPSUMinimum"));

                            strFilePath += strFileName + strExtension;
                            PutMinimum.SaveAs(strFilePath);

                            FileInfo newFile = new FileInfo(strFilePath);

                            #region Source and Destination Column
                            string strSourceColumn = objUploadFileMaster.SourceColumn;
                            string[] arrSourceColumn = null;
                            if (strSourceColumn != "")
                                arrSourceColumn = strSourceColumn.Split('|');

                            DataTable dtData = new DataTable();

                            for (int i = 0; i < arrSourceColumn.Length; i++)
                            {
                                dtData.Columns.Add(arrSourceColumn[i]);
                            }

                            DataTable dtColumnList = new DataTable();
                            dtColumnList.Columns.Add("ColumnName");

                            string strDestinationColumn = objUploadFileMaster.DestinationColumn;
                            string[] arrDestinationColumn = null;

                            if (strDestinationColumn != "")
                                arrDestinationColumn = strDestinationColumn.Split('|');

                            string strTableName = objUploadFileMaster.TableName;
                            #endregion


                            using (ExcelPackage xlPackage = new ExcelPackage(newFile))
                            {
                                ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets[objUploadFileMaster.SheetName];

                                string strVersion = worksheet.Cell(1, 2).Value.Trim();

                                if (strVersion == "" || strVersion == "0")
                                {
                                    var IVRF1 = intAdjustmentUnderlyingID + "/" + intSkewUnderlyingID + "/" + intMinimumUnderlyingID;

                                    return RedirectToAction("PutAdjustmentSurface", new { IVRF = IVRF1, Version = strVersion });
                                }
                                else
                                {

                                    DataRow drNew;

                                    //string strType = "Call";
                                    // output the data in column 2
                                    for (int iRow = 4; iRow < 36; iRow++)
                                    {
                                        if (Convert.ToString(worksheet.Cell(iRow, 1).Value) != "")
                                        {
                                            drNew = dtData.NewRow();

                                            drNew["VERSION"] = strVersion;
                                            drNew["TENURE"] = worksheet.Cell(iRow, 1).Value;
                                            drNew["STRIKE_DIFFERENCE"] = worksheet.Cell(iRow, 2).Value;
                                            drNew["MINIMUM_VOL_DIFFERENCE"] = worksheet.Cell(iRow, 3).Value;
                                            drNew["CREATED_DATE"] = DateTime.Now;
                                            drNew["UNDERLYINGID"] = intMinimumUnderlyingID;

                                            dtData.Rows.Add(drNew);

                                        }
                                    }
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
                                dsIV = General.ExecuteDataSet(objUploadFileMaster.ExtraProcedure, false);

                                var result = Convert.ToInt32(dsIV.Tables[0].Rows[0]["Result"]);
                                if (result == -1)
                                {
                                    var IVRF1 = intAdjustmentUnderlyingID + "/" + intSkewUnderlyingID + "/" + intMinimumUnderlyingID;

                                    return RedirectToAction("PutAdjustmentSurface", new { IVRF = IVRF1, Version = -1 });
                                }
                            }
                            else
                            {
                                blnUploadStatus = false;
                            }

                            if (blnUploadStatus)
                            {
                                ManageUploadFileInfo(intMinimumUnderlyingID, PutMinimum.FileName, strFilePath, blnUploadStatus, blnUploadDataStatus);

                                ViewBag.Message = "Imported successfully";
                            }
                            objUnderlying = new Underlying();
                            objUnderlying = FetchDefaultDetails();

                            var IVRF = intAdjustmentUnderlyingID + "/" + intSkewUnderlyingID + "/" + intMinimumUnderlyingID;

                            return RedirectToAction("PutAdjustmentSurface", new { IVRF = IVRF, Version = 1 });
                        }
                    }
                    #endregion

                    #region Spread Adjustment Surface
                    else if (Command == "PASDownload")
                    {
                        UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "PAS"; });
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

                        int intAdjustmentUnderlyingID = Convert.ToInt32(collection.Get("FilterPSUAdjustment"));
                        int intSkewUnderlyingID = Convert.ToInt32(collection.Get("FilterPSUSkew"));
                        int intMinimumUnderlyingID = Convert.ToInt32(collection.Get("FilterPSUMinimum"));
                        var IVRF = intAdjustmentUnderlyingID + "/" + intSkewUnderlyingID + "/" + intMinimumUnderlyingID;
                        return RedirectToAction("PutAdjustmentSurface", new { IVRF = IVRF, Version = 1 });
                    }
                    #endregion

                    #region Spread Skew Adjustment
                    else if (Command == "PSDownload")
                    {
                        UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "PS"; });
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

                        int intAdjustmentUnderlyingID = Convert.ToInt32(collection.Get("FilterPSUAdjustment"));
                        int intSkewUnderlyingID = Convert.ToInt32(collection.Get("FilterPSUSkew"));
                        int intMinimumUnderlyingID = Convert.ToInt32(collection.Get("FilterPSUMinimum"));
                        var IVRF = intAdjustmentUnderlyingID + "/" + intSkewUnderlyingID + "/" + intMinimumUnderlyingID;
                        return RedirectToAction("PutAdjustmentSurface", new { IVRF = IVRF, Version = 1 });
                    }
                    #endregion

                    #region Spread Minimum
                    else if (Command == "PMDownload")
                    {
                        UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "PIV"; });
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

                        int intAdjustmentUnderlyingID = Convert.ToInt32(collection.Get("FilterPSUAdjustment"));
                        int intSkewUnderlyingID = Convert.ToInt32(collection.Get("FilterPSUSkew"));
                        int intMinimumUnderlyingID = Convert.ToInt32(collection.Get("FilterPSUMinimum"));
                        var IVRF = intAdjustmentUnderlyingID + "/" + intSkewUnderlyingID + "/" + intMinimumUnderlyingID;
                        return RedirectToAction("PutAdjustmentSurface", new { IVRF = IVRF, Version = 1 });
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
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "PutAdjustmentSurface Post", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        public JsonResult FetchPutAdjusment(int UnderlyingId)
        {
            try
            {
                if (UnderlyingId == 0)
                {
                    UnderlyingId = Convert.ToInt32(Session["UnderlyingID"]);
                }
                List<Call_Put_AdjustmentSurface> AdjustmentSurfaceList = new List<Call_Put_AdjustmentSurface>();

                var lookupCode = "PSA";
                DataSet dsResult1 = new DataSet();
                dsResult1 = General.ExecuteDataSet("FETCH_PUT_ADJUSMENT_SURFACE", lookupCode, UnderlyingId);

                string strHTML = "<table border='1'>";
                strHTML += "<tr style='background-color:#BCBCBC;color:Black;'>";

                for (int intCol = 0; intCol < dsResult1.Tables[0].Columns.Count; intCol++)
                {
                    strHTML += "<td>";
                    strHTML += dsResult1.Tables[0].Columns[intCol].Caption;
                    strHTML += "</td>";
                }
                strHTML += "</tr>";

                for (int intRow = 0; intRow < dsResult1.Tables[0].Rows.Count; intRow++)
                {
                    strHTML += "<tr>";
                    for (int intCol = 0; intCol < dsResult1.Tables[0].Columns.Count; intCol++)
                    {
                        strHTML += "<td>";
                        strHTML += dsResult1.Tables[0].Rows[intRow][intCol];
                        strHTML += "</td>";
                    }
                    strHTML += "</tr>";
                }
                strHTML += "</table>";

                if (dsResult1.Tables[0].Rows.Count == 0)
                    strHTML = "No Records Found";

                return Json(strHTML, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "FetchPutAdjusment", objUserMaster.UserID);
                return Json("");
            }
        }


        public JsonResult FetchPutSkewAdjustment(int UnderlyingId)
        {
            try
            {
                if (UnderlyingId == 0)
                {
                    UnderlyingId = Convert.ToInt32(Session["UnderlyingID"]);
                }
                List<PutSkewAdjustment> SkewAdjustmentList = new List<PutSkewAdjustment>();


                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("FETCH_PUT_SKEW_ADJUSTMENT", UnderlyingId);

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in dsResult.Tables[0].Rows)
                    {
                        PutSkewAdjustment obj = new PutSkewAdjustment();

                        obj.Version = Convert.ToInt32(dr["VERSION"]);
                        obj.Tenure = Convert.ToString(dr["TENURE"]);
                        obj.StrikeDifference = Convert.ToDouble(dr["STRIKE_DIFFERENCE"]);
                        obj.MinimumVolDifference = Convert.ToDouble(dr["MINIMUM_VOL_DIFFERENCE"]);
                        obj.CreatedDate = Convert.ToDateTime(dr["CREATED_DATE"]);
                        obj.UnderlyingShortName = Convert.ToString(dr["UNDERLYING_SHORTNAME"]);

                        SkewAdjustmentList.Add(obj);
                    }
                }

                var TSkewAdjustmentListData = SkewAdjustmentList.ToList();
                return Json(TSkewAdjustmentListData, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "FetchPutSkewAdjustment", objUserMaster.UserID);
                return Json("");
            }
        }


        public JsonResult FetchPutMinimum(int UnderlyingId)
        {
            try
            {
                if (UnderlyingId == 0)
                {
                    UnderlyingId = Convert.ToInt32(Session["UnderlyingID"]);
                }
                List<PutMinimumGap> PutMinimumGapList = new List<PutMinimumGap>();


                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("FETCH_PUT_MINIMUM_GAP", UnderlyingId);

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in dsResult.Tables[0].Rows)
                    {
                        PutMinimumGap obj = new PutMinimumGap();

                        obj.Version = Convert.ToInt32(dr["VERSION"]);
                        obj.Tenure = Convert.ToString(dr["TENURE"]);
                        obj.StrikeDifference = Convert.ToDouble(dr["STRIKE_DIFFERENCE"]);
                        obj.MinimumVolDifference = Convert.ToDouble(dr["MINIMUM_VOL_DIFFERENCE"]);
                        obj.CreatedDate = Convert.ToDateTime(dr["CREATED_DATE"]);
                        obj.UnderlyingShortName = Convert.ToString(dr["UNDERLYING_SHORTNAME"]);

                        PutMinimumGapList.Add(obj);
                    }
                }

                var TSkewAdjustmentListData = PutMinimumGapList.ToList();
                return Json(TSkewAdjustmentListData, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "FetchPutMinimum", objUserMaster.UserID);
                return Json("");
            }
        }

        #endregion

        #region Pricers Tolerance

        [HttpGet]
        public ActionResult PricerTolerance()
        {
            LoginController objLoginController = new LoginController();
            Underlying objUnderlying = new Underlying();

            try
            {
                if (ValidateSession())
                {
                    #region Menu Access By on Role

                    Int32 intResult = 0;
                    // bool PPorNonPP = false;

                    UserMaster objUserMaster = new UserMaster();
                    objUserMaster = (UserMaster)Session["LoggedInUser"];

                    var Result = objSP_PRICINGEntities.VALIDATE_MENU_BY_ROLE(objUserMaster.RoleID, "UCPT");
                    intResult = Convert.ToInt32(Result.SingleOrDefault());

                    if (intResult == 0)
                        return RedirectToAction("UserNotAuthorize", "Login");

                    #endregion

                    FetchUploadFileMasterList();

                    return View(objUnderlying);
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
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "PricerTolerance Get", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        [HttpPost]
        public ActionResult PricerTolerance(Underlying objUnderlying, string Command, FormCollection collection, HttpPostedFileBase file)
        {
            LoginController objLoginController = new LoginController();
            List<UploadFileMaster> UploadFileMasterList = (List<UploadFileMaster>)Session["UploadFileMasterList"];

            try
            {
                if (ValidateSession())
                {
                    bool blnUploadStatus = false;
                    bool blnUploadDataStatus = true;

                    #region Price Tolerance Upload
                    if (Command == "PDUpload")
                    {
                        if (file != null && file.ContentLength > 0)
                        {
                            UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "PTL"; });
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
                            #endregion

                            using (ExcelPackage xlPackage = new ExcelPackage(newFile))
                            {
                                ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets[objUploadFileMaster.SheetName];

                                string strVersion = worksheet.Cell(1, 2).Value.Trim();
                                if (strVersion == "")
                                    strVersion = "0";

                                UserMaster objUserMaster = new UserMaster();
                                objUserMaster = (UserMaster)Session["LoggedInUser"];

                                DataRow drNew;

                                //string strType = "Call";
                                // output the data in column 2
                                for (int iRow = 4; iRow < 36; iRow++)
                                {
                                    if (Convert.ToString(worksheet.Cell(iRow, 1).Value) != "")
                                    {
                                        drNew = dtData.NewRow();

                                        drNew["VERSION"] = strVersion;
                                        drNew["PRICERS"] = worksheet.Cell(iRow, 1).Value;
                                        drNew["BUILT_IN"] = worksheet.Cell(iRow, 2).Value;
                                        drNew["NO_OF_DAYS_REDEMPTION_PERIOD"] = worksheet.Cell(iRow, 3).Value;
                                        drNew["FIXED_COUPON"] = worksheet.Cell(iRow, 4).Value;
                                        drNew["OTHER_COUPON"] = worksheet.Cell(iRow, 5).Value;
                                        //drNew["COUPON_MAX_MIN_RATE"] = worksheet.Cell(iRow, 6).Value;     --Commented by Shweta on 5th May 2016
                                        drNew["CREATED_DATE"] = DateTime.Now;
                                        drNew["CREATED_BY"] = objUserMaster.UserID;

                                        dtData.Rows.Add(drNew);
                                    }
                                    else
                                        break;
                                }
                                //}
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
                            objUnderlying = new Underlying();
                            objUnderlying = FetchDefaultDetails();

                            return View(objUnderlying);
                        }
                    }
                    #endregion

                    #region Pricing Tolerance Download
                    else if (Command == "PDDownload")
                    {
                        UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "PTL"; });
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
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "PricerTolerance Post", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        public JsonResult FetchPricersTolerance()
        {
            try
            {
                List<PricerTolerance> PricerToleranceList = new List<PricerTolerance>();


                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("FETCH_PRICER_TOLERANCE");

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in dsResult.Tables[0].Rows)
                    {
                        PricerTolerance obj = new PricerTolerance();

                        obj.Version = Convert.ToInt32(dr["VERSION"]);
                        obj.Pricers = Convert.ToString(dr["PRICERS"]);
                        obj.DistributorBuiltIn = Convert.ToDouble(dr["BUILT_IN"]);
                        //obj.EdelweissBuiltIn = Convert.ToDouble(dr["EDELWEISS_BUILT_IN"]);
                        obj.DaysRedemptionPeriod = Convert.ToString(dr["NO_OF_DAYS_REDEMPTION_PERIOD"]);
                        obj.FixedCoupon = Convert.ToDouble(dr["FIXED_COUPON"]);
                        obj.OtherCoupon = Convert.ToDouble(dr["OTHER_COUPON"]);
                        obj.CreatedDate = Convert.ToDateTime(dr["CREATED_DATE"]);

                        PricerToleranceList.Add(obj);
                    }
                }

                var PricerToleranceListData = PricerToleranceList.ToList();
                return Json(PricerToleranceListData, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "FetchPricersTolerance", objUserMaster.UserID);
                return Json("");
            }
        }

        #endregion

        #region Underlying Price
        [HttpGet]
        public ActionResult UnderlyingPrice()
        {
            LoginController objLoginController = new LoginController();
            UnderlyingPrice objUnderlyingPrice = new UnderlyingPrice();

            try
            {
                if (ValidateSession())
                {
                    #region Menu Access By on Role

                    Int32 intResult = 0;
                    // bool PPorNonPP = false;

                    UserMaster objUserMaster = new UserMaster();
                    objUserMaster = (UserMaster)Session["LoggedInUser"];

                    var Result = objSP_PRICINGEntities.VALIDATE_MENU_BY_ROLE(objUserMaster.RoleID, "UPR");
                    intResult = Convert.ToInt32(Result.SingleOrDefault());

                    if (intResult == 0)
                        return RedirectToAction("UserNotAuthorize", "Login");

                    #endregion

                    FetchUploadFileMasterList();
                    return View(objUnderlyingPrice);
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
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "UnderlyingPrice Get", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        [HttpPost]
        public ActionResult UnderlyingPrice(Underlying objUnderlying, string Command, FormCollection collection, HttpPostedFileBase file)
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
                            UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "UP"; });
                            string strFilePath = System.Web.HttpContext.Current.Server.MapPath("~/Uploads/");
                            string strFileName = Path.GetFileNameWithoutExtension(file.FileName) + DateTime.Now.ToString("dd_mmm_yyyy_hh_mm");
                            string strExtension = Path.GetExtension(file.FileName);

                            strFilePath += strFileName + strExtension;
                            file.SaveAs(strFilePath);

                            //string connectionString;
                            //string strCommand;

                            //connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + strFilePath + "';Extended Properties='Excel 12.0;HDR=YES;IMEX=1'";
                            //strCommand = "SELECT * FROM [Sheet1$]";
                            //// strCommand = "SELECT * FROM " + objUploadFileMaster.SheetName +  "";

                            //OleDbConnection objConnection = new System.Data.OleDb.OleDbConnection(connectionString);
                            //objConnection.Open();
                            //DbDataAdapter objAdapter = new System.Data.OleDb.OleDbDataAdapter(strCommand, objConnection);
                            //DataSet dsResult = new DataSet();
                            //objAdapter.Fill(dsResult);

                            //objConnection.Close();

                            FileInfo newFile = new FileInfo(strFilePath);

                            #region Source and Destination Column
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

                                        drNew["UNDERLYING_NAME"] = worksheet.Cell(iRow, 1).Value;
                                        //drNew["UNDERLYING_ID"] = "0";
                                        drNew["DATE"] = DateTime.FromOADate(Convert.ToDouble(worksheet.Cell(iRow, 2).Value)).ToString("dd-MM-yyyy");
                                        drNew["PRICE"] = worksheet.Cell(iRow, 3).Value;

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

                            return RedirectToAction("UnderlyingPrice");
                        }
                    }
                    else if (Command == "Download")
                    {

                        UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "UP"; });
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

                            objUnderlying = new Underlying();
                            objUnderlying = FetchDefaultDetails();

                            return RedirectToAction("UnderlyingPrice");
                        }
                        else
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
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "UnderlyingPrice Post", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        public JsonResult FetchUnderlyingPriceDetails()
        {
            try
            {
                List<UnderlyingPrice> UnderlyingPriceList = new List<UnderlyingPrice>();

                ObjectResult<UnderlyingPriceResult> objUnderlyingPriceResult = objSP_PRICINGEntities.SP_FETCH_UNDERLYING_PRICE_DETAILS();
                List<UnderlyingPriceResult> UnderlyingPriceResultList = objUnderlyingPriceResult.ToList();

                foreach (UnderlyingPriceResult oUnderlyingPriceResult in UnderlyingPriceResultList)
                {
                    UnderlyingPrice objUnderlyingPrice = new UnderlyingPrice();

                    General.ReflectSingleData(objUnderlyingPrice, oUnderlyingPriceResult);
                    UnderlyingPriceList.Add(objUnderlyingPrice);
                }

                var UnderlyingPriceListData = UnderlyingPriceList.ToList();
                return Json(UnderlyingPriceListData, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "FetchBasketInstruments", objUserMaster.UserID);
                return Json("");
            }
        }

        public JsonResult FetchUnderlyingPriceErrorDetails()
        {
            try
            {
                List<UnderlyingPrice> UnderlyingPriceErrorList = new List<UnderlyingPrice>();

                ObjectResult<UnderlyingPriceErrorResult> objUnderlyingPriceErrorResult = objSP_PRICINGEntities.SP_FETCH_UNDERLYING_PRICE_ERROR_DETAILS();
                List<UnderlyingPriceErrorResult> UnderlyingPriceErrorResultList = objUnderlyingPriceErrorResult.ToList();

                foreach (UnderlyingPriceErrorResult oUnderlyingPriceErrorResult in UnderlyingPriceErrorResultList)
                {
                    UnderlyingPrice objUnderlyingPrice = new UnderlyingPrice();

                    General.ReflectSingleData(objUnderlyingPrice, oUnderlyingPriceErrorResult);
                    UnderlyingPriceErrorList.Add(objUnderlyingPrice);
                }

                var UnderlyingPriceErrorListData = UnderlyingPriceErrorList.ToList();
                return Json(UnderlyingPriceErrorListData, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "FetchBasketInstruments", objUserMaster.UserID);
                return Json("");
            }
        }
        #endregion

        #region Underlying IV

        [HttpGet]
        public ActionResult UnderlyingIV()
        {
            LoginController objLoginController = new LoginController();
            UnderlyingIV objUnderlyingIV = new UnderlyingIV();

            try
            {
                if (ValidateSession())
                {
                    #region Menu Access By on Role

                    Int32 intResult = 0;
                    // bool PPorNonPP = false;

                    UserMaster objUserMaster = new UserMaster();
                    objUserMaster = (UserMaster)Session["LoggedInUser"];

                    var Result = objSP_PRICINGEntities.VALIDATE_MENU_BY_ROLE(objUserMaster.RoleID, "CUIV");
                    intResult = Convert.ToInt32(Result.SingleOrDefault());

                    if (intResult == 0)
                        return RedirectToAction("UserNotAuthorize", "Login");

                    #endregion

                    return View(objUnderlyingIV);
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
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "UnderlyingIV Get", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        [HttpPost]
        public ActionResult UnderlyingIV(Underlying objUnderlying, string Command, FormCollection collection, HttpPostedFileBase file)
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
                            UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "UIV"; });
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
                                        drNew["UNDERLYING"] = worksheet.Cell(iRow, 2).Value;
                                        drNew["STRIKE_MULTIPLIER"] = worksheet.Cell(iRow, 3).Value;
                                        drNew["INSTRUMENT_TYPE"] = worksheet.Cell(iRow, 4).Value;
                                        //drNew["EXPIRY_DATE"] = DateTime.FromOADate(Convert.ToDouble(worksheet.Cell(iRow, 5).Value));
                                        drNew["EXPIRY_DATE"] = DateTime.FromOADate(Convert.ToDouble(worksheet.Cell(iRow, 5).Value)).ToString("yyyy-MM-dd");
                                        drNew["DIRECTION"] = worksheet.Cell(iRow, 6).Value;
                                        drNew["OPTION_TYPE"] = worksheet.Cell(iRow, 7).Value;
                                        //drNew["FROM_DATE"] = DateTime.FromOADate(Convert.ToDouble(worksheet.Cell(iRow, 7).Value));
                                        //drNew["TO_DATE"] = DateTime.FromOADate(Convert.ToDouble(worksheet.Cell(iRow, 8).Value));
                                        drNew["IV_VALUE"] = worksheet.Cell(iRow, 8).Value;
                                        drNew["RFR_VALUE"] = worksheet.Cell(iRow, 9).Value;
                                        drNew["IS_ACTIVE"] = worksheet.Cell(iRow, 10).Value;

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

                            return RedirectToAction("UnderlyingIV");
                        }
                    }
                    #region IVDownload
                    else if (Command == "Download")
                    {

                        UploadFileMaster objUploadFileMaster = UploadFileMasterList.Find(delegate(UploadFileMaster oUploadFileMaster) { return oUploadFileMaster.UploadTypeCode == "UIV"; });
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

                            objUnderlying = new Underlying();
                            objUnderlying = FetchDefaultDetails();

                            return RedirectToAction("UnderlyingIV");
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
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "UnderlyingIV Post", objUserMaster.UserID);
                return RedirectToAction("ErrorDetails", "Login");
            }
        }

        public JsonResult FetchUnderlyingIVDetails(string ProductID, string Underlying, string StrikeMultiplier, string InstrumentType, string ExpiryDate, string Direction, string IVValue, string RFRValue, string OptionType, string FromDate, string ToDate)
        {
            try
            {
                List<UnderlyingIV> UnderlyingIVList = new List<UnderlyingIV>();

                if (ProductID == "" || ProductID == "--Select--")
                    ProductID = "ALL";

                if (OptionType == "" || OptionType == "--Select--")
                    OptionType = "ALL";

                if (Underlying == "" || Underlying == "0")
                    Underlying = "ALL";

                if (StrikeMultiplier == "0")
                    StrikeMultiplier = "ALL";

                if (InstrumentType == "" || InstrumentType == "0")
                    InstrumentType = "ALL";

                if (ExpiryDate == "")
                    ExpiryDate = "1900-01-01";

                if (Direction == "" || Direction == "0")
                    Direction = "ALL";

                if (IVValue == "" || IVValue == "0")
                    IVValue = "ALL";

                if (RFRValue == "" || RFRValue == "0")
                    RFRValue = "ALL";

                if (FromDate == "")
                    FromDate = "1900-01-01";

                if (ToDate == "")
                    ToDate = "2900-01-01";

                ObjectResult<UnderlyingIVResult> objUnderlyingIVResult = objSP_PRICINGEntities.SP_FETCH_UNDERLYING_IV_DETAILS(ProductID, Underlying, StrikeMultiplier, InstrumentType, ExpiryDate, Direction, IVValue, RFRValue, OptionType, FromDate, ToDate);
                List<UnderlyingIVResult> UnderlyingIVResultList = objUnderlyingIVResult.ToList();

                foreach (UnderlyingIVResult oUnderlyingIVResult in UnderlyingIVResultList)
                {
                    UnderlyingIV objUnderlyingIV = new UnderlyingIV();

                    General.ReflectSingleData(objUnderlyingIV, oUnderlyingIVResult);
                    UnderlyingIVList.Add(objUnderlyingIV);
                }

                var UnderlyingIVListData = UnderlyingIVList.ToList();
                return Json(UnderlyingIVListData, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "FetchUnderlyingIVDetails", objUserMaster.UserID);
                return Json("");
            }
        }

        public JsonResult FetchUnderlyingIVErrorDetails()
        {
            try
            {
                List<UnderlyingIV> UnderlyingIVErrorList = new List<UnderlyingIV>();

                ObjectResult<UnderlyingIVErrorResult> objUnderlyingIVErrorResult = objSP_PRICINGEntities.SP_FETCH_UNDERLYING_IV_ERROR_DETAILS();
                List<UnderlyingIVErrorResult> UnderlyingIVErrorResultList = objUnderlyingIVErrorResult.ToList();

                foreach (UnderlyingIVErrorResult oUnderlyingIVErrorResult in UnderlyingIVErrorResultList)
                {
                    UnderlyingIV objUnderlyingIV = new UnderlyingIV();

                    General.ReflectSingleData(objUnderlyingIV, oUnderlyingIVErrorResult);
                    UnderlyingIVErrorList.Add(objUnderlyingIV);
                }

                var UnderlyingIVErrorListData = UnderlyingIVErrorList.ToList();
                return Json(UnderlyingIVErrorListData, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "FetchUnderlyingIVErrorDetails", objUserMaster.UserID);
                return Json("");
            }
        }
        #endregion


        public void ManageUploadFileInfo(Int32 intUnderlyingID, string strOriginalFileName, string strFilePath, bool blnUploadStatus, bool blnUploadDataStatus)
        {
            Int32 intUploadType = 0;
            UserMaster objUserMaster = new UserMaster();
            objUserMaster = (UserMaster)Session["LoggedInUser"];


            Int32 intResult = 0;
            var Count = objSP_PRICINGEntities.SP_MANAGE_UPLOAD_FILE_INFO(intUnderlyingID, intUploadType, strOriginalFileName, Path.GetFileName(strFilePath), strFilePath, blnUploadStatus, blnUploadDataStatus, objUserMaster.UserID);
            intResult = Count.SingleOrDefault().Value;
        }

        public string GenerateUniqueFileName(HttpPostedFileBase file)
        {
            string strFilePath = System.Web.HttpContext.Current.Server.MapPath("~/Uploads/");
            string strFileName = Path.GetFileNameWithoutExtension(file.FileName) + DateTime.Now.ToString("dd_mmm_yyyy_hh_mm") + Path.GetExtension(file.FileName);

            strFilePath = strFilePath + "" + strFileName;

            return strFilePath;
        }

        public JsonResult FetchBasketInstruments(int UnderlyingId)
        {
            try
            {
                if (UnderlyingId == 0)
                {
                    UnderlyingId = Convert.ToInt32(Session["UnderlyingID"]);
                }
                List<BasketInstruments> BasketInstrumentsList = new List<BasketInstruments>();

                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("FETCH_BASKET_INSTRUMENT", UnderlyingId);

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in dsResult.Tables[0].Rows)
                    {
                        BasketInstruments obj = new BasketInstruments();

                        obj.UnderlyingID = Convert.ToInt32(dr["UNDERLYING_ID"]);
                        obj.BasketInstrument = Convert.ToString(dr["BASKET_INSTRUMENT"]);
                        obj.Weightage = Convert.ToInt32(dr["WEIGHTAGE"]);

                        BasketInstrumentsList.Add(obj);
                    }
                }

                var BasketInstrumentsListtData = BasketInstrumentsList.ToList();
                return Json(BasketInstrumentsListtData, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "FetchBasketInstruments", objUserMaster.UserID);
                return Json("");
            }
        }

        public JsonResult FetchCallAdjusmentSurface(int UnderlyingId)
        {
            try
            {
                if (UnderlyingId == 0)
                {
                    UnderlyingId = Convert.ToInt32(Session["UnderlyingID"]);
                }
                List<Call_Put_AdjustmentSurface> AdjustmentSurfaceList = new List<Call_Put_AdjustmentSurface>();

                var lookupCode = "CSA";
                DataSet dsResult1 = new DataSet();
                dsResult1 = General.ExecuteDataSet("FETCH_CALL_PUT_ADJUSMENT_SURFACE", lookupCode, UnderlyingId);


                string strHTML = "<table border='1'>";
                strHTML += "<tr style='background-color:#BCBCBC;color:Black;'>";

                for (int intCol = 0; intCol < dsResult1.Tables[0].Columns.Count; intCol++)
                {
                    strHTML += "<td>";
                    strHTML += dsResult1.Tables[0].Columns[intCol].Caption;
                    strHTML += "</td>";
                }
                strHTML += "</tr>";

                for (int intRow = 0; intRow < dsResult1.Tables[0].Rows.Count; intRow++)
                {
                    strHTML += "<tr>";
                    for (int intCol = 0; intCol < dsResult1.Tables[0].Columns.Count; intCol++)
                    {
                        strHTML += "<td>";
                        strHTML += dsResult1.Tables[0].Rows[intRow][intCol];
                        strHTML += "</td>";
                    }
                    strHTML += "</tr>";
                }
                strHTML += "</table>";

                if (dsResult1.Tables[0].Rows.Count == 0)
                    strHTML = "No Records Found";

                return Json(strHTML, JsonRequestBehavior.AllowGet);
                //if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                //{
                //    foreach (DataRow dr in dsResult.Tables[0].Rows)
                //    {
                //        Call_Put_AdjustmentSurface obj = new Call_Put_AdjustmentSurface();

                //        obj.Version = Convert.ToInt32(dr["VERSION"]);
                //        obj.Tenure = Convert.ToString(dr["TENURE"]);
                //        obj.Moneyness = Convert.ToDouble(dr["MONEYNESS"]);
                //        obj.Value = Convert.ToDouble(dr["VALUE"]);
                //        obj.CreatedDate = Convert.ToDateTime(dr["CREATED_DATE"]);
                //        obj.LookUpId = Convert.ToInt32(dr["TYPE"]);
                //        //obj.UnderlyingID = Convert.ToInt32(dr["UNDERLYINGID"]);
                //        obj.UnderlyingShortName = Convert.ToString(dr["UNDERLYING_SHORTNAME"]);

                //        AdjustmentSurfaceList.Add(obj);
                //    }
                //}

                //var AdjustmentSurfaceListData = AdjustmentSurfaceList.ToList();
                //return Json(AdjustmentSurfaceListData, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "FetchCallAdjusmentSurface", objUserMaster.UserID);
                return Json("");
            }
        }

        public JsonResult FetchPutAdjusmentSurface(int UnderlyingId)
        {
            try
            {
                if (UnderlyingId == 0)
                {
                    UnderlyingId = Convert.ToInt32(Session["UnderlyingID"]);
                }
                List<Call_Put_AdjustmentSurface> AdjustmentSurfaceList = new List<Call_Put_AdjustmentSurface>();

                var lookupCode = "PSA";
                DataSet dsResult1 = new DataSet();
                dsResult1 = General.ExecuteDataSet("FETCH_CALL_PUT_ADJUSMENT_SURFACE", lookupCode, UnderlyingId);

                string strHTML = "<table border='1'>";
                strHTML += "<tr style='background-color:#BCBCBC;color:Black;'>";

                for (int intCol = 0; intCol < dsResult1.Tables[0].Columns.Count; intCol++)
                {
                    strHTML += "<td>";
                    strHTML += dsResult1.Tables[0].Columns[intCol].Caption;
                    strHTML += "</td>";
                }
                strHTML += "</tr>";

                for (int intRow = 0; intRow < dsResult1.Tables[0].Rows.Count; intRow++)
                {
                    strHTML += "<tr>";
                    for (int intCol = 0; intCol < dsResult1.Tables[0].Columns.Count; intCol++)
                    {
                        strHTML += "<td>";
                        strHTML += dsResult1.Tables[0].Rows[intRow][intCol];
                        strHTML += "</td>";
                    }
                    strHTML += "</tr>";
                }
                strHTML += "</table>";

                if (dsResult1.Tables[0].Rows.Count == 0)
                    strHTML = "No Records Found";

                return Json(strHTML, JsonRequestBehavior.AllowGet);

                //if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                //{
                //    foreach (DataRow dr in dsResult.Tables[0].Rows)
                //    {
                //        Call_Put_AdjustmentSurface obj = new Call_Put_AdjustmentSurface();

                //        obj.Version = Convert.ToInt32(dr["VERSION"]);
                //        obj.Tenure = Convert.ToString(dr["TENURE"]);
                //        obj.Moneyness = Convert.ToDouble(dr["MONEYNESS"]);
                //        obj.Value = Convert.ToDouble(dr["VALUE"]);
                //        obj.CreatedDate = Convert.ToDateTime(dr["CREATED_DATE"]);
                //        obj.LookUpId = Convert.ToInt32(dr["TYPE"]);
                //        //obj.UnderlyingID = Convert.ToInt32(dr["UNDERLYINGID"]);
                //        obj.UnderlyingShortName = Convert.ToString(dr["UNDERLYING_SHORTNAME"]);

                //        AdjustmentSurfaceList.Add(obj);
                //    }
                //}

                //var AdjustmentSurfaceListData = AdjustmentSurfaceList.ToList();
                //return Json(AdjustmentSurfaceListData, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "FetchPutAdjusmentSurface", objUserMaster.UserID);
                return Json("");
            }
        }

        public JsonResult FetchPutAdjusmentException(int UnderlyingId)
        {
            try
            {
                if (UnderlyingId == 0)
                {
                    UnderlyingId = Convert.ToInt32(Session["UnderlyingID"]);
                }
                List<Call_Put_AdjustmentSurface> AdjustmentSurfaceList = new List<Call_Put_AdjustmentSurface>();

                var lookupCode = "PSA";
                DataSet dsResult1 = new DataSet();
                dsResult1 = General.ExecuteDataSet("FETCH_PUT_ADJUSMENT_EXCEPTION_DETAILS", lookupCode, UnderlyingId);

                string strHTML = "<table border='1'>";
                strHTML += "<tr style='background-color:#BCBCBC;color:Black;'>";

                for (int intCol = 0; intCol < dsResult1.Tables[0].Columns.Count; intCol++)
                {
                    strHTML += "<td>";
                    strHTML += dsResult1.Tables[0].Columns[intCol].Caption;
                    strHTML += "</td>";
                }
                strHTML += "</tr>";

                for (int intRow = 0; intRow < dsResult1.Tables[0].Rows.Count; intRow++)
                {
                    strHTML += "<tr>";
                    for (int intCol = 0; intCol < dsResult1.Tables[0].Columns.Count; intCol++)
                    {
                        strHTML += "<td>";
                        strHTML += dsResult1.Tables[0].Rows[intRow][intCol];
                        strHTML += "</td>";
                    }
                    strHTML += "</tr>";
                }
                strHTML += "</table>";

                if (dsResult1.Tables[0].Rows.Count == 0)
                    strHTML = "No Records Found";

                return Json(strHTML, JsonRequestBehavior.AllowGet);

            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "FetchPutAdjusmentSurface", objUserMaster.UserID);
                return Json("");
            }
        }


        public JsonResult GetIvVersion(int LookUpID, int UnderlyingId)
        {
            if (UnderlyingId == 0)
            {
                UnderlyingId = Convert.ToInt32(Session["UnderlyingID"]);
            }
            try
            {
                int version = 0;
                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("GET_IV_VERSION", LookUpID, UnderlyingId);

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count == 1 && dsResult.Tables[0].Rows[0]["VERSION"].ToString() != string.Empty)
                {
                    version = Convert.ToInt32(dsResult.Tables[0].Rows[0]["VERSION"]);
                }
                else
                {
                    version = 0;
                }

                return Json(version, JsonRequestBehavior.AllowGet);
            }

            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "GetIvVersion", objUserMaster.UserID);
                return Json("");
            }

            //return Json(objResponseDetails, JsonRequestBehavior.AllowGet);
        }

        public JsonResult GetRollCostVersion(int LookUpID, int UnderlyingId)
        {
            if (UnderlyingId == 0)
            {
                UnderlyingId = Convert.ToInt32(Session["UnderlyingID"]);
            }
            try
            {
                int version = 0;
                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("GET_RC_VERSION", LookUpID, UnderlyingId);

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count == 1 && dsResult.Tables[0].Rows[0]["VERSION"].ToString() != string.Empty)
                {
                    version = Convert.ToInt32(dsResult.Tables[0].Rows[0]["VERSION"]);
                }
                else
                {
                    version = 0;
                }

                return Json(version, JsonRequestBehavior.AllowGet);
            }

            catch (Exception ex)
            {

                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "GetRollCostVersion", objUserMaster.UserID);
                return Json("");
            }

            //return Json(objResponseDetails, JsonRequestBehavior.AllowGet);
        }

        public JsonResult GetLocalVocalityVersion(int LookUpID, int UnderlyingId)
        {
            if (UnderlyingId == 0)
            {
                UnderlyingId = Convert.ToInt32(Session["UnderlyingID"]);
            }
            try
            {
                int version = 0;
                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("GET_LVS_VERSION", LookUpID, UnderlyingId);

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count == 1 && dsResult.Tables[0].Rows[0]["VERSION"].ToString() != string.Empty)
                {
                    version = Convert.ToInt32(dsResult.Tables[0].Rows[0]["VERSION"]);
                }
                else
                {
                    version = 0;
                }

                return Json(version, JsonRequestBehavior.AllowGet);
            }

            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "GetLocalVocalityVersion", objUserMaster.UserID);
                return Json("");
            }

            //return Json(objResponseDetails, JsonRequestBehavior.AllowGet);
        }


        public JsonResult GetAdjusmentVersion(int LookUpID, int UnderlyingId)
        {
            if (UnderlyingId == 0)
            {
                UnderlyingId = Convert.ToInt32(Session["UnderlyingID"]);
            }
            try
            {
                int version = 0;
                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("GET_ADJUSTMENTSURFACE_VERSION", LookUpID, UnderlyingId);

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count == 1 && dsResult.Tables[0].Rows[0]["VERSION"].ToString() != string.Empty)
                {
                    version = Convert.ToInt32(dsResult.Tables[0].Rows[0]["VERSION"]);
                }
                else
                {
                    version = 0;
                }

                return Json(version, JsonRequestBehavior.AllowGet);
            }

            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "GetAdjusmentVersion", objUserMaster.UserID);
                return Json("");
            }

            //return Json(objResponseDetails, JsonRequestBehavior.AllowGet);
        }

        public JsonResult GetAdjusmentExceptionVersion(int LookUpID, int UnderlyingId)
        {
            if (UnderlyingId == 0)
            {
                UnderlyingId = Convert.ToInt32(Session["UnderlyingID"]);
            }
            try
            {
                int version = 0;
                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("GET_ADJUSTMENT_EXCEPTION_VERSION", LookUpID, UnderlyingId);

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count == 1 && dsResult.Tables[0].Rows[0]["VERSION"].ToString() != string.Empty)
                {
                    version = Convert.ToInt32(dsResult.Tables[0].Rows[0]["VERSION"]);
                }
                else
                {
                    version = 0;
                }

                return Json(version, JsonRequestBehavior.AllowGet);
            }

            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "GetAdjusmentVersion", objUserMaster.UserID);
                return Json("");
            }

            //return Json(objResponseDetails, JsonRequestBehavior.AllowGet);
        }

        public JsonResult GetCallThresholdStrikeVersion(int LookUpID, int UnderlyingId)
        {
            if (UnderlyingId == 0)
            {
                UnderlyingId = Convert.ToInt32(Session["UnderlyingID"]);
            }
            try
            {
                int version = 0;
                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("GET_CALL_THRESHOLD_VERSION", LookUpID, UnderlyingId);

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count == 1 && dsResult.Tables[0].Rows[0]["VERSION"].ToString() != string.Empty)
                {
                    version = Convert.ToInt32(dsResult.Tables[0].Rows[0]["VERSION"]);
                }
                else
                {
                    version = 0;
                }

                return Json(version, JsonRequestBehavior.AllowGet);
            }

            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "GetCallThresholdStrikeVersion", objUserMaster.UserID);
                return Json("");
            }

            //return Json(objResponseDetails, JsonRequestBehavior.AllowGet);
        }

        public JsonResult GetCallMinimumVersion(int LookUpID, int UnderlyingId)
        {
            if (UnderlyingId == 0)
            {
                UnderlyingId = Convert.ToInt32(Session["UnderlyingID"]);
            }
            try
            {
                int version = 0;
                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("GET_CALL_MINIMUM_VERSION", LookUpID, UnderlyingId);

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count == 1 && dsResult.Tables[0].Rows[0]["VERSION"].ToString() != string.Empty)
                {
                    version = Convert.ToInt32(dsResult.Tables[0].Rows[0]["VERSION"]);
                }
                else
                {
                    version = 0;
                }
                return Json(version, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "GetCallMinimumVersion", objUserMaster.UserID);
                return Json("");
            }

            //return Json(objResponseDetails, JsonRequestBehavior.AllowGet);
        }

        public JsonResult GetSkewAdjustmentVersion(int LookUpID, int UnderlyingId)
        {
            if (UnderlyingId == 0)
            {
                UnderlyingId = Convert.ToInt32(Session["UnderlyingID"]);
            }
            try
            {
                int version = 0;
                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("GET_PUT_SKEW_VERSION", LookUpID, UnderlyingId);

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count == 1 && dsResult.Tables[0].Rows[0]["VERSION"].ToString() != string.Empty)
                {
                    version = Convert.ToInt32(dsResult.Tables[0].Rows[0]["VERSION"]);
                }
                else
                {
                    version = 0;
                }

                return Json(version, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "GetSkewAdjustmentVersion", objUserMaster.UserID);
                return Json("");
            }

            //return Json(objResponseDetails, JsonRequestBehavior.AllowGet);
        }

        public JsonResult GetMinimum4PSVersion(int LookUpID, int UnderlyingId)
        {
            if (UnderlyingId == 0)
            {
                UnderlyingId = Convert.ToInt32(Session["UnderlyingID"]);
            }
            try
            {
                int version = 0;
                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("GET_PUT_MINIMUM_VERSION", LookUpID, UnderlyingId);

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count == 1 && dsResult.Tables[0].Rows[0]["VERSION"].ToString() != string.Empty)
                {
                    version = Convert.ToInt32(dsResult.Tables[0].Rows[0]["VERSION"]);
                }
                else
                {
                    version = 0;
                }
                return Json(version, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "GetMinimum4PSVersion", objUserMaster.UserID);
                return Json("");
            }


            //return Json(objResponseDetails, JsonRequestBehavior.AllowGet);
        }

        public JsonResult FetchUnderlyingList()
        {
            try
            {
                List<Underlying> UnderlyingList = new List<Underlying>();

                ObjectResult<UnderlyingListResult> objUnderlyingListResult;

                objUnderlyingListResult = objSP_PRICINGEntities.SP_FETCH_UNDERLYING_DETAILS();
                var UnderlyingListData = objUnderlyingListResult.ToList();
                return Json(UnderlyingListData, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "FetchUnderlyingList", objUserMaster.UserID);
                return Json("");
            }
        }

        public JsonResult FetchUnderlyingListBasketInstruments()
        {
            try
            {
                DataTable dtData = (DataTable)Session["BasketCorrelationFileDetails"];

                List<Underlying> UnderlyingList = new List<Underlying>();

                foreach (DataRow dr in dtData.Rows)
                {
                    Underlying objUnderlying = new Underlying();

                    objUnderlying.UnderlyingID = 0;
                    objUnderlying.UnderlyingShortName = Convert.ToString(dr["UNDERLYING_1"]);

                    UnderlyingList.Add(objUnderlying);
                }

                foreach (DataRow dr in dtData.Rows)
                {
                    Underlying objUnderlying = new Underlying();

                    objUnderlying.UnderlyingID = 0;
                    objUnderlying.UnderlyingShortName = Convert.ToString(dr["UNDERLYING_2"]);

                    UnderlyingList.Add(objUnderlying);
                }

                var UnderlyingListData = UnderlyingList.Select(r => new { r.UnderlyingID, r.UnderlyingShortName }).Distinct();
                return Json(UnderlyingListData, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "FetchUnderlyingListBasketInstruments", objUserMaster.UserID);
                return Json("");
            }
        }

        public ActionResult AutoCompleteQuoteID(string term)
        {
            try
            {
                List<UnderlyingIV> UnderlyingIVList = new List<UnderlyingIV>();

                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("SP_FETCH_UNDERLYING_IV_DETAILS", "ALL", "ALL", "ALL", "ALL", "1900-01-01", "ALL", "ALL", "ALL", "1900-01-01", "2900-01-01");

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in dsResult.Tables[0].Rows)
                    {
                        UnderlyingIV obj = new UnderlyingIV();

                        obj.ProductCode = Convert.ToString(dr["ProductCode"]);
                        obj.UnderlyingName = Convert.ToString(dr["UnderlyingName"]);
                        obj.StrikeMultiplier = Convert.ToDouble(dr["StrikeMultiplier"]);
                        obj.InstrumentType = Convert.ToString(dr["InstrumentType"]);
                        obj.ExpiryDateString = Convert.ToString(dr["ExpiryDate"]);
                        obj.Direction = Convert.ToString(dr["Direction"]);


                        UnderlyingIVList.Add(obj);
                    }
                }

                var DistinctItems = UnderlyingIVList.GroupBy(x => x.ProductCode).Select(y => y.First());

                var result = (from objRuleList in DistinctItems
                              where objRuleList.ProductCode.ToLower().StartsWith(term.ToLower())
                              select objRuleList);

                return Json(result);

            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "AutoCompleteQuoteID", objUserMaster.UserID);
                Session["ErrorData"] = ex.Message;
                return RedirectToAction("ErrorPage", "Login");
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
    }
}
