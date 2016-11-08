using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using SPPricing.Models;
using System.Data;
using System.Data.Objects;
using Newtonsoft.Json;
using OfficeOpenXml;
using System.IO;
using System.Web.Script.Serialization;
using DotNet.Highcharts;
using DotNet.Highcharts.Enums;
using DotNet.Highcharts.Options;
using DotNet.Highcharts.Helpers;
using CRYPTOGRAPHY;
using System.Net.Http;
using System.Net;
using System.Diagnostics;
using System.Data.SqlClient;
using Microsoft.Win32;

namespace SPPricing.Controllers
{
    public class BlackscholesPricersController : Controller
    {
        //
        // GET: /BlackscholesPricers/

        SP_PRICINGEntities objSP_PRICINGEntities = new SP_PRICINGEntities();

        public ActionResult Index()
        {
            return View();
        }

        #region Fixed Coupon
        [HttpGet]
        public ActionResult FixedCoupon(string ProductID, bool IsQuotron = false)
        {
            LoginController objLoginController = new LoginController();

            try
            {
                if (ValidateSession())
                {
                    FixedCoupon objFixedCoupon = new FixedCoupon();

                    #region Menu Access By on Role

                    Int32 intResult = 0;
                    // bool PPorNonPP = false;

                    //objFixedCoupon.IsIRR = true;

                    UserMaster objUserMaster = new UserMaster();
                    objUserMaster = (UserMaster)Session["LoggedInUser"];

                    var Result = objSP_PRICINGEntities.VALIDATE_MENU_BY_ROLE(objUserMaster.RoleID, "BFC");
                    intResult = Convert.ToInt32(Result.SingleOrDefault());

                    if (intResult == 0)
                        return RedirectToAction("UserNotAuthorize", "Login");

                    #endregion

                    objFixedCoupon.EntityID = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["DefaultEntityID"]);
                    objFixedCoupon.IsSecuredID = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["DefaultIsSecuredID"]);

                    if (ProductID != "" && ProductID != null)
                    {
                        ObjectResult<FixedCouponEditResult> objFixedCouponEditResult = objSP_PRICINGEntities.FETCH_FIXED_COUPON_EDIT_DETAILS(ProductID);
                        List<FixedCouponEditResult> FixedCouponEditResultList = objFixedCouponEditResult.ToList();

                        General.ReflectSingleData(objFixedCoupon, FixedCouponEditResultList[0]);
                    }
                    else
                    {
                        objFixedCoupon.IsIRR = true;
                        objFixedCoupon.IsRedemptionPeriodMonth = true;
                    }

                    if (Session["FixedCouponCopyQuote"] != null)
                    {
                        ViewBag.Message = "Quote is copied";
                        objFixedCoupon = (FixedCoupon)Session["FixedCouponCopyQuote"];

                        ObjectResult<FixedCouponEditResult> objFixedCouponEditResult = objSP_PRICINGEntities.FETCH_FIXED_COUPON_EDIT_DETAILS(objFixedCoupon.ProductID);
                        List<FixedCouponEditResult> FixedCouponEditResultList = objFixedCouponEditResult.ToList();
                        FixedCoupon oFixedCoupon = new FixedCoupon();

                        if (FixedCouponEditResultList != null && FixedCouponEditResultList.Count > 0)
                            General.ReflectSingleData(oFixedCoupon, FixedCouponEditResultList[0]);

                        objFixedCoupon.ParentProductID = objFixedCoupon.ProductID;
                        objFixedCoupon.ProductID = "";
                        objFixedCoupon.Status = "";
                        objFixedCoupon.SaveStatus = "";
                        objFixedCoupon.IsIRR = oFixedCoupon.IsIRR;
                        objFixedCoupon.IsRedemptionPeriodMonth = oFixedCoupon.IsRedemptionPeriodMonth;

                        //-------------Added by Shweta on 22nd July 2016 to Fetch Latest Deployment Rate------------START--------
                        string strDeploymentRate = "";
                        var DeploymentRate = objSP_PRICINGEntities.SP_FETCH_PRICING_DEPLOYMENT_RATE(Convert.ToInt32(objFixedCoupon.RedemptionPeriodDays), Convert.ToInt32(objFixedCoupon.EntityID), Convert.ToInt32(objFixedCoupon.IsSecuredID));
                        strDeploymentRate = Convert.ToString(DeploymentRate.SingleOrDefault());
                        objFixedCoupon.DeploymentRate = Convert.ToDouble(strDeploymentRate);
                        //-------------Added by Shweta on 22nd July 2016 to Fetch Latest Deployment Rate------------END----------
                    }

                    else if (Session["FixedCouponChildQuote"] != null)
                    {
                        ViewBag.Message = true;
                        objFixedCoupon = (FixedCoupon)Session["FixedCouponChildQuote"];

                        ObjectResult<FixedCouponEditResult> objFixedCouponEditResult = objSP_PRICINGEntities.FETCH_FIXED_COUPON_EDIT_DETAILS("");
                        List<FixedCouponEditResult> FixedCouponEditResultList = objFixedCouponEditResult.ToList();
                        FixedCoupon oFixedCoupon = new FixedCoupon();

                        if (FixedCouponEditResultList != null && FixedCouponEditResultList.Count > 0)
                            General.ReflectSingleData(oFixedCoupon, FixedCouponEditResultList[0]);

                        objFixedCoupon.ParentProductID = objFixedCoupon.ProductID;
                        objFixedCoupon.ProductID = "";
                        objFixedCoupon.Status = oFixedCoupon.Status;
                        objFixedCoupon.SaveStatus = oFixedCoupon.SaveStatus;
                        objFixedCoupon.IsChildQuote = true;
                    }
                    else if (Session["CancelQuote"] != null)
                    {
                        ViewBag.Message = "Quote is copied";
                        objFixedCoupon = (FixedCoupon)Session["CancelQuote"];

                        ObjectResult<FixedCouponEditResult> objFixedCouponEditResult = objSP_PRICINGEntities.FETCH_FIXED_COUPON_EDIT_DETAILS(objFixedCoupon.ProductID);
                        List<FixedCouponEditResult> FixedCouponEditResultList = objFixedCouponEditResult.ToList();
                        FixedCoupon oFixedCoupon = new FixedCoupon();

                        if (FixedCouponEditResultList != null && FixedCouponEditResultList.Count > 0)
                            General.ReflectSingleData(oFixedCoupon, FixedCouponEditResultList[0]);

                        objFixedCoupon.Status = oFixedCoupon.Status;
                        objFixedCoupon.SaveStatus = oFixedCoupon.SaveStatus;

                        Session.Remove("CancelQuote");
                    }
                    else
                    {
                        Session.Remove("IsChildQuote");
                        Session.Remove("ParentProductID");
                    }

                    if (IsQuotron == true)
                    {
                        objFixedCoupon.IsQuotron = true;
                    }

                    if (Session["FixedCouponChildQuote"] == null && Session["FixedCouponCopyQuote"] == null)
                        objFixedCoupon.SaveStatus = "";

                    if (Session["FixedCouponCopyQuote"] != null)
                        Session.Remove("FixedCouponCopyQuote");

                    if (Session["FixedCouponChildQuote"] != null)
                        Session.Remove("FixedCouponChildQuote");

                    return View(objFixedCoupon);
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

                ClearFixedCouponSession();

                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "FixedCoupon Get", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        public ActionResult ManageFixedCoupon(string ProductID, string Distributor, string EdelweissBuiltIn, string DistributorBuiltIn, string TotalBuiltIn, string BuiltInAdjustment, string Remaining, string IRR, string IsIRR, string DeploymentRate, string CustomerDeploymentRate, string FixedCoupon, string OptionTenureMonth, string RedemptionPeriodMonth, string IsRedemptionPeriodMonth, string RedemptionPeriodDays, string SalesComments, string TradingComments, string CouponScenario, string CopyProductID, string Entity, string IsSecured)
        {
            LoginController objLoginController = new LoginController();
            try
            {
                if (ValidateSession())
                {
                    UserMaster objUserMaster = new UserMaster();
                    objUserMaster = (UserMaster)Session["LoggedInUser"];

                    if (CustomerDeploymentRate == "")
                        CustomerDeploymentRate = "0";

                    string ParentProductID = "";
                    if (Session["ParentProductID"] != null)
                        ParentProductID = (string)Session["ParentProductID"];

                    ObjectResult<ManageFixedCouponResult> objManageFixedCouponResult = objSP_PRICINGEntities.SP_MANAGE_FIXED_COUPON_DETAILS(ProductID, ParentProductID, Distributor, Convert.ToDouble(EdelweissBuiltIn), Convert.ToDouble(DistributorBuiltIn), Convert.ToDouble(BuiltInAdjustment), Convert.ToDouble(TotalBuiltIn), Convert.ToDouble(Remaining), Convert.ToDouble(IRR), Convert.ToBoolean(IsIRR), Convert.ToDouble(DeploymentRate), Convert.ToDouble(CustomerDeploymentRate), Convert.ToDouble(FixedCoupon), Convert.ToDouble(OptionTenureMonth), Convert.ToDouble(RedemptionPeriodMonth), Convert.ToBoolean(IsRedemptionPeriodMonth), Convert.ToInt32(RedemptionPeriodDays), SalesComments, TradingComments, CouponScenario, Convert.ToInt32(Entity), Convert.ToInt32(IsSecured), objUserMaster.UserID, CopyProductID);
                    List<ManageFixedCouponResult> ManageFixedCouponResultList = objManageFixedCouponResult.ToList();

                    Session.Remove("ParentProductID");

                    return Json(ManageFixedCouponResultList[0].ProductID);
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

                ClearFixedCouponSession();

                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "ManageFixedCoupon Get", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }

        }

        [HttpPost, ValidateInput(false)]
        public ActionResult FixedCoupon(string Command, FixedCoupon objFixedCoupon)
        {
            LoginController objLoginController = new LoginController();
            try
            {
                if (ValidateSession())
                {
                    if (Command == "ExportToExcel")
                    {
                        ExportFixedCoupon(objFixedCoupon);
                    }
                    else if (Command == "CopyQuote")
                    {
                        Session["FixedCouponCopyQuote"] = objFixedCoupon;

                        return RedirectToAction("FixedCoupon");
                    }
                    else if (Command == "CreateChildQuote")
                    {
                        Session.Remove("ParentProductID");
                        Session.Remove("IsChildQuote");

                        Session["ParentProductID"] = objFixedCoupon.ProductID;

                        objFixedCoupon.IsChildQuote = true;

                        Session["FixedCouponChildQuote"] = objFixedCoupon;
                        Session["IsChildQuote"] = objFixedCoupon.IsChildQuote;

                        return RedirectToAction("FixedCoupon");
                    }
                    else if (Command == "AddNewProduct")
                    {
                        var productID = objFixedCoupon.ProductID;
                        UserMaster objUserMaster = new UserMaster();
                        objUserMaster = (UserMaster)Session["LoggedInUser"];

                        EncryptDecrypt obj = new EncryptDecrypt();
                        var encryptedpaswd = obj.Encrypt(objUserMaster.Password, "SPPricing", CryptographyEngine.AlgorithmType.DES);
                        var ProductType = "PP";

                        var Url = "http://edemumnewuatvm4:63400/Login.aspx?UserId=" + objUserMaster.LoginName + "&Key=" + encryptedpaswd + "&ProductId=" + productID + "&ProductType=" + ProductType;
                        return Redirect(Url);
                    }
                    else if (Command == "Cancel")
                    {
                        Session["CancelQuote"] = objFixedCoupon;

                        return RedirectToAction("FixedCoupon");
                    }
                    else if (Command == "PricingInExcel")
                    {
                        objFixedCoupon.IsWorkingFileExport = OpenWorkingExcelFile("FC", objFixedCoupon.ProductID);

                        if (!objFixedCoupon.IsWorkingFileExport)
                            objFixedCoupon.WorkingFileStatus = "File Not Found";

                        return View(objFixedCoupon);
                    }

                    return View(objFixedCoupon);
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

                ClearFixedCouponSession();

                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "FixedCoupon Post", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }

        }

        public virtual void ExportFixedCoupon(FixedCoupon objFixedCoupon)
        {
            try
            {
                string strTemplateFilePath = Server.MapPath("~/Templates");
                string strTemplateFileName = strTemplateFilePath + "\\FixedCouponTemplate.xlsx";

                string strTargetFilePath = Server.MapPath("~/OutputFiles");
                string strTargetFileName = strTargetFilePath + "\\" + objFixedCoupon.ProductID + "_FixedCoupon.xlsx";

                if (System.IO.File.Exists(strTargetFileName))
                    System.IO.File.Delete(strTargetFileName);

                FileInfo objTemplateFileInfo = new FileInfo(strTemplateFileName);
                objTemplateFileInfo.CopyTo(strTargetFileName);

                FileInfo objTargetFileInfo = new FileInfo(strTargetFileName);

                using (var xlPackage = new ExcelPackage(objTargetFileInfo))
                {
                    var worksheet = xlPackage.Workbook.Worksheets["FixedCoupon"];

                    worksheet.Cell(1, 2).Value = objFixedCoupon.ProductID.ToString();
                    worksheet.Cell(1, 4).Value = objFixedCoupon.Distributor.ToString().ToUpper();

                    worksheet.Cell(2, 2).Formula = "=" + objFixedCoupon.EdelweissBuiltIn.ToString() + "%";
                    worksheet.Cell(2, 4).Formula = "=" + objFixedCoupon.DistributorBuiltIn.ToString() + "%";
                    worksheet.Cell(2, 6).Formula = "=B2+D2+H2";
                    worksheet.Cell(2, 8).Formula = "=" + objFixedCoupon.BuiltInAdjustment.ToString() + "%";

                    worksheet.Cell(3, 2).Formula = "=((POWER((1+D3),(12/B4))-1) * 100) %";
                    worksheet.Cell(3, 4).Formula = "=" + objFixedCoupon.FixedCouponValue.ToString() + "%";

                    worksheet.Cell(4, 2).Formula = "=ROUND(D4/30.417,0)";
                    worksheet.Cell(4, 4).Value = objFixedCoupon.RedemptionPeriodDays.ToString();

                    #region Get Entity Name
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
                    LookupMaster objLookupMasterEntity = EntityList.Find(delegate(LookupMaster oLookupMaster) { return oLookupMaster.LookupID == objFixedCoupon.EntityID; });
                    worksheet.Cell(5, 2).Value = objLookupMasterEntity.LookupDescription;
                    #endregion

                    #region Get Is Secured
                    objLookupResult = null;
                    LookupResultList = null;
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
                    LookupMaster objLookupMasterIsSecured = IsSecuredList.Find(delegate(LookupMaster oLookupMaster) { return oLookupMaster.LookupID == objFixedCoupon.IsSecuredID; });
                    worksheet.Cell(5, 4).Value = objLookupMasterIsSecured.LookupDescription;
                    #endregion

                    worksheet.Cell(5, 6).Formula = "=" + objFixedCoupon.DeploymentRate.ToString() + "%";
                    worksheet.Cell(5, 8).Formula = "=" + objFixedCoupon.CustomDeploymentRate.ToString() + "%";

                    worksheet.Cell(6, 2).Formula = "=(100-(B2+D2)*100)-(100*(1+ROUND(D3, 4)))/(POWER((1+(IF(H5>0,H5,F5))),(B4/12)))";

                    if (objFixedCoupon.SalesComments != null)
                        worksheet.Cell(8, 2).Value = objFixedCoupon.SalesComments.ToString();
                    else
                        worksheet.Cell(8, 2).Value = "";

                    if (objFixedCoupon.TradingComments != null)
                        worksheet.Cell(9, 2).Value = objFixedCoupon.TradingComments.ToString();
                    else
                        worksheet.Cell(9, 2).Value = "";

                    if (objFixedCoupon.CouponScenario != null)
                        worksheet.Cell(10, 2).Value = objFixedCoupon.CouponScenario.ToString();
                    else
                        worksheet.Cell(10, 2).Value = "";


                    xlPackage.Save();
                }

                if (System.IO.File.Exists(strTargetFileName))
                {
                    FileInfo TemplateFile = new FileInfo(strTargetFileName);

                    Response.Clear();
                    Response.ClearHeaders();
                    Response.ClearContent();
                    Response.AddHeader("content-disposition", "attachment; filename=" + Path.GetFileName(strTargetFileName));
                    Response.AddHeader("Content-Type", "application/Excel");
                    Response.ContentType = "application/vnd.xls";
                    Response.AddHeader("Content-Length", TemplateFile.Length.ToString());
                    Response.WriteFile(TemplateFile.FullName);
                    Response.End();
                }
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];

                ClearFixedCouponSession();

                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "ExportFixedCoupon", objUserMaster.UserID);

            }
        }

        public JsonResult ExportFixedCouponWorkingFile(string ProductID, string Distributor, string EdelweissBuiltIn, string DistributorBuiltIn, string TotalBuiltIn, string BuiltInAdjustment, string Remaining, string IRR, string IsIRR, string DeploymentRate, string CustomerDeploymentRate, string FixedCoupon, string OptionTenureMonth, string RedemptionPeriodMonth, string IsRedemptionPeriodMonth, string RedemptionPeriodDays, string SalesComments, string TradingComments, string CouponScenario, string Entity, string IsSecured)
        {
            try
            {
                string strTemplateFilePath = Server.MapPath("~/Templates");
                string strTemplateFileName = strTemplateFilePath + "\\FixedCouponTemplateWorkingFile.xlsx";

                //string strTargetFilePath = Server.MapPath("~/WorkingFiles");
                string strTargetFilePath = System.Configuration.ConfigurationManager.AppSettings["WorkingFilePath"];
                string strTargetFileName = strTargetFilePath + "\\" + ProductID + "_FixedCoupon.xlsx";

                if (System.IO.File.Exists(strTargetFileName))
                    System.IO.File.Delete(strTargetFileName);

                FileInfo objTemplateFileInfo = new FileInfo(strTemplateFileName);
                objTemplateFileInfo.CopyTo(strTargetFileName);

                FileInfo objTargetFileInfo = new FileInfo(strTargetFileName);

                using (var xlPackage = new ExcelPackage(objTargetFileInfo))
                {
                    var worksheet = xlPackage.Workbook.Worksheets["FixedCoupon"];

                    worksheet.Cell(1, 2).Value = ProductID.ToString();
                    worksheet.Cell(1, 4).Value = Distributor.ToString().ToUpper();

                    worksheet.Cell(2, 2).Formula = "=" + EdelweissBuiltIn.ToString() + "%";
                    worksheet.Cell(2, 4).Formula = "=" + DistributorBuiltIn.ToString() + "%";
                    worksheet.Cell(2, 6).Formula = "=B2+D2+H2";
                    worksheet.Cell(2, 8).Formula = "=" + BuiltInAdjustment.ToString() + "%";

                    if (IsIRR.ToUpper() == "TRUE")
                    {
                        worksheet.Cell(3, 2).Formula = "=" + IRR.ToString() + "%";
                        worksheet.Cell(3, 4).Formula = "=(POWER((1 + B3), (D4 / 365)) - 1)*100%";
                    }
                    else
                    {
                        worksheet.Cell(3, 2).Formula = "=((POWER((1+ROUND(D3,4)),(12/ROUND(B4,2)))-1) * 100) %";
                        worksheet.Cell(3, 4).Formula = "=" + FixedCoupon.ToString() + "%";
                    }

                    if (IsRedemptionPeriodMonth.ToUpper() == "TRUE")
                    {
                        worksheet.Cell(4, 2).Formula = RedemptionPeriodMonth;
                        worksheet.Cell(4, 4).Formula = "=ROUND(B4*30.417, 0)";
                    }
                    else
                    {
                        worksheet.Cell(4, 2).Formula = "=ROUND(D4/30.417,2)";
                        worksheet.Cell(4, 4).Formula = RedemptionPeriodDays.ToString();
                    }

                    #region Get Entity Name
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
                    LookupMaster objLookupMasterEntity = EntityList.Find(delegate(LookupMaster oLookupMaster) { return oLookupMaster.LookupID == Convert.ToInt32(Entity); });
                    worksheet.Cell(5, 2).Value = objLookupMasterEntity.LookupDescription;
                    #endregion

                    #region Get Is Secured
                    objLookupResult = null;
                    LookupResultList = null;
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
                    LookupMaster objLookupMasterIsSecured = IsSecuredList.Find(delegate(LookupMaster oLookupMaster) { return oLookupMaster.LookupID == Convert.ToInt32(IsSecured); });
                    worksheet.Cell(5, 4).Value = objLookupMasterIsSecured.LookupDescription;
                    #endregion

                    worksheet.Cell(5, 6).Formula = "=" + DeploymentRate.ToString() + "%";

                    if (CustomerDeploymentRate == "")
                        CustomerDeploymentRate = "0";
                    worksheet.Cell(5, 8).Formula = "=" + CustomerDeploymentRate.ToString() + "%";

                    worksheet.Cell(6, 2).Formula = "=(100-(B2+D2)*100)-(100*(1+ROUND(D3,4)))/(POWER((1+(IF(H5>0,H5,F5))),(B4/12)))";

                    if (SalesComments != null)
                        worksheet.Cell(8, 2).Value = SalesComments.ToString();
                    else
                        worksheet.Cell(8, 2).Value = "";

                    if (TradingComments != null)
                        worksheet.Cell(9, 2).Value = TradingComments.ToString();
                    else
                        worksheet.Cell(9, 2).Value = "";

                    if (CouponScenario != null)
                        worksheet.Cell(10, 2).Value = CouponScenario.ToString();
                    else
                        worksheet.Cell(10, 2).Value = "";

                    xlPackage.Save();

                    return Json("");
                }

            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];

                ClearFixedCouponSession();

                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "ExportFixedCoupon", objUserMaster.UserID);
                return Json("");
            }
        }
        #endregion

        #region Fixed Coupon MLD
        [HttpGet]
        public ActionResult FixedCouponMLD(string ProductID, string GenerateGraph, bool IsQuotron = false)
        {
            LoginController objLoginController = new LoginController();
            try
            {
                if (ValidateSession())
                {
                    FixedCouponMLD objFixedCouponMLD = new FixedCouponMLD();

                    #region Menu Access By on Role

                    Int32 intResult = 0;
                    // bool PPorNonPP = false;

                    UserMaster objUserMaster = new UserMaster();
                    objUserMaster = (UserMaster)Session["LoggedInUser"];

                    var Result = objSP_PRICINGEntities.VALIDATE_MENU_BY_ROLE(objUserMaster.RoleID, "BFCM");
                    intResult = Convert.ToInt32(Result.SingleOrDefault());

                    if (intResult == 0)
                        return RedirectToAction("UserNotAuthorize", "Login");

                    #endregion

                    #region Bind Underlying List
                    List<Underlying> UnderlyingList = new List<Underlying>();

                    ObjectResult<UnderlyingListResult> objUnderlyingListResult;
                    List<UnderlyingListResult> UnderlyingListResultList = new List<UnderlyingListResult>();

                    objUnderlyingListResult = objSP_PRICINGEntities.SP_FETCH_UNDERLYING_DETAILS();
                    UnderlyingListResultList = objUnderlyingListResult.ToList();

                    foreach (UnderlyingListResult oUnderlyingListResult in UnderlyingListResultList)
                    {
                        Underlying objUnderlying = new Underlying();
                        General.ReflectSingleData(objUnderlying, oUnderlyingListResult);

                        UnderlyingList.Add(objUnderlying);
                    }
                    objFixedCouponMLD.UnderlyingList = UnderlyingList;

                    //--Set default underlying--Added by Shweta on 3rd May 2016------------START--------------------
                    string strDefaultUnderlying = System.Configuration.ConfigurationManager.AppSettings["DefaultUnderlying"].ToUpper();
                    Underlying objDefaulyUnderlying = objFixedCouponMLD.UnderlyingList.Find(delegate(Underlying oUnderlying) { return oUnderlying.UnderlyingShortName == strDefaultUnderlying; });
                    objFixedCouponMLD.UnderlyingID = objDefaulyUnderlying.UnderlyingID;

                    objFixedCouponMLD.EntityID = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["DefaultEntityID"]);
                    objFixedCouponMLD.IsSecuredID = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["DefaultIsSecuredID"]);
                    //--Set default underlying--Added by Shweta on 3rd May 2016------------END----------------------
                    #endregion

                    if (ProductID != "" && ProductID != null)
                    {
                        ObjectResult<FixedMLDEditResult> objFixedCouponMLDEditResult = objSP_PRICINGEntities.FETCH_FIXED_COUPON_MLD_EDIT_DETAILS(ProductID);
                        List<FixedMLDEditResult> FixedCouponMLDEditResultList = objFixedCouponMLDEditResult.ToList();

                        General.ReflectSingleData(objFixedCouponMLD, FixedCouponMLDEditResultList[0]);

                        DataSet dsResult = new DataSet();
                        dsResult = General.ExecuteDataSet("FETCH_UNDERLYING_CREATION_BYID", objFixedCouponMLD.UnderlyingID);

                        if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                        {
                            ViewBag.UnderlyingShortName = Convert.ToString(dsResult.Tables[0].Rows[0]["UNDERLYING_SHORTNAME"]);
                        }
                    }
                    else
                    {
                        objFixedCouponMLD.IsCouponAboveIRR = true;
                        objFixedCouponMLD.IsCouponBelowIRR = true;
                        objFixedCouponMLD.IsRedemptionPeriodMonth = true;
                    }

                    if (GenerateGraph == "GenerateGraph")
                    {
                        objFixedCouponMLD = (FixedCouponMLD)TempData["FixedMLDGraph"];
                        ObjectResult<FixedMLDEditResult> objFixedMLDEditResult = objSP_PRICINGEntities.FETCH_FIXED_COUPON_MLD_EDIT_DETAILS(objFixedCouponMLD.ProductID);
                        List<FixedMLDEditResult> FixedMLDEditResultList = objFixedMLDEditResult.ToList();
                        FixedCouponMLD oFixedCouponMLD = new FixedCouponMLD();
                        General.ReflectSingleData(oFixedCouponMLD, FixedMLDEditResultList[0]);

                        objFixedCouponMLD.Status = oFixedCouponMLD.Status;
                        // objFixedCouponMLD.SaveStatus = oFixedCouponMLD.SaveStatus;
                        return GenerateFixedCouponMLDGraph(objFixedCouponMLD);
                    }

                    else if (Session["FixedCouponMLDCopyQuote"] != null)
                    {
                        ViewBag.Message = "Quote is copied";
                        objFixedCouponMLD = (FixedCouponMLD)Session["FixedCouponMLDCopyQuote"];
                        objFixedCouponMLD.UnderlyingID = Convert.ToInt32(Session["UnderlyingID"]);

                        ObjectResult<FixedMLDEditResult> objFixedMLDEditResult = objSP_PRICINGEntities.FETCH_FIXED_COUPON_MLD_EDIT_DETAILS(objFixedCouponMLD.ProductID);
                        List<FixedMLDEditResult> FixedMLDEditResultList = objFixedMLDEditResult.ToList();
                        FixedCouponMLD oFixedCouponMLD = new FixedCouponMLD();

                        if (FixedMLDEditResultList != null && FixedMLDEditResultList.Count > 0)
                            General.ReflectSingleData(oFixedCouponMLD, FixedMLDEditResultList[0]);

                        objFixedCouponMLD.ParentProductID = objFixedCouponMLD.ProductID;
                        objFixedCouponMLD.ProductID = "";
                        objFixedCouponMLD.Status = "";
                        objFixedCouponMLD.SaveStatus = "";
                        objFixedCouponMLD.IsCouponAboveIRR = oFixedCouponMLD.IsCouponAboveIRR;
                        objFixedCouponMLD.IsCouponBelowIRR = oFixedCouponMLD.IsCouponBelowIRR;
                        objFixedCouponMLD.IsRedemptionPeriodMonth = oFixedCouponMLD.IsRedemptionPeriodMonth;

                        //-------------Added by Shweta on 22nd July 2016 to Fetch Latest Deployment Rate------------START--------
                        string strDeploymentRate = "";
                        var DeploymentRate = objSP_PRICINGEntities.SP_FETCH_PRICING_DEPLOYMENT_RATE(Convert.ToInt32(objFixedCouponMLD.RedemptionPeriodDays), objFixedCouponMLD.EntityID, objFixedCouponMLD.IsSecuredID);
                        strDeploymentRate = Convert.ToString(DeploymentRate.SingleOrDefault());
                        objFixedCouponMLD.DeploymentRate = Convert.ToDouble(strDeploymentRate);
                        //-------------Added by Shweta on 22nd July 2016 to Fetch Latest Deployment Rate------------END----------
                    }

                    else if (Session["FixedCouponMLDChildQuote"] != null)
                    {
                        ViewBag.Message = true;
                        objFixedCouponMLD = (FixedCouponMLD)Session["FixedCouponMLDChildQuote"];

                        objFixedCouponMLD.UnderlyingID = Convert.ToInt32(Session["UnderlyingID"]);

                        ObjectResult<FixedMLDEditResult> objFixedMLDEditResult = objSP_PRICINGEntities.FETCH_FIXED_COUPON_MLD_EDIT_DETAILS("");
                        List<FixedMLDEditResult> FixedMLDEditResultList = objFixedMLDEditResult.ToList();
                        FixedCouponMLD oFixedCouponMLD = new FixedCouponMLD();
                        if (FixedMLDEditResultList != null && FixedMLDEditResultList.Count > 0)
                            General.ReflectSingleData(oFixedCouponMLD, FixedMLDEditResultList[0]);

                        objFixedCouponMLD.ParentProductID = objFixedCouponMLD.ProductID;
                        objFixedCouponMLD.ProductID = "";
                        objFixedCouponMLD.Status = oFixedCouponMLD.Status;
                        objFixedCouponMLD.SaveStatus = oFixedCouponMLD.SaveStatus;
                        objFixedCouponMLD.IsChildQuote = true;
                    }
                    else if (Session["CancelQuote"] != null)
                    {
                        ViewBag.Message = "Quote is copied";
                        objFixedCouponMLD = (FixedCouponMLD)Session["CancelQuote"];

                        ObjectResult<FixedMLDEditResult> objFixedMLDEditResult = objSP_PRICINGEntities.FETCH_FIXED_COUPON_MLD_EDIT_DETAILS(objFixedCouponMLD.ProductID);
                        List<FixedMLDEditResult> FixedMLDEditResultList = objFixedMLDEditResult.ToList();
                        FixedCouponMLD oFixedCouponMLD = new FixedCouponMLD();
                        if (FixedMLDEditResultList != null && FixedMLDEditResultList.Count > 0)
                            General.ReflectSingleData(oFixedCouponMLD, FixedMLDEditResultList[0]);

                        objFixedCouponMLD.Status = oFixedCouponMLD.Status;
                        objFixedCouponMLD.SaveStatus = oFixedCouponMLD.SaveStatus;

                        Session.Remove("CancelQuote");
                    }
                    else
                    {
                        Session.Remove("IsChildQuoteMLD");
                        Session.Remove("ParentProductID");
                        Session.Remove("UnderlyingID");
                    }

                    if (IsQuotron == true)
                    {
                        objFixedCouponMLD.IsQuotron = true;
                    }

                    if (Session["FixedCouponMLDChildQuote"] == null && Session["FixedCouponMLDCopyQuote"] == null)
                        objFixedCouponMLD.SaveStatus = "";

                    if (Session["FixedCouponMLDCopyQuote"] != null)
                        Session.Remove("FixedCouponMLDCopyQuote");

                    if (Session["FixedCouponMLDChildQuote"] != null)
                        Session.Remove("FixedCouponMLDChildQuote");

                    if (ProductID == null)
                    {
                        objFixedCouponMLD.isGraphActive = false;
                        return View(objFixedCouponMLD);
                    }
                    else
                    {
                        return GenerateFixedCouponMLDGraph(objFixedCouponMLD);
                    }
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

                ClearFixedCouponMLDSession();

                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "FixedCouponMLD Get", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        private ActionResult GenerateFixedCouponMLDGraph(FixedCouponMLD objFixedCouponMLD)
        {
            LoginController objLoginController = new LoginController();
            try
            {
                if (ValidateSession())
                {
                    var transactionCounts = new List<Graph>();
                    transactionCounts = GenerateFixedMLDGraphCalculation(objFixedCouponMLD.Strike, objFixedCouponMLD.CouponBelowStrike, objFixedCouponMLD.CouponAboveStrike, objFixedCouponMLD.RedemptionPeriodDays);

                    //FixedCouponMLD obj = new FixedCouponMLD();

                    #region Pie Chart For FixedCouponMLD
                    var xDataMonths = transactionCounts.Select(i => i.Column1).ToArray();
                    var yDataCounts = transactionCounts.Select(i => new object[] { i.Column2 }).ToArray();
                    var yDataCounts1 = transactionCounts.Select(i => new object[] { i.Column3 }).ToArray();

                    var chart = new Highcharts("pie")
                        //define the type of chart 
                                .InitChart(new Chart { DefaultSeriesType = ChartTypes.Line })
                        //overall Title of the chart 
                                .SetTitle(new Title { Text = "Fixed Coupon MLD" })
                        ////small label below the main Title
                        //        .SetSubtitle(new Subtitle { Text = "Accounting" })
                        //load the X values
                        .SetXAxis(new XAxis { Title = new XAxisTitle { Text = "Underlying Returns" }, Categories = xDataMonths, Labels = new XAxisLabels { Step = 2 } })
                        //set the Y title
                                .SetYAxis(new YAxis { Title = new YAxisTitle { Text = "Product Returns" } })
                                .SetTooltip(new Tooltip
                                {
                                    //PointFormat = "{series.name}: <b>{point.percentage:.1f}%</b>",
                                    Enabled = true,
                                    Formatter = @"function() { return '<b>'+ this.series.name +'</b><br/>'+ this.x +': '+ this.y; }"
                                })
                                .SetPlotOptions(new PlotOptions
                                {
                                    Line = new PlotOptionsLine
                                    {
                                        DataLabels = new PlotOptionsLineDataLabels
                                        {
                                            Enabled = false
                                        },
                                        EnableMouseTracking = true
                                    }
                                })
                        //load the Y values 
                                .SetSeries(new[]
                    {
                        new Series {Name = "Coupon", Data = new Data(yDataCounts)}
                        //,//you can add more y data to create a second line
                        //new Series { Name = "Strike", Data = new Data(yDataCounts1) }
                    });
                    #endregion

                    if (Session["FixedCouponMLDCopyQuote"] != null)
                    {
                        ViewBag.Message = "Quote is copied";
                        objFixedCouponMLD = (FixedCouponMLD)Session["FixedCouponMLDCopyQuote"];
                        Session.Remove("FixedCouponMLDCopyQuote");
                    }

                    objFixedCouponMLD.FixedCouponMLDChart = chart;

                    return View(objFixedCouponMLD);
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

                ClearFixedCouponMLDSession();

                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "GenerateFixedCouponMLDGraph", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        [HttpPost, ValidateInput(false)]
        public ActionResult FixedCouponMLD(string Command, FixedCouponMLD objFixedCouponMLD, FormCollection objFormCollection)
        {
            LoginController objLoginController = new LoginController();
            try
            {
                #region Bind Underlying List
                List<Underlying> UnderlyingList = new List<Underlying>();

                ObjectResult<UnderlyingListResult> objUnderlyingListResult;
                List<UnderlyingListResult> UnderlyingListResultList = new List<UnderlyingListResult>();

                objUnderlyingListResult = objSP_PRICINGEntities.SP_FETCH_UNDERLYING_DETAILS();
                UnderlyingListResultList = objUnderlyingListResult.ToList();

                foreach (UnderlyingListResult oUnderlyingListResult in UnderlyingListResultList)
                {
                    Underlying objUnderlying = new Underlying();
                    General.ReflectSingleData(objUnderlying, oUnderlyingListResult);

                    UnderlyingList.Add(objUnderlying);
                }
                objFixedCouponMLD.UnderlyingList = UnderlyingList;
                #endregion
                if (ValidateSession())
                {
                    if (Command == "ExportToExcel")
                    {
                        ExportFixedCouponMLD(objFixedCouponMLD, objFormCollection);

                        return RedirectToAction("FixedCouponMLD");
                    }
                    else if (Command == "GenerateGraph")
                    {
                        objFixedCouponMLD.isGraphActive = true;

                        TempData["FixedMLDGraph"] = objFixedCouponMLD;
                        return RedirectToAction("FixedCouponMLD", new { GenerateGraph = "GenerateGraph" });
                    }
                    else if (Command == "CopyQuote")
                    {
                        Session["FixedCouponMLDCopyQuote"] = objFixedCouponMLD;
                        Session["UnderlyingID"] = objFixedCouponMLD.UnderlyingID;

                        return RedirectToAction("FixedCouponMLD");
                    }
                    else if (Command == "CreateChildQuote")
                    {
                        Session.Remove("ParentProductID");
                        Session.Remove("IsChildQuoteMLD");
                        Session.Remove("UnderlyingID");

                        Session["ParentProductID"] = objFixedCouponMLD.ProductID;
                        Session["UnderlyingID"] = objFixedCouponMLD.UnderlyingID;

                        objFixedCouponMLD.IsChildQuote = true;

                        Session["FixedCouponMLDChildQuote"] = objFixedCouponMLD;
                        Session["IsChildQuoteMLD"] = objFixedCouponMLD.IsChildQuote;

                        return RedirectToAction("FixedCouponMLD");
                    }
                    else if (Command == "AddNewProduct")
                    {
                        var productID = objFixedCouponMLD.ProductID;
                        UserMaster objUserMaster = new UserMaster();
                        objUserMaster = (UserMaster)Session["LoggedInUser"];

                        EncryptDecrypt obj = new EncryptDecrypt();
                        var encryptedpaswd = obj.Encrypt(objUserMaster.Password, "SPPricing", CryptographyEngine.AlgorithmType.DES);
                        var ProductType = "PP";

                        var Url = "http://edemumnewuatvm4:63400/Login.aspx?UserId=" + objUserMaster.LoginName + "&Key=" + encryptedpaswd + "&ProductId=" + productID + "&ProductType=" + ProductType;

                        // return Json(new { Url = Url });
                        //return Json(Url,JsonRequestBehavior.AllowGet);
                        return Redirect(Url);
                    }
                    else if (Command == "Cancel")
                    {
                        Session["CancelQuote"] = objFixedCouponMLD;

                        return RedirectToAction("FixedCouponMLD");
                    }
                    else if (Command == "PricingInExcel")
                    {
                        objFixedCouponMLD.IsWorkingFileExport = OpenWorkingExcelFile("FCM", objFixedCouponMLD.ProductID);

                        if (!objFixedCouponMLD.IsWorkingFileExport)
                            objFixedCouponMLD.WorkingFileStatus = "File Not Found";

                        return View(objFixedCouponMLD);
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

                ClearFixedCouponMLDSession();

                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "FixedCouponMLD Post", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        //public JsonResult GenerateFixedCouponMLDGraph(string Strike1, string Strike2, string Instrument, string Tenure, string UnderlyingID)
        //{
        //    try
        //    {
        //        ObjectResult<FinalIVRFResult> objFinalIVRFResult = objSP_PRICINGEntities.SP_FETCH_FINAL_IV_RF_VALUE(Convert.ToDouble(Strike1), Convert.ToDouble(Strike2), Instrument, Convert.ToInt32(Tenure), Convert.ToInt32(UnderlyingID));
        //        List<FinalIVRFResult> FinalIVRFResultList = objFinalIVRFResult.ToList();

        //        return Json(FinalIVRFResultList[0]);
        //    }
        //    catch (Exception ex)
        //    {
        //        UserMaster objUserMaster = new UserMaster();
        //        objUserMaster = (UserMaster)Session["LoggedInUser"];

        //        LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "GenerateFixedCouponMLDGraph", objUserMaster.UserID);
        //        return Json("");
        //    }
        //}

        public ActionResult ManageFixedCouponMLD(string ProductID, string Distributor, string EdelweissBuiltIn, string DistributorBuiltIn, string TotalBuiltIn, string BuiltInAdjustment, string Underlying, string Remaining, string Strike, string DeploymentRate, string CustomerDeploymentRate, string CouponAboveStrike, string CouponBelowStrike, string CouponAboveIRR, string IsCouponAboveIRR, string CouponBelowIRR, string IsCouponBelowIRR, string OptionTenureMonth, string RedemptionPeriodMonth, string IsRedemptionPeriodMonth, string RedemptionPeriodDays, string InitialAveragingMonth, string InitialAveragingDaysDiff, string FinalAveragingMonth, string FinalAveragingDaysDiff, string SalesComments, string TradingComments, string CouponScenario, string CopyProductID, string Entity, string IsSecured)
        {
            try
            {
                if (ValidateSession())
                {
                    UserMaster objUserMaster = new UserMaster();
                    objUserMaster = (UserMaster)Session["LoggedInUser"];

                    if (CustomerDeploymentRate == "")
                        CustomerDeploymentRate = "0";

                    string ParentProductID = "";
                    if (Session["ParentProductID"] != null)
                        ParentProductID = (string)Session["ParentProductID"];

                    ObjectResult<ManageFixedCouponMLDResult> objManageFixedCouponMLDResult = objSP_PRICINGEntities.SP_MANAGE_FIXED_COUPON_MLD_DETAILS(ProductID, ParentProductID, Distributor, Convert.ToDouble(EdelweissBuiltIn), Convert.ToDouble(DistributorBuiltIn), Convert.ToDouble(BuiltInAdjustment), Convert.ToDouble(TotalBuiltIn), Convert.ToInt32(Underlying), Convert.ToDouble(Remaining), Convert.ToDouble(Strike), Convert.ToDouble(DeploymentRate), Convert.ToDouble(CustomerDeploymentRate), Convert.ToDouble(CouponAboveStrike), Convert.ToDouble(CouponBelowStrike), Convert.ToDouble(CouponAboveIRR), Convert.ToBoolean(IsCouponAboveIRR), Convert.ToDouble(CouponBelowIRR), Convert.ToBoolean(IsCouponBelowIRR), Convert.ToInt32(OptionTenureMonth), Convert.ToDouble(RedemptionPeriodMonth), Convert.ToBoolean(IsRedemptionPeriodMonth), Convert.ToInt32(RedemptionPeriodDays), Convert.ToInt32(InitialAveragingMonth), Convert.ToInt32(InitialAveragingDaysDiff), Convert.ToInt32(FinalAveragingMonth), Convert.ToInt32(FinalAveragingDaysDiff), SalesComments, TradingComments, CouponScenario, Convert.ToInt32(Entity), Convert.ToInt32(IsSecured), objUserMaster.UserID, CopyProductID);
                    List<ManageFixedCouponMLDResult> ManageFixedCouponMLDResultList = objManageFixedCouponMLDResult.ToList();

                    Session.Remove("ParentProductID");

                    return Json(ManageFixedCouponMLDResultList[0].ProductID);
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

                ClearFixedCouponMLDSession();

                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "ManageFixedCouponMLD", objUserMaster.UserID);
                return Json("");
            }
        }

        public virtual void ExportFixedCouponMLD(FixedCouponMLD objFixedCouponMLD, FormCollection objFormCollection)
        {
            try
            {
                string strTemplateFilePath = Server.MapPath("~/Templates");
                string strTemplateFileName = strTemplateFilePath + "//FixedCouponMLDTemplate.xlsx";

                string strTargetFilePath = Server.MapPath("~/OutputFiles");
                string strTargetFileName = strTargetFilePath + "//" + objFixedCouponMLD.ProductID + "_FixedCouponMLD.xlsx";

                if (System.IO.File.Exists(strTargetFileName))
                    System.IO.File.Delete(strTargetFileName);

                FileInfo objTemplateFileInfo = new FileInfo(strTemplateFileName);
                objTemplateFileInfo.CopyTo(strTargetFileName);

                FileInfo objTargetFileInfo = new FileInfo(strTargetFileName);
                Underlying objUnderlying = objFixedCouponMLD.UnderlyingList.Find(delegate(Underlying oUnderlying) { return oUnderlying.UnderlyingID == objFixedCouponMLD.UnderlyingID; });

                using (var xlPackage = new ExcelPackage(objTargetFileInfo))
                {
                    var worksheet = xlPackage.Workbook.Worksheets["FixedCouponMLD"];

                    worksheet.Cell(1, 2).Value = objFixedCouponMLD.ProductID.ToString();
                    worksheet.Cell(1, 4).Value = objFixedCouponMLD.Distributor.ToString().ToUpper();
                    worksheet.Cell(1, 6).Value = objUnderlying.UnderlyingShortName;

                    worksheet.Cell(2, 2).Formula = "=" + objFixedCouponMLD.EdelweissBuiltIn.ToString() + "%";
                    worksheet.Cell(2, 4).Formula = "=" + objFixedCouponMLD.DistributorBuiltIn.ToString() + "%";
                    worksheet.Cell(2, 6).Formula = "=B2+D2+H2";
                    worksheet.Cell(2, 8).Formula = "=" + objFixedCouponMLD.BuiltInAdjustment.ToString() + "%";

                    //worksheet.Cell(4, 2).Formula = "=" + objFixedCouponMLD.FixedCouponValue.ToString() + "%";
                    //worksheet.Cell(4, 4).Formula = "=" + objFixedCouponMLD.Strike.ToString() + "%";

                    worksheet.Cell(3, 2).Value = objFixedCouponMLD.OptionTenureMonth.ToString();
                    worksheet.Cell(3, 4).Formula = "=ROUND(F3/30.417,0)";
                    worksheet.Cell(3, 6).Value = objFixedCouponMLD.RedemptionPeriodDays.ToString();

                    #region Get Entity Name
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
                    LookupMaster objLookupMasterEntity = EntityList.Find(delegate(LookupMaster oLookupMaster) { return oLookupMaster.LookupID == objFixedCouponMLD.EntityID; });
                    worksheet.Cell(4, 2).Value = objLookupMasterEntity.LookupDescription;
                    #endregion

                    #region Get Is Secured
                    objLookupResult = null;
                    LookupResultList = null;
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
                    LookupMaster objLookupMasterIsSecured = IsSecuredList.Find(delegate(LookupMaster oLookupMaster) { return oLookupMaster.LookupID == objFixedCouponMLD.IsSecuredID; });
                    worksheet.Cell(4, 4).Value = objLookupMasterIsSecured.LookupDescription;
                    #endregion
                    worksheet.Cell(4, 6).Formula = "=" + objFixedCouponMLD.DeploymentRate.ToString() + "%";
                    worksheet.Cell(4, 8).Formula = "=" + objFixedCouponMLD.CustomDeploymentRate.ToString() + "%";

                    worksheet.Cell(5, 2).Formula = "=((POWER((1+D5),(12/D3))-1) * 100) %";
                    worksheet.Cell(5, 4).Formula = "=" + objFixedCouponMLD.CouponAboveStrike.ToString() + "%";

                    worksheet.Cell(6, 2).Formula = "=((POWER((1+D6),(12/D3))-1) * 100) %";
                    worksheet.Cell(6, 4).Formula = "=" + objFixedCouponMLD.CouponBelowStrike.ToString() + "%";

                    worksheet.Cell(7, 2).Formula = "=(100-(B2+D2)*100)-(100*(1+ROUND(D5,4)))/(POWER((1+(IF(H4>0,H4,F4))),(D3/12)))";
                    worksheet.Cell(7, 4).Formula = "=" + objFixedCouponMLD.Strike.ToString() + "%";

                    worksheet.Cell(9, 2).Value = objFixedCouponMLD.InitialAveragingMonth.ToString();
                    worksheet.Cell(9, 4).Value = objFixedCouponMLD.InitialAveragingDaysDiff.ToString();
                    worksheet.Cell(9, 6).Value = objFixedCouponMLD.FinalAveragingMonth.ToString();
                    worksheet.Cell(9, 8).Value = objFixedCouponMLD.FinalAveragingDaysDiff.ToString();

                    if (objFixedCouponMLD.SalesComments != null)
                        worksheet.Cell(11, 2).Value = objFixedCouponMLD.SalesComments.ToString();
                    else
                        worksheet.Cell(11, 2).Value = "";

                    if (objFixedCouponMLD.TradingComments != null)
                        worksheet.Cell(12, 2).Value = objFixedCouponMLD.TradingComments.ToString();
                    else
                        worksheet.Cell(12, 2).Value = "";

                    if (objFixedCouponMLD.CouponScenario != null)
                        worksheet.Cell(13, 2).Value = objFixedCouponMLD.CouponScenario.ToString();
                    else
                        worksheet.Cell(13, 2).Value = "";

                    xlPackage.Save();
                }

                if (System.IO.File.Exists(strTargetFileName))
                {
                    FileInfo TemplateFile = new FileInfo(strTargetFileName);

                    Response.Clear();
                    Response.ClearHeaders();
                    Response.ClearContent();
                    Response.AddHeader("content-disposition", "attachment; filename=" + Path.GetFileName(strTargetFileName));
                    Response.AddHeader("Content-Type", "application/Excel");
                    Response.ContentType = "application/vnd.xls";
                    Response.AddHeader("Content-Length", TemplateFile.Length.ToString());
                    Response.WriteFile(TemplateFile.FullName);
                    Response.End();
                }
            }
            catch (Exception ex)
            {

                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];

                ClearFixedCouponMLDSession();

                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "ExportFixedCouponMLD", objUserMaster.UserID);

            }

        }

        public JsonResult ExportFixedCouponMLDWorkingFile(string ProductID, string Distributor, string EdelweissBuiltIn, string DistributorBuiltIn, string TotalBuiltIn, string BuiltInAdjustment, string Underlying, string Remaining, string Strike, string DeploymentRate, string CustomerDeploymentRate, string CouponAboveStrike, string CouponBelowStrike, string CouponAboveIRR, string IsCouponAboveIRR, string CouponBelowIRR, string IsCouponBelowIRR, string OptionTenureMonth, string RedemptionPeriodMonth, string IsRedemptionPeriodMonth, string RedemptionPeriodDays, string InitialAveragingMonth, string InitialAveragingDaysDiff, string FinalAveragingMonth, string FinalAveragingDaysDiff, string SalesComments, string TradingComments, string CouponScenario, string Entity, string IsSecured)
        {
            try
            {
                string strTemplateFilePath = Server.MapPath("~/Templates");
                string strTemplateFileName = strTemplateFilePath + "//FixedCouponMLDTemplateWorkingFile.xlsx";

                string strTargetFilePath = System.Configuration.ConfigurationManager.AppSettings["WorkingFilePath"];
                string strTargetFileName = strTargetFilePath + "//" + ProductID + "_FixedCouponMLD.xlsx";

                if (System.IO.File.Exists(strTargetFileName))
                    System.IO.File.Delete(strTargetFileName);

                FileInfo objTemplateFileInfo = new FileInfo(strTemplateFileName);
                objTemplateFileInfo.CopyTo(strTargetFileName);

                FileInfo objTargetFileInfo = new FileInfo(strTargetFileName);

                using (var xlPackage = new ExcelPackage(objTargetFileInfo))
                {
                    var worksheet = xlPackage.Workbook.Worksheets["FixedCouponMLD"];

                    worksheet.Cell(1, 2).Value = ProductID.ToString();
                    worksheet.Cell(1, 4).Value = Distributor.ToString().ToUpper();
                    worksheet.Cell(1, 6).Value = Underlying;

                    worksheet.Cell(2, 2).Formula = "=" + EdelweissBuiltIn.ToString() + "%";
                    worksheet.Cell(2, 4).Formula = "=" + DistributorBuiltIn.ToString() + "%";
                    worksheet.Cell(2, 6).Formula = "=B2+D2+H2";
                    worksheet.Cell(2, 8).Formula = "=" + BuiltInAdjustment.ToString() + "%";

                    worksheet.Cell(3, 2).Value = OptionTenureMonth.ToString();

                    if (IsRedemptionPeriodMonth.ToUpper() == "TRUE")
                    {
                        worksheet.Cell(3, 4).Formula = RedemptionPeriodMonth;
                        worksheet.Cell(3, 6).Formula = "=ROUND(D3*30.417, 0)";
                    }
                    else
                    {
                        worksheet.Cell(3, 4).Formula = "=ROUND(F3/30.417,2)";
                        worksheet.Cell(3, 6).Formula = RedemptionPeriodDays.ToString();
                    }

                    #region Get Entity Name
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
                    LookupMaster objLookupMasterEntity = EntityList.Find(delegate(LookupMaster oLookupMaster) { return oLookupMaster.LookupID == Convert.ToInt32(Entity); });
                    worksheet.Cell(4, 2).Value = objLookupMasterEntity.LookupDescription;
                    #endregion

                    #region Get Is Secured
                    objLookupResult = null;
                    LookupResultList = null;
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
                    LookupMaster objLookupMasterIsSecured = IsSecuredList.Find(delegate(LookupMaster oLookupMaster) { return oLookupMaster.LookupID == Convert.ToInt32(IsSecured); });
                    worksheet.Cell(4, 4).Value = objLookupMasterIsSecured.LookupDescription;
                    #endregion
                    worksheet.Cell(4, 6).Formula = "=" + DeploymentRate.ToString() + "%";

                    if (CustomerDeploymentRate == "")
                        CustomerDeploymentRate = "0";
                    worksheet.Cell(4, 8).Formula = "=" + CustomerDeploymentRate.ToString() + "%";

                    if (IsCouponAboveIRR.ToUpper() == "TRUE")
                    {
                        worksheet.Cell(5, 2).Formula = "=" + CouponAboveIRR.ToString() + "%";
                        worksheet.Cell(5, 4).Formula = "=(POWER((1 + B5), (F3 / 365)) - 1)*100%";
                    }
                    else
                    {
                        worksheet.Cell(5, 2).Formula = "=((POWER((1+D5),(12/D3))-1) * 100) %";
                        worksheet.Cell(5, 4).Formula = "=" + CouponAboveStrike.ToString() + "%";
                    }

                    if (IsCouponBelowIRR.ToUpper() == "TRUE")
                    {
                        worksheet.Cell(6, 2).Formula = "=" + CouponBelowIRR.ToString() + "%";
                        worksheet.Cell(6, 4).Formula = "=(POWER((1 + B6), (F3 / 365)) - 1)*100%";
                    }
                    else
                    {
                        worksheet.Cell(6, 2).Formula = "=((POWER((1+D6),(12/D3))-1) * 100) %";
                        worksheet.Cell(6, 4).Formula = "=" + CouponBelowStrike.ToString() + "%";
                    }

                    worksheet.Cell(7, 2).Formula = "=(100-(B2+D2)*100)-(100*(1+ROUND(D5,4)))/(POWER((1+(IF(H4>0,H4,F4))),(D3/12)))";
                    worksheet.Cell(7, 4).Formula = "=" + Strike.ToString() + "%";

                    worksheet.Cell(9, 2).Value = InitialAveragingMonth.ToString();
                    worksheet.Cell(9, 4).Value = InitialAveragingDaysDiff.ToString();
                    worksheet.Cell(9, 6).Value = FinalAveragingMonth.ToString();
                    worksheet.Cell(9, 8).Value = FinalAveragingDaysDiff.ToString();

                    if (SalesComments != null)
                        worksheet.Cell(11, 2).Value = SalesComments.ToString();
                    else
                        worksheet.Cell(11, 2).Value = "";

                    if (TradingComments != null)
                        worksheet.Cell(12, 2).Value = TradingComments.ToString();
                    else
                        worksheet.Cell(12, 2).Value = "";

                    if (CouponScenario != null)
                        worksheet.Cell(13, 2).Value = CouponScenario.ToString();
                    else
                        worksheet.Cell(13, 2).Value = "";

                    xlPackage.Save();
                }

                return Json("");
            }
            catch (Exception ex)
            {

                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];

                ClearFixedCouponMLDSession();

                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "ExportFixedCouponMLD", objUserMaster.UserID);
                return Json("");
            }

        }

        public List<Graph> GenerateFixedMLDGraphCalculation(double Strike, double BelowStrikeCoupon, double AfterStrikeCoupon, Int32 RedemptionDays)
        {
            var transactionCounts = new List<Graph>();
            try
            {
                DataTable dtGraph = new DataTable();

                dtGraph.Columns.Add("INITIAL");
                dtGraph.Columns.Add("FINAL");
                dtGraph.Columns.Add("NIFTY_PERFORMANCE");
                dtGraph.Columns.Add("COUPON");
                dtGraph.Columns.Add("ANNUALIZED_RETURN");

                int Count = 21;

                DataRow dr;
                var Initial = 100;
                var Nifty = -110;
                var Final = 0;
                var PrevFinal = 0;
                bool strike = false;

                double[] arrStrike = { Strike };

                //var transactionCounts = new List<Graph>();
                for (int i = 0; i <= Count; i++)
                {
                    dr = dtGraph.NewRow();

                    Nifty = Nifty + 10;
                    Final = (int)(Initial + ((100 * Nifty) / 100));

                    if (Final == 210)
                    {
                        return transactionCounts;
                    }

                    //if (Final == Strike)                    
                    if (arrStrike.Contains(Final))
                    {
                        GenerateRow(dtGraph, Final, Nifty, BelowStrikeCoupon, AfterStrikeCoupon, RedemptionDays, transactionCounts);
                        strike = true;
                    }
                    else if (arrStrike[0] > PrevFinal && arrStrike[0] < Final)
                    {
                        GenerateRow(dtGraph, arrStrike[0], Initial - arrStrike[0], BelowStrikeCoupon, AfterStrikeCoupon, RedemptionDays, transactionCounts);
                        strike = true;
                        Nifty = Nifty - 10;
                    }
                    else
                    {
                        dr["INITIAL"] = Initial;
                        dr["FINAL"] = Final;
                        dr["NIFTY_PERFORMANCE"] = Nifty;

                        if (strike)
                        {
                            dr["COUPON"] = AfterStrikeCoupon;
                            dr["ANNUALIZED_RETURN"] = Math.Pow((1 + AfterStrikeCoupon), (365 * 1.000 / RedemptionDays * 1.000)) - 1;
                        }
                        else
                        {
                            dr["COUPON"] = BelowStrikeCoupon;
                            dr["ANNUALIZED_RETURN"] = Math.Pow((1 + BelowStrikeCoupon), (365 * 1.000 / RedemptionDays * 1.000)) - 1;

                        }

                        dtGraph.Rows.Add(dr);
                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["COUPON"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                    }

                    PrevFinal = Final;
                }

                return transactionCounts;
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];

                ClearFixedCouponMLDSession();

                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "GenerateFixedMLDGraphCalculation", objUserMaster.UserID);
                return transactionCounts;
            }
        }
        #endregion

        #region Fixed Plus PR
        public JsonResult ManageFixedPlusPR(string ProductID, string Distributor, string EdelweissBuiltIn, string DistributorBuiltIn, string BuiltInAdjustment, string TotalBuiltIn, string IRR, string IsIRR, string DeploymentRate, string CustomerDeploymentRate, string FixedCoupon, string OptionTenureMonth, string RedemptionPeriodMonth, string IsRedemptionPeriodMonth, string RedemptionPeriodDays, string InitialAveragingMonth, string InitialAveragingDaysDiff, string FinalAveragingMonth, string FinalAveragingDaysDiff, string Underlying, string Remaining, string TotalOptionPrice, string NetRemaining, string SalesComments, string TradingComments, string CouponScenario1, string CouponScenario2,
            string CallOptionType, string CallStrike1, string CallStrike2, string CallParticipatoryRatio, string CallPrice, string CallDiscountedPrice, string CallPRAdjustedPrice, string CallIV1, string CallCustomIV1, string CallRF1, string CallCustomRF1, string CallIV2, string CallCustomIV2, string CallRF2, string CallCustomRF2,
            string PutOptionType, string PutStrike1, string PutStrike2, string PutParticipatoryRatio, string PutPrice, string PutDiscountedPrice, string PutPRAdjustedPrice, string PutIV1, string PutCustomIV1, string PutRF1, string PutCustomRF1, string PutIV2, string PutCustomIV2, string PutRF2, string PutCustomRF2, string IsPrincipalProtected, string CopyProductID,
            string CallStrike1Summary, string CallStrike2Summary, string PutStrike1Summary, string PutStrike2Summary,
            string ExportCallStrike1Summary, string ExportCallStrike2Summary, string ExportPutStrike1Summary, string ExportPutStrike2Summary, string Entity, string IsSecured)
        {
            try
            {

                ExportCallStrike1Summary = System.Uri.UnescapeDataString(ExportCallStrike1Summary);
                ExportCallStrike2Summary = System.Uri.UnescapeDataString(ExportCallStrike2Summary);
                ExportPutStrike1Summary = System.Uri.UnescapeDataString(ExportPutStrike1Summary);
                ExportPutStrike2Summary = System.Uri.UnescapeDataString(ExportPutStrike2Summary);
                CallStrike1Summary = System.Uri.UnescapeDataString(CallStrike1Summary);
                CallStrike2Summary = System.Uri.UnescapeDataString(CallStrike2Summary);
                PutStrike1Summary = System.Uri.UnescapeDataString(PutStrike1Summary);
                PutStrike2Summary = System.Uri.UnescapeDataString(PutStrike2Summary);
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];

                if (CustomerDeploymentRate == "")
                    CustomerDeploymentRate = "0";

                if (TotalOptionPrice == "")
                    TotalOptionPrice = "0";

                if (NetRemaining == "")
                    NetRemaining = "0";

                if (CallStrike1 == "")
                    CallStrike1 = "0";

                if (CallStrike2 == "")
                    CallStrike2 = "0";

                if (CallParticipatoryRatio == "")
                    CallParticipatoryRatio = "0";

                if (CallPrice == "")
                    CallPrice = "0";

                if (CallDiscountedPrice == "")
                    CallDiscountedPrice = "0";

                if (CallPRAdjustedPrice == "")
                    CallPRAdjustedPrice = "0";

                if (CallIV1 == "")
                    CallIV1 = "0";

                if (CallCustomIV1 == "")
                    CallCustomIV1 = "0";

                if (CallRF1 == "")
                    CallRF1 = "0";

                if (CallCustomRF1 == "")
                    CallCustomRF1 = "0";

                if (CallIV2 == "")
                    CallIV2 = "0";

                if (CallCustomIV2 == "")
                    CallCustomIV2 = "0";

                if (CallRF2 == "")
                    CallRF2 = "0";

                if (CallCustomRF2 == "")
                    CallCustomRF2 = "0";

                if (PutOptionType == null)
                    PutOptionType = "";

                if (PutStrike1 == "")
                    PutStrike1 = "0";

                if (PutStrike2 == "")
                    PutStrike2 = "0";

                if (PutParticipatoryRatio == "" || PutParticipatoryRatio == "NaN")
                    PutParticipatoryRatio = "0";

                if (PutPrice == "" || PutPrice == "NaN")
                    PutPrice = "0";

                if (PutDiscountedPrice == "" || PutDiscountedPrice == "NaN")
                    PutDiscountedPrice = "0";

                if (PutPRAdjustedPrice == "" || PutPRAdjustedPrice == "NaN")
                    PutPRAdjustedPrice = "0";

                if (PutIV1 == "" || PutIV1 == "NaN")
                    PutIV1 = "0";

                if (PutCustomIV1 == "")
                    PutCustomIV1 = "0";

                if (PutRF1 == "" || PutRF1 == "NaN")
                    PutRF1 = "0";

                if (PutCustomRF1 == "")
                    PutCustomRF1 = "0";

                if (PutIV2 == "" || PutIV2 == "NaN")
                    PutIV2 = "0";

                if (PutCustomIV2 == "")
                    PutCustomIV2 = "0";

                if (PutRF2 == "" || PutRF2 == "NaN")
                    PutRF2 = "0";

                if (PutCustomRF2 == "")
                    PutCustomRF2 = "0";

                string ParentProductID = "";
                if (Session["ParentProductID"] != null)
                    ParentProductID = (string)Session["ParentProductID"];

                ObjectResult<ManageFixedPlusPRResult> objManageFixedPlusPRResult = objSP_PRICINGEntities.SP_MANAGE_FIXED_PLUS_PR_DETAILS(ProductID, ParentProductID, Distributor, Convert.ToDouble(EdelweissBuiltIn), Convert.ToDouble(DistributorBuiltIn), Convert.ToDouble(BuiltInAdjustment), Convert.ToDouble(TotalBuiltIn), Convert.ToDouble(IRR), Convert.ToBoolean(IsIRR), Convert.ToDouble(DeploymentRate), Convert.ToDouble(CustomerDeploymentRate), Convert.ToDouble(FixedCoupon), Convert.ToInt32(OptionTenureMonth), Convert.ToDouble(RedemptionPeriodMonth), Convert.ToBoolean(IsRedemptionPeriodMonth), Convert.ToInt32(RedemptionPeriodDays), Convert.ToInt32(InitialAveragingMonth), Convert.ToInt32(InitialAveragingDaysDiff), Convert.ToInt32(FinalAveragingMonth), Convert.ToInt32(FinalAveragingDaysDiff), Convert.ToInt32(Underlying), Convert.ToDouble(Remaining), Convert.ToDouble(TotalOptionPrice), Convert.ToDouble(NetRemaining), SalesComments, TradingComments, CouponScenario1, CouponScenario2, Convert.ToInt32(Entity), Convert.ToInt32(IsSecured), objUserMaster.UserID,
                    CallOptionType, Convert.ToDouble(CallStrike1), Convert.ToDouble(CallStrike2), Convert.ToDouble(CallParticipatoryRatio), Convert.ToDouble(CallPrice), Convert.ToDouble(CallDiscountedPrice), Convert.ToDouble(CallPRAdjustedPrice), Convert.ToDouble(CallIV1), Convert.ToDouble(CallCustomIV1), Convert.ToDouble(CallRF1), Convert.ToDouble(CallCustomRF1), Convert.ToDouble(CallIV2), Convert.ToDouble(CallCustomIV2), Convert.ToDouble(CallRF2), Convert.ToDouble(CallCustomRF2),
                    PutOptionType, Convert.ToDouble(PutStrike1), Convert.ToDouble(PutStrike2), Convert.ToDouble(PutParticipatoryRatio), Convert.ToDouble(PutPrice), Convert.ToDouble(PutDiscountedPrice), Convert.ToDouble(PutPRAdjustedPrice), Convert.ToDouble(PutIV1), Convert.ToDouble(PutCustomIV1), Convert.ToDouble(PutRF1), Convert.ToDouble(PutCustomRF1), Convert.ToDouble(PutIV2), Convert.ToDouble(PutCustomIV2), Convert.ToDouble(PutRF2), Convert.ToDouble(PutCustomRF2), Convert.ToBoolean(IsPrincipalProtected), CopyProductID, CallStrike1Summary, CallStrike2Summary, PutStrike1Summary, PutStrike2Summary,
                    ExportCallStrike1Summary, ExportCallStrike2Summary, ExportPutStrike1Summary, ExportPutStrike2Summary);
                List<ManageFixedPlusPRResult> ManageFixedPlusPRResultList = objManageFixedPlusPRResult.ToList();

                return Json(ManageFixedPlusPRResultList[0].ProductID);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];

                ClearFixedPlusPRSession();

                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "ManageFixedPlusPR", objUserMaster.UserID);
                return Json("");
            }
        }

        public virtual void ExportFixedPlusPR(FixedPlusPR objFixedPlusPR, FormCollection objFormCollection)
        {
            try
            {
                string strTemplateFilePath = Server.MapPath("~/Templates");
                string strTemplateFileName = strTemplateFilePath + "\\FixedPlusPRTemplate.xlsx";

                string strTargetFilePath = Server.MapPath("~/OutputFiles");
                string strTargetFileName = strTargetFilePath + "\\" + objFixedPlusPR.ProductID + "_FixedPlusPR.xlsx";

                string Role = Convert.ToString(Session["Role"]);

                if (System.IO.File.Exists(strTargetFileName))
                    System.IO.File.Delete(strTargetFileName);

                FileInfo objTemplateFileInfo = new FileInfo(strTemplateFileName);
                objTemplateFileInfo.CopyTo(strTargetFileName);

                FileInfo objTargetFileInfo = new FileInfo(strTargetFileName);

                Underlying objUnderlying = objFixedPlusPR.UnderlyingList.Find(delegate(Underlying oUnderlying) { return oUnderlying.UnderlyingID == objFixedPlusPR.UnderlyingID; });

                using (var xlPackage = new ExcelPackage(objTargetFileInfo))
                {
                    var worksheet = xlPackage.Workbook.Worksheets["FixedPlusPR"];

                    worksheet.Cell(1, 2).Value = objFixedPlusPR.ProductID.ToString();
                    worksheet.Cell(1, 4).Value = objFixedPlusPR.Distributor.ToString().ToUpper();
                    worksheet.Cell(1, 6).Value = objUnderlying.UnderlyingShortName;

                    if (objFixedPlusPR.IsPrincipalProtected)
                        worksheet.Cell(1, 8).Value = "Yes";
                    else
                        worksheet.Cell(1, 8).Value = "No";

                    worksheet.Cell(2, 2).Formula = "=" + objFixedPlusPR.EdelweissBuiltIn.ToString() + "%";
                    worksheet.Cell(2, 4).Formula = "=" + objFixedPlusPR.DistributorBuiltIn.ToString() + "%";
                    worksheet.Cell(2, 6).Formula = "=B2+D2+H2";
                    worksheet.Cell(2, 8).Formula = "=" + objFixedPlusPR.BuiltInAdjustment.ToString() + "%";

                    worksheet.Cell(3, 2).Formula = "=((POWER((1+D3),(12/D4))-1)*100) %";
                    worksheet.Cell(3, 4).Formula = "=" + objFixedPlusPR.FixedCouponValue.ToString() + "%";

                    worksheet.Cell(4, 2).Value = objFixedPlusPR.OptionTenureMonth.ToString();
                    worksheet.Cell(4, 4).Formula = "=ROUND(F4/30.417,0)";
                    worksheet.Cell(4, 6).Value = objFixedPlusPR.RedemptionPeriodDays.ToString();

                    #region Get Entity Name
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
                    LookupMaster objLookupMasterEntity = EntityList.Find(delegate(LookupMaster oLookupMaster) { return oLookupMaster.LookupID == objFixedPlusPR.EntityID; });
                    worksheet.Cell(5, 2).Value = objLookupMasterEntity.LookupDescription;
                    #endregion

                    #region Get Is Secured
                    objLookupResult = null;
                    LookupResultList = null;
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
                    LookupMaster objLookupMasterIsSecured = IsSecuredList.Find(delegate(LookupMaster oLookupMaster) { return oLookupMaster.LookupID == objFixedPlusPR.IsSecuredID; });
                    worksheet.Cell(5, 4).Value = objLookupMasterIsSecured.LookupDescription;
                    #endregion
                    worksheet.Cell(5, 6).Formula = "=" + objFixedPlusPR.DeploymentRate.ToString() + "%";
                    worksheet.Cell(5, 8).Formula = "=" + objFixedPlusPR.CustomDeploymentRate.ToString() + "%";

                    worksheet.Cell(6, 2).Formula = "=(100-(B2+D2)*100)-(100*(1+ROUND(D3,4)))/(POWER((1+(IF(H5>0,H5,F5))),(D4/12)))";
                    worksheet.Cell(6, 4).Formula = "=H11 + H12";
                    worksheet.Cell(6, 6).Formula = "=D6 + B6";

                    worksheet.Cell(8, 2).Value = objFixedPlusPR.InitialAveragingMonth.ToString();
                    worksheet.Cell(8, 4).Value = objFixedPlusPR.InitialAveragingDaysDiff.ToString();
                    worksheet.Cell(8, 6).Value = objFixedPlusPR.FinalAveragingMonth.ToString();
                    worksheet.Cell(8, 8).Value = objFixedPlusPR.FinalAveragingDaysDiff.ToString();

                    if (objFixedPlusPR.CallOptionTypeId != null)
                        worksheet.Cell(11, 1).Value = objFixedPlusPR.CallOptionTypeId.ToString();
                    else
                        worksheet.Cell(11, 1).Value = "";

                    worksheet.Cell(11, 2).Value = objUnderlying.UnderlyingShortName;
                    worksheet.Cell(11, 3).Value = objFixedPlusPR.CallStrike1.ToString();
                    worksheet.Cell(11, 4).Value = objFixedPlusPR.CallStrike2.ToString();
                    worksheet.Cell(11, 5).Value = objFixedPlusPR.CallPaticipatoryRatio.ToString();
                    worksheet.Cell(11, 6).Value = objFixedPlusPR.CallPrice.ToString();
                    worksheet.Cell(11, 7).Value = objFixedPlusPR.CallDiscountedPrice.ToString();
                    worksheet.Cell(11, 8).Value = objFixedPlusPR.CallPrAdjustmentPrice.ToString();

                    if (Role == "Sales")
                    {
                        worksheet.Cell(10, 9).Value = "";
                        worksheet.Cell(10, 10).Value = "";
                        worksheet.Cell(10, 11).Value = "";
                        worksheet.Cell(10, 12).Value = "";
                        worksheet.Cell(10, 13).Value = "";
                        worksheet.Cell(10, 14).Value = "";
                        worksheet.Cell(10, 15).Value = "";
                        worksheet.Cell(10, 16).Value = "";
                    }
                    else
                    {
                        worksheet.Cell(11, 9).Formula = "=" + objFixedPlusPR.CallIV1.ToString() + "%";
                        worksheet.Cell(11, 10).Formula = "=" + objFixedPlusPR.CallCustomIV1.ToString() + "%";
                        worksheet.Cell(11, 11).Formula = "=" + objFixedPlusPR.CallRF1.ToString() + "%";
                        worksheet.Cell(11, 12).Formula = "=" + objFixedPlusPR.CallCustomRF1.ToString() + "%";
                        worksheet.Cell(11, 13).Formula = "=" + objFixedPlusPR.CallIV2.ToString() + "%";
                        worksheet.Cell(11, 14).Formula = "=" + objFixedPlusPR.CallCustomIV2.ToString() + "%";
                        worksheet.Cell(11, 15).Formula = "=" + objFixedPlusPR.CallRF2.ToString() + "%";
                        worksheet.Cell(11, 16).Formula = "=" + objFixedPlusPR.CallCustomRF2.ToString() + "%";
                    }

                    if (!objFixedPlusPR.IsPrincipalProtected)
                        if (objFixedPlusPR.PutStrike1 != 0)
                        {
                            if (objFixedPlusPR.PutOptionTypeId != null)
                                worksheet.Cell(12, 1).Value = objFixedPlusPR.PutOptionTypeId.ToString();
                            else
                                worksheet.Cell(12, 1).Value = "";
                            worksheet.Cell(12, 2).Value = objUnderlying.UnderlyingShortName;
                            worksheet.Cell(12, 3).Value = objFixedPlusPR.PutStrike1.ToString();
                            worksheet.Cell(12, 4).Value = objFixedPlusPR.PutStrike2.ToString();
                            worksheet.Cell(12, 5).Value = objFixedPlusPR.PutPaticipatoryRatio.ToString();
                            worksheet.Cell(12, 6).Value = objFixedPlusPR.PutPrice.ToString();
                            worksheet.Cell(12, 7).Value = objFixedPlusPR.PutDiscountedPrice.ToString();
                            worksheet.Cell(12, 8).Value = objFixedPlusPR.PutPrAdjustmentPrice.ToString();
                            if (Role == "Sales")
                            {
                                worksheet.Cell(10, 9).Value = "";
                                worksheet.Cell(10, 10).Value = "";
                                worksheet.Cell(10, 11).Value = "";
                                worksheet.Cell(10, 12).Value = "";
                                worksheet.Cell(10, 13).Value = "";
                                worksheet.Cell(10, 14).Value = "";
                                worksheet.Cell(10, 15).Value = "";
                                worksheet.Cell(10, 16).Value = "";
                            }
                            else
                            {
                                worksheet.Cell(12, 9).Formula = "=" + objFixedPlusPR.PutIV1.ToString() + "%";
                                worksheet.Cell(12, 10).Formula = "=" + objFixedPlusPR.PutCustomIV1.ToString() + "%";
                                worksheet.Cell(12, 11).Formula = "=" + objFixedPlusPR.PutRF1.ToString() + "%";
                                worksheet.Cell(12, 12).Formula = "=" + objFixedPlusPR.PutCustomRF1.ToString() + "%";
                                worksheet.Cell(12, 13).Formula = "=" + objFixedPlusPR.PutIV2.ToString() + "%";
                                worksheet.Cell(12, 14).Formula = "=" + objFixedPlusPR.PutCustomIV2.ToString() + "%";
                                worksheet.Cell(12, 15).Formula = "=" + objFixedPlusPR.PutRF2.ToString() + "%";
                                worksheet.Cell(12, 16).Formula = "=" + objFixedPlusPR.PutCustomRF2.ToString() + "%";
                            }
                        }

                    if (objFixedPlusPR.SalesComments != null)
                        worksheet.Cell(14, 2).Value = objFixedPlusPR.SalesComments.ToString();
                    else
                        worksheet.Cell(14, 2).Value = "";

                    if (objFixedPlusPR.TradingComments != null)
                        worksheet.Cell(15, 2).Value = objFixedPlusPR.TradingComments.ToString();
                    else
                        worksheet.Cell(15, 2).Value = "";

                    if (objFixedPlusPR.CouponScenario1 != null)
                        worksheet.Cell(16, 2).Value = objFixedPlusPR.CouponScenario1.ToString();
                    else
                        worksheet.Cell(16, 2).Value = "";

                    if (objFixedPlusPR.CouponScenario2 != null)
                        worksheet.Cell(17, 2).Value = objFixedPlusPR.CouponScenario2.ToString();
                    else
                        worksheet.Cell(17, 2).Value = "";

                    xlPackage.Save();
                }

                if (System.IO.File.Exists(strTargetFileName))
                {
                    FileInfo TemplateFile = new FileInfo(strTargetFileName);

                    Response.Clear();
                    Response.ClearHeaders();
                    Response.ClearContent();
                    Response.AddHeader("content-disposition", "attachment; filename=" + Path.GetFileName(strTargetFileName));
                    Response.AddHeader("Content-Type", "application/Excel");
                    Response.ContentType = "application/vnd.xls";
                    Response.AddHeader("Content-Length", TemplateFile.Length.ToString());
                    Response.WriteFile(TemplateFile.FullName);
                    Response.End();
                }
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];

                ClearFixedPlusPRSession();

                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "ExportFixedPlusPR", objUserMaster.UserID);

            }
        }

        public JsonResult ExportFixedPlusPRWorkingFile(string ProductID, string Distributor, string EdelweissBuiltIn, string DistributorBuiltIn, string BuiltInAdjustment, string TotalBuiltIn, string IRR, string IsIRR, string DeploymentRate, string CustomerDeploymentRate, string FixedCoupon, string OptionTenureMonth, string RedemptionPeriodMonth, string IsRedemptionPeriodMonth, string RedemptionPeriodDays, string InitialAveragingMonth, string InitialAveragingDaysDiff, string FinalAveragingMonth, string FinalAveragingDaysDiff, string Underlying, string Remaining, string TotalOptionPrice, string NetRemaining, string SalesComments, string TradingComments, string CouponScenario1, string CouponScenario2,
            string CallOptionType, string CallStrike1, string CallStrike2, string CallParticipatoryRatio, string CallPrice, string CallDiscountedPrice, string CallPRAdjustedPrice, string CallIV1, string CallCustomIV1, string CallRF1, string CallCustomRF1, string CallIV2, string CallCustomIV2, string CallRF2, string CallCustomRF2,
            string PutOptionType, string PutStrike1, string PutStrike2, string PutParticipatoryRatio, string PutPrice, string PutDiscountedPrice, string PutPRAdjustedPrice, string PutIV1, string PutCustomIV1, string PutRF1, string PutCustomRF1, string PutIV2, string PutCustomIV2, string PutRF2, string PutCustomRF2, string IsPrincipalProtected,
            string CallStrike1Summary, string CallStrike2Summary, string PutStrike1Summary, string PutStrike2Summary,
            string ExportCallStrike1Summary, string ExportCallStrike2Summary, string ExportPutStrike1Summary, string ExportPutStrike2Summary, string Entity, string IsSecured)
        {
            try
            {
                string strTemplateFilePath = Server.MapPath("~/Templates");
                string strTemplateFileName = strTemplateFilePath + "\\FixedPlusPRTemplateWorkingFile.xlsx";

                string strTargetFilePath = System.Configuration.ConfigurationManager.AppSettings["WorkingFilePath"];
                string strTargetFileName = strTargetFilePath + "\\" + ProductID + "_FixedPlusPR.xlsx";

                string Role = Convert.ToString(Session["Role"]);

                if (System.IO.File.Exists(strTargetFileName))
                    System.IO.File.Delete(strTargetFileName);

                FileInfo objTemplateFileInfo = new FileInfo(strTemplateFileName);
                objTemplateFileInfo.CopyTo(strTargetFileName);

                FileInfo objTargetFileInfo = new FileInfo(strTargetFileName);

                using (var xlPackage = new ExcelPackage(objTargetFileInfo))
                {
                    var worksheet = xlPackage.Workbook.Worksheets["FixedPlusPR"];

                    worksheet.Cell(1, 2).Value = ProductID.ToString();
                    worksheet.Cell(1, 4).Value = Distributor.ToString().ToUpper();
                    worksheet.Cell(1, 6).Value = Underlying;

                    if (IsPrincipalProtected == "True")
                        worksheet.Cell(1, 8).Value = "Yes";
                    else
                        worksheet.Cell(1, 8).Value = "No";

                    worksheet.Cell(2, 2).Formula = "=" + EdelweissBuiltIn.ToString() + "%";
                    worksheet.Cell(2, 4).Formula = "=" + DistributorBuiltIn.ToString() + "%";
                    worksheet.Cell(2, 6).Formula = "=B2+D2+H2";
                    worksheet.Cell(2, 8).Formula = "=" + BuiltInAdjustment.ToString() + "%";

                    if (IsIRR.ToUpper() == "TRUE")
                    {
                        worksheet.Cell(3, 2).Formula = "=" + IRR.ToString() + "%";
                        worksheet.Cell(3, 4).Formula = "=(POWER((1 + B3), (F4 / 365)) - 1)*100%";
                    }
                    else
                    {
                        worksheet.Cell(3, 2).Formula = "=((POWER((1+D3),(12/D4))-1) * 100) %";
                        worksheet.Cell(3, 4).Formula = "=" + FixedCoupon.ToString() + "%";
                    }

                    worksheet.Cell(4, 2).Value = OptionTenureMonth.ToString();
                    if (IsRedemptionPeriodMonth.ToUpper() == "TRUE")
                    {
                        worksheet.Cell(4, 4).Formula = RedemptionPeriodMonth;
                        worksheet.Cell(4, 6).Formula = "=ROUND(D4*30.417, 0)";
                    }
                    else
                    {
                        worksheet.Cell(4, 4).Formula = "=ROUND(F4/30.417,2)";
                        worksheet.Cell(4, 6).Formula = RedemptionPeriodDays.ToString();
                    }

                    #region Get Entity Name
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
                    LookupMaster objLookupMasterEntity = EntityList.Find(delegate(LookupMaster oLookupMaster) { return oLookupMaster.LookupID == Convert.ToInt32(Entity); });
                    worksheet.Cell(5, 2).Value = objLookupMasterEntity.LookupDescription;
                    #endregion

                    #region Get Is Secured
                    objLookupResult = null;
                    LookupResultList = null;
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
                    LookupMaster objLookupMasterIsSecured = IsSecuredList.Find(delegate(LookupMaster oLookupMaster) { return oLookupMaster.LookupID == Convert.ToInt32(IsSecured); });
                    worksheet.Cell(5, 4).Value = objLookupMasterIsSecured.LookupDescription;
                    #endregion
                    worksheet.Cell(5, 6).Formula = "=" + DeploymentRate.ToString() + "%";

                    if (CustomerDeploymentRate == "")
                        CustomerDeploymentRate = "0";
                    worksheet.Cell(5, 8).Formula = "=" + CustomerDeploymentRate.ToString() + "%";

                    worksheet.Cell(6, 2).Formula = "=(100-(B2+D2)*100)-(100*(1+ROUND(D3,4)))/(POWER((1+(IF(H5>0,H5,F5))),(D4/12)))";
                    worksheet.Cell(6, 4).Formula = "=H11 + H12";
                    worksheet.Cell(6, 6).Formula = "=ROUND(D6,4) + ROUND(B6,4)";

                    worksheet.Cell(8, 2).Value = InitialAveragingMonth.ToString();
                    worksheet.Cell(8, 4).Value = InitialAveragingDaysDiff.ToString();
                    worksheet.Cell(8, 6).Value = FinalAveragingMonth.ToString();
                    worksheet.Cell(8, 8).Value = FinalAveragingDaysDiff.ToString();

                    if (CallOptionType != null)
                        worksheet.Cell(11, 1).Value = CallOptionType.ToString();
                    else
                        worksheet.Cell(11, 1).Value = "";

                    worksheet.Cell(11, 2).Value = Underlying;
                    worksheet.Cell(11, 3).Value = CallStrike1.ToString();
                    worksheet.Cell(11, 4).Value = CallStrike2.ToString();

                    if (PutPRAdjustedPrice == "")
                        PutPRAdjustedPrice = "0";

                    if (PutStrike1 != "" && PutStrike1 != "")
                        worksheet.Cell(11, 5).Formula = "=(B6+H12)/F11";
                    else
                        worksheet.Cell(11, 5).Formula = "=(B6+0)/F11";

                    if (CallStrike2 != "0" && CallStrike2 != "")
                        worksheet.Cell(11, 6).Formula = "=(AVERAGE(INDIRECT(\"$S$2:\"&ADDRESS(1+$B$8,18+$F$8))))-(AVERAGE(INDIRECT(\"$AA$2:\"&ADDRESS(1+$B$8,26+$F$8))))";
                    else
                        worksheet.Cell(11, 6).Formula = "=AVERAGE(INDIRECT(\"$S$2:\"&ADDRESS(1+$B$8,18+$F$8)))";

                    worksheet.Cell(11, 7).Formula = "=F11*-1";
                    worksheet.Cell(11, 8).Formula = "=ROUND(G11,4)*ROUND(E11,4)";

                    worksheet.Cell(11, 9).Formula = "=" + CallIV1.ToString() + "%";

                    if (CallCustomIV1 == "")
                        CallCustomIV1 = "0";
                    worksheet.Cell(11, 10).Formula = "=" + CallCustomIV1.ToString() + "%";

                    worksheet.Cell(11, 11).Formula = "=" + CallRF1.ToString() + "%";

                    if (CallCustomRF1 == "")
                        CallCustomRF1 = "0";
                    worksheet.Cell(11, 12).Formula = "=" + CallCustomRF1.ToString() + "%";

                    worksheet.Cell(11, 13).Formula = "=" + CallIV2.ToString() + "%";

                    if (CallCustomIV2 == "")
                        CallCustomIV2 = "0";
                    worksheet.Cell(11, 14).Formula = "=" + CallCustomIV2.ToString() + "%";

                    worksheet.Cell(11, 15).Formula = "=" + CallRF2.ToString() + "%";

                    if (CallCustomRF2 == "")
                        CallCustomRF2 = "0";
                    worksheet.Cell(11, 16).Formula = "=" + CallCustomRF2.ToString() + "%";

                    if (IsPrincipalProtected.ToUpper() == "FALSE")
                    {
                        if (PutStrike1 != "0")
                        {
                            if (PutOptionType != null)
                                worksheet.Cell(12, 1).Value = PutOptionType.ToString();
                            else
                                worksheet.Cell(12, 1).Value = "";

                            worksheet.Cell(12, 2).Value = Underlying;
                            worksheet.Cell(12, 3).Value = PutStrike1.ToString();
                            worksheet.Cell(12, 4).Value = PutStrike2.ToString();
                            worksheet.Cell(12, 5).Value = PutParticipatoryRatio.ToString();

                            if (PutStrike2 != "0" && PutStrike2 != "")
                                worksheet.Cell(12, 6).Formula = "=(AVERAGE(INDIRECT(\"$S$10:\"&ADDRESS(9+$B$8,18+$F$8))))-(AVERAGE(INDIRECT(\"$AA$10:\"&ADDRESS(9+$B$8,26+$F$8))))";
                            else
                                worksheet.Cell(12, 6).Formula = "=AVERAGE(INDIRECT(\"$S$10:\"&ADDRESS(9+$B$8,18+$F$8)))";

                            worksheet.Cell(12, 7).Formula = "=F12";
                            worksheet.Cell(12, 8).Formula = "=ROUND(G12,4)*ROUND(E12,4)";

                            worksheet.Cell(12, 9).Formula = "=" + PutIV1.ToString() + "%";

                            if (PutCustomIV1 == "")
                                PutCustomIV1 = "0";
                            worksheet.Cell(12, 10).Formula = "=" + PutCustomIV1.ToString() + "%";

                            worksheet.Cell(12, 11).Formula = "=" + PutRF1.ToString() + "%";

                            if (PutCustomRF1 == "")
                                PutCustomRF1 = "0";
                            worksheet.Cell(12, 12).Formula = "=" + PutCustomRF1.ToString() + "%";

                            worksheet.Cell(12, 13).Formula = "=" + PutIV2.ToString() + "%";

                            if (PutCustomIV2 == "")
                                PutCustomIV2 = "0";
                            worksheet.Cell(12, 14).Formula = "=" + PutCustomIV2.ToString() + "%";

                            worksheet.Cell(12, 15).Formula = "=" + PutRF2.ToString() + "%";

                            if (PutCustomRF2 == "")
                                PutCustomRF2 = "0";
                            worksheet.Cell(12, 16).Formula = "=" + PutCustomRF2.ToString() + "%";
                        }
                    }

                    if (SalesComments != null)
                        worksheet.Cell(14, 2).Value = SalesComments.ToString();
                    else
                        worksheet.Cell(14, 2).Value = "";

                    if (TradingComments != null)
                        worksheet.Cell(15, 2).Value = TradingComments.ToString();
                    else
                        worksheet.Cell(15, 2).Value = "";

                    if (CouponScenario1 != null)
                        worksheet.Cell(16, 2).Value = CouponScenario1.ToString();
                    else
                        worksheet.Cell(16, 2).Value = "";

                    if (CouponScenario2 != null)
                        worksheet.Cell(17, 2).Value = CouponScenario2.ToString();
                    else
                        worksheet.Cell(17, 2).Value = "";

                    //---------------Write Put Spread Strike 1 IV Grid-----------------START------------
                    worksheet.Cell(1, 18).Formula = "=C11";
                    worksheet.Cell(1, 19).Formula = "=ROUND($B$4*30.417,0)";
                    worksheet.Cell(1, 20).Formula = "=ROUND($B$4*30.417,0) - (1*$H$8)";
                    worksheet.Cell(1, 21).Formula = "=ROUND($B$4*30.417,0) - (2*$H$8)";
                    worksheet.Cell(1, 22).Formula = "=ROUND($B$4*30.417,0) - (3*$H$8)";
                    worksheet.Cell(1, 23).Formula = "=ROUND($B$4*30.417,0) - (4*$H$8)";
                    worksheet.Cell(1, 24).Formula = "=ROUND($B$4*30.417,0) - (5*$H$8)";

                    worksheet.Cell(2, 18).Formula = "0";
                    worksheet.Cell(3, 18).Formula = "=$R$2+1*$D$8";
                    worksheet.Cell(4, 18).Formula = "=$R$2+2*$D$8";
                    worksheet.Cell(5, 18).Formula = "=$R$2+3*$D$8";
                    worksheet.Cell(6, 18).Formula = "=$R$2+4*$D$8";
                    worksheet.Cell(7, 18).Formula = "=$R$2+5*$D$8";

                    worksheet.Cell(2, 19).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,S1,IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(3, 19).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($S$1-R3),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(4, 19).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($S$1-R4),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(5, 19).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($S$1-R5),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(6, 19).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($S$1-R6),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(7, 19).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($S$1-R7),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";

                    worksheet.Cell(2, 20).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,T1,IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(3, 20).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($T$1-R3),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(4, 20).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($T$1-R4),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(5, 20).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($T$1-R5),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(6, 20).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($T$1-R6),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(7, 20).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($T$1-R7),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";

                    worksheet.Cell(2, 21).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,U1,IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(3, 21).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($U$1-R3),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(4, 21).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($U$1-R4),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(5, 21).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($U$1-R5),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(6, 21).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($U$1-R6),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(7, 21).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($U$1-R7),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";

                    worksheet.Cell(2, 22).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,V1,IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(3, 22).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($V$1-R3),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(4, 22).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($V$1-R4),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(5, 22).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($V$1-R5),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(6, 22).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($V$1-R6),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(7, 22).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($V$1-R7),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";

                    worksheet.Cell(2, 23).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,W1,IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(3, 23).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($W$1-R3),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(4, 23).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($W$1-R4),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(5, 23).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($W$1-R5),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(6, 23).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($W$1-R6),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(7, 23).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($W$1-R7),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";

                    worksheet.Cell(2, 24).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,X1,IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(3, 24).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($X$1-R3),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(4, 24).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($X$1-R4),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(5, 24).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($X$1-R5),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(6, 24).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($X$1-R6),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(7, 24).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($X$1-R7),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    //---------------Write Put Spread Strike 1 IV Grid-----------------END--------------

                    //---------------Write Put Spread Strike 2 IV Grid-----------------START------------
                    worksheet.Cell(1, 26).Formula = "=D11";
                    worksheet.Cell(1, 27).Formula = "=ROUND($B$4*30.417,0)";
                    worksheet.Cell(1, 28).Formula = "=ROUND($B$4*30.417,0) - (1*$H$8)";
                    worksheet.Cell(1, 29).Formula = "=ROUND($B$4*30.417,0) - (2*$H$8)";
                    worksheet.Cell(1, 30).Formula = "=ROUND($B$4*30.417,0) - (3*$H$8)";
                    worksheet.Cell(1, 31).Formula = "=ROUND($B$4*30.417,0) - (4*$H$8)";
                    worksheet.Cell(1, 32).Formula = "=ROUND($B$4*30.417,0) - (5*$H$8)";

                    worksheet.Cell(2, 26).Formula = "0";
                    worksheet.Cell(3, 26).Formula = "=$Z$2+1*$D$8";
                    worksheet.Cell(4, 26).Formula = "=$Z$2+2*$D$8";
                    worksheet.Cell(5, 26).Formula = "=$Z$2+3*$D$8";
                    worksheet.Cell(6, 26).Formula = "=$Z$2+4*$D$8";
                    worksheet.Cell(7, 26).Formula = "=$Z$2+5*$D$8";

                    worksheet.Cell(2, 27).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,AA1,IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(3, 27).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AA$1-Z3),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(4, 27).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AA$1-Z4),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(5, 27).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AA$1-Z5),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(6, 27).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AA$1-Z6),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(7, 27).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AA$1-Z7),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";

                    worksheet.Cell(2, 28).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,AB1,IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(3, 28).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AB$1-Z3),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(4, 28).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AB$1-Z4),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(5, 28).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AB$1-Z5),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(6, 28).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AB$1-Z6),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(7, 28).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AB$1-Z7),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";

                    worksheet.Cell(2, 29).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,AC1,IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(3, 29).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AC$1-Z3),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(4, 29).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AC$1-Z4),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(5, 29).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AC$1-Z5),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(6, 29).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AC$1-Z6),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(7, 29).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AC$1-Z7),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";

                    worksheet.Cell(2, 30).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,AD1,IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(3, 30).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AD$1-Z3),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(4, 30).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AD$1-Z4),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(5, 30).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AD$1-Z5),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(6, 30).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AD$1-Z6),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(7, 30).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AD$1-Z7),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";

                    worksheet.Cell(2, 31).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,AE1,IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(3, 31).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AE$1-Z3),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(4, 31).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AE$1-Z4),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(5, 31).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AE$1-Z5),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(6, 31).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AE$1-Z6),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(7, 31).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AE$1-Z7),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";

                    worksheet.Cell(2, 32).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,AF1,IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(3, 32).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AF$1-Z3),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(4, 32).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AF$1-Z4),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(5, 32).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AF$1-Z5),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(6, 32).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AF$1-Z6),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(7, 32).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AF$1-Z7),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    //---------------Write Put Spread Strike 2 IV Grid-----------------END--------------

                    //---------------Write Put Strike 1 IV Grid-----------------START------------
                    if (PutStrike1 != "")
                    {
                        worksheet.Cell(9, 18).Formula = "=C12";
                        worksheet.Cell(9, 19).Formula = "=ROUND($B$4*30.417,0)";
                        worksheet.Cell(9, 20).Formula = "=ROUND($B$4*30.417,0) - (1*$H$8)";
                        worksheet.Cell(9, 21).Formula = "=ROUND($B$4*30.417,0) - (2*$H$8)";
                        worksheet.Cell(9, 22).Formula = "=ROUND($B$4*30.417,0) - (3*$H$8)";
                        worksheet.Cell(9, 23).Formula = "=ROUND($B$4*30.417,0) - (4*$H$8)";
                        worksheet.Cell(9, 24).Formula = "=ROUND($B$4*30.417,0) - (5*$H$8)";

                        worksheet.Cell(10, 18).Formula = "0";
                        worksheet.Cell(11, 18).Formula = "=$R$10+1*$D$8";
                        worksheet.Cell(12, 18).Formula = "=$R$10+2*$D$8";
                        worksheet.Cell(13, 18).Formula = "=$R$10+3*$D$8";
                        worksheet.Cell(14, 18).Formula = "=$R$10+4*$D$8";
                        worksheet.Cell(15, 18).Formula = "=$R$10+5*$D$8";

                        worksheet.Cell(10, 19).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,S9,IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(11, 19).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($S$9-R11),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(12, 19).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($S$9-R12),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(13, 19).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($S$9-R13),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(14, 19).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($S$9-R14),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(15, 19).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($S$9-R15),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";

                        worksheet.Cell(10, 20).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,T9,IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(11, 20).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($T$9-R11),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(12, 20).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($T$9-R12),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(13, 20).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($T$9-R13),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(14, 20).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($T$9-R14),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(15, 20).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($T$9-R15),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";

                        worksheet.Cell(10, 21).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,U9,IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(11, 21).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($U$9-R11),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(12, 21).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($U$9-R12),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(13, 21).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($U$9-R13),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(14, 21).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($U$9-R14),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(15, 21).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($U$9-R15),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";

                        worksheet.Cell(10, 22).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,V9,IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(11, 22).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($V$9-R11),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(12, 22).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($V$9-R12),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(13, 22).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($V$9-R13),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(14, 22).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($V$9-R14),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(15, 22).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($V$9-R15),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";

                        worksheet.Cell(10, 23).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,W9,IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(11, 23).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($W$9-R11),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(12, 23).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($W$9-R12),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(13, 23).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($W$9-R13),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(14, 23).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($W$9-R14),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(15, 23).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($W$9-R15),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";

                        worksheet.Cell(10, 24).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,X9,IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(11, 24).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($X$9-R11),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(12, 24).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($X$9-R12),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(13, 24).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($X$9-R13),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(14, 24).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($X$9-R14),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(15, 24).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($X$9-R15),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                    }
                    //---------------Write Put Strike 1 IV Grid-----------------END--------------

                    //---------------Write Put Strike 2 IV Grid-----------------START------------
                    if (PutStrike2 != "")
                    {
                        worksheet.Cell(9, 26).Formula = "=D12";
                        worksheet.Cell(9, 27).Formula = "=ROUND($B$4*30.417,0)";
                        worksheet.Cell(9, 28).Formula = "=ROUND($B$4*30.417,0) - (1*$H$8)";
                        worksheet.Cell(9, 29).Formula = "=ROUND($B$4*30.417,0) - (2*$H$8)";
                        worksheet.Cell(9, 30).Formula = "=ROUND($B$4*30.417,0) - (3*$H$8)";
                        worksheet.Cell(9, 31).Formula = "=ROUND($B$4*30.417,0) - (4*$H$8)";
                        worksheet.Cell(9, 32).Formula = "=ROUND($B$4*30.417,0) - (5*$H$8)";

                        worksheet.Cell(10, 26).Formula = "0";
                        worksheet.Cell(11, 26).Formula = "=$Z$10+1*$D$8";
                        worksheet.Cell(12, 26).Formula = "=$Z$10+2*$D$8";
                        worksheet.Cell(13, 26).Formula = "=$Z$10+3*$D$8";
                        worksheet.Cell(14, 26).Formula = "=$Z$10+4*$D$8";
                        worksheet.Cell(15, 26).Formula = "=$Z$10+5*$D$8";

                        worksheet.Cell(10, 27).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,AA9,IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(11, 27).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AA$9-Z11),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(12, 27).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AA$9-Z12),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(13, 27).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AA$9-Z13),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(14, 27).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AA$9-Z14),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(15, 27).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AA$9-Z15),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";

                        worksheet.Cell(10, 28).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,AB9,IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(11, 28).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AB$9-Z11),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(12, 28).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AB$9-Z12),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(13, 28).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AB$9-Z13),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(14, 28).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AB$9-Z14),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(15, 28).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AB$9-Z15),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";

                        worksheet.Cell(10, 29).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,AC9,IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(11, 29).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AC$9-Z11),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(12, 29).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AC$9-Z12),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(13, 29).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AC$9-Z13),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(14, 29).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AC$9-Z14),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(15, 29).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AC$9-Z15),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";

                        worksheet.Cell(10, 30).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,AD9,IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(11, 30).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AD$9-Z11),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(12, 30).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AD$9-Z12),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(13, 30).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AD$9-Z13),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(14, 30).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AD$9-Z14),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(15, 30).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AD$9-Z15),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";

                        worksheet.Cell(10, 31).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,AE9,IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(11, 31).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AE$9-Z11),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(12, 31).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AE$9-Z12),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(13, 31).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AE$9-Z13),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(14, 31).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AE$9-Z14),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(15, 31).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AE$9-Z15),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";

                        worksheet.Cell(10, 32).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,AF9,IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(11, 32).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AF$9-Z11),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(12, 32).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AF$9-Z12),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(13, 32).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AF$9-Z13),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(14, 32).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AF$9-Z14),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(15, 32).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AF$9-Z15),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                    }
                    //---------------Write Put Strike 2 IV Grid-----------------END--------------

                    xlPackage.Save();

                    return Json("");
                }
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];

                ClearFixedPlusPRSession();

                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "ExportFixedPlusPR", objUserMaster.UserID);
                return Json("");
            }
        }

        [HttpGet]
        public ActionResult FixedPlusPR(string ProductID, string GenerateGraph, bool IsQuotron = false)
        {
            LoginController objLoginController = new LoginController();
            try
            {
                if (ValidateSession())
                {
                    FixedPlusPR objFixedPlusPR = new FixedPlusPR();

                    #region Menu Access By on Role

                    Int32 intResult = 0;
                    // bool PPorNonPP = false;

                    UserMaster objUserMaster = new UserMaster();
                    objUserMaster = (UserMaster)Session["LoggedInUser"];

                    var Result = objSP_PRICINGEntities.VALIDATE_MENU_BY_ROLE(objUserMaster.RoleID, "BFPP");
                    intResult = Convert.ToInt32(Result.SingleOrDefault());

                    if (intResult == 0)
                        return RedirectToAction("UserNotAuthorize", "Login");

                    #endregion

                    #region Bind Underlying List
                    List<Underlying> UnderlyingList = new List<Underlying>();

                    ObjectResult<UnderlyingListResult> objUnderlyingListResult;
                    List<UnderlyingListResult> UnderlyingListResultList = new List<UnderlyingListResult>();

                    objUnderlyingListResult = objSP_PRICINGEntities.SP_FETCH_UNDERLYING_DETAILS();
                    UnderlyingListResultList = objUnderlyingListResult.ToList();

                    foreach (UnderlyingListResult oUnderlyingListResult in UnderlyingListResultList)
                    {
                        Underlying objUnderlying = new Underlying();
                        General.ReflectSingleData(objUnderlying, oUnderlyingListResult);

                        UnderlyingList.Add(objUnderlying);
                    }
                    objFixedPlusPR.UnderlyingList = UnderlyingList;

                    //--Set default underlying--Added by Shweta on 3rd May 2016------------START--------------------
                    string strDefaultUnderlying = System.Configuration.ConfigurationManager.AppSettings["DefaultUnderlying"].ToUpper();
                    Underlying objDefaulyUnderlying = objFixedPlusPR.UnderlyingList.Find(delegate(Underlying oUnderlying) { return oUnderlying.UnderlyingShortName == strDefaultUnderlying; });
                    objFixedPlusPR.UnderlyingID = objDefaulyUnderlying.UnderlyingID;

                    objFixedPlusPR.EntityID = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["DefaultEntityID"]);
                    objFixedPlusPR.IsSecuredID = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["DefaultIsSecuredID"]);
                    //--Set default underlying--Added by Shweta on 3rd May 2016------------END----------------------
                    #endregion

                    if (ProductID != "" && ProductID != null)
                    {
                        ObjectResult<FixedPlusPREditResult> objFixedPlusPREditResult = objSP_PRICINGEntities.FETCH_FIXED_PLUS_PR_EDIT_DETAILS(ProductID);
                        List<FixedPlusPREditResult> FixedPlusPREditResultList = objFixedPlusPREditResult.ToList();

                        General.ReflectSingleData(objFixedPlusPR, FixedPlusPREditResultList[0]);

                        DataSet dsResult = new DataSet();
                        dsResult = General.ExecuteDataSet("FETCH_UNDERLYING_CREATION_BYID", objFixedPlusPR.UnderlyingID);

                        if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                        {
                            ViewBag.UnderlyingShortName = Convert.ToString(dsResult.Tables[0].Rows[0]["UNDERLYING_SHORTNAME"]);
                        }
                    }
                    else
                    {
                        objFixedPlusPR.IsIRR = true;
                        objFixedPlusPR.IsRedemptionPeriodMonth = true;
                    }

                    if (GenerateGraph == "GenerateGraph")
                    {
                        objFixedPlusPR = (FixedPlusPR)TempData["FixedPlusGraph"];
                        ObjectResult<FixedPlusPREditResult> objFixedPlusPREditResult = objSP_PRICINGEntities.FETCH_FIXED_PLUS_PR_EDIT_DETAILS(objFixedPlusPR.ProductID);
                        List<FixedPlusPREditResult> FixedPlusPREditResultList = objFixedPlusPREditResult.ToList();
                        FixedPlusPR oFixedPlusPR = new FixedPlusPR();

                        General.ReflectSingleData(oFixedPlusPR, FixedPlusPREditResultList[0]);

                        objFixedPlusPR.Status = oFixedPlusPR.Status;
                        //objFixedPlusPR.SaveStatus = oFixedPlusPR.SaveStatus;
                        return GenerateFixedPlusPRGraph(objFixedPlusPR);
                    }

                    else if (Session["FixedPlusPRCopyQuote"] != null)
                    {
                        ViewBag.Message = "Quote is copied";
                        objFixedPlusPR = (FixedPlusPR)Session["FixedPlusPRCopyQuote"];

                        objFixedPlusPR.UnderlyingID = Convert.ToInt32(Session["UnderlyingID"]);

                        ObjectResult<FixedPlusPREditResult> objFixedPlusPREditResult = objSP_PRICINGEntities.FETCH_FIXED_PLUS_PR_EDIT_DETAILS(objFixedPlusPR.ProductID);
                        List<FixedPlusPREditResult> FixedPlusPREditResultList = objFixedPlusPREditResult.ToList();
                        FixedPlusPR oFixedPlusPR = new FixedPlusPR();
                        if (FixedPlusPREditResultList != null && FixedPlusPREditResultList.Count > 0)
                            General.ReflectSingleData(oFixedPlusPR, FixedPlusPREditResultList[0]);

                        objFixedPlusPR.ParentProductID = objFixedPlusPR.ProductID;
                        objFixedPlusPR.ProductID = "";
                        objFixedPlusPR.Status = "";
                        objFixedPlusPR.SaveStatus = "";
                        objFixedPlusPR.IsCopyQuote = true;
                        objFixedPlusPR.IsIRR = oFixedPlusPR.IsIRR;
                        objFixedPlusPR.IsRedemptionPeriodMonth = oFixedPlusPR.IsRedemptionPeriodMonth;

                        //Added by Shweta on 10th May---------------START-------
                        objFixedPlusPR.CallCustomIV1 = 0;
                        objFixedPlusPR.CallCustomIV2 = 0;
                        objFixedPlusPR.CallCustomRF1 = 0;
                        objFixedPlusPR.CallCustomRF2 = 0;

                        objFixedPlusPR.PutCustomIV1 = 0;
                        objFixedPlusPR.PutCustomIV2 = 0;
                        objFixedPlusPR.PutCustomRF1 = 0;
                        objFixedPlusPR.PutCustomRF2 = 0;
                        //Added by Shweta on 10th May---------------END---------

                        //-------------Added by Shweta on 22nd July 2016 to Fetch Latest Deployment Rate------------START--------
                        string strDeploymentRate = "";
                        var DeploymentRate = objSP_PRICINGEntities.SP_FETCH_PRICING_DEPLOYMENT_RATE(Convert.ToInt32(objFixedPlusPR.RedemptionPeriodDays), objFixedPlusPR.EntityID, objFixedPlusPR.IsSecuredID);
                        strDeploymentRate = Convert.ToString(DeploymentRate.SingleOrDefault());
                        objFixedPlusPR.DeploymentRate = Convert.ToDouble(strDeploymentRate);
                        //-------------Added by Shweta on 22nd July 2016 to Fetch Latest Deployment Rate------------END----------
                    }

                    else if (Session["FixedPlusPRChildQuote"] != null)
                    {
                        ViewBag.Message = true;
                        objFixedPlusPR = (FixedPlusPR)Session["FixedPlusPRChildQuote"];
                        objFixedPlusPR.UnderlyingID = Convert.ToInt32(Session["UnderlyingID"]);

                        ObjectResult<FixedPlusPREditResult> objFixedPlusPREditResult = objSP_PRICINGEntities.FETCH_FIXED_PLUS_PR_EDIT_DETAILS("");
                        List<FixedPlusPREditResult> FixedPlusPREditResultList = objFixedPlusPREditResult.ToList();
                        FixedPlusPR oFixedPlusPR = new FixedPlusPR();
                        if (FixedPlusPREditResultList != null && FixedPlusPREditResultList.Count > 0)
                            General.ReflectSingleData(oFixedPlusPR, FixedPlusPREditResultList[0]);

                        objFixedPlusPR.ParentProductID = objFixedPlusPR.ProductID;
                        objFixedPlusPR.ProductID = "";
                        objFixedPlusPR.Status = oFixedPlusPR.Status;
                        objFixedPlusPR.SaveStatus = oFixedPlusPR.SaveStatus;
                        objFixedPlusPR.IsChildQuote = true;
                    }
                    else if (Session["CancelQuote"] != null)
                    {
                        objFixedPlusPR = (FixedPlusPR)Session["CancelQuote"];

                        ObjectResult<FixedPlusPREditResult> objFixedPlusPREditResult = objSP_PRICINGEntities.FETCH_FIXED_PLUS_PR_EDIT_DETAILS(objFixedPlusPR.ProductID);
                        List<FixedPlusPREditResult> FixedPlusPREditResultList = objFixedPlusPREditResult.ToList();
                        FixedPlusPR oFixedPlusPR = new FixedPlusPR();
                        if (FixedPlusPREditResultList != null && FixedPlusPREditResultList.Count > 0)
                            General.ReflectSingleData(oFixedPlusPR, FixedPlusPREditResultList[0]);

                        objFixedPlusPR.Status = oFixedPlusPR.Status;
                        objFixedPlusPR.SaveStatus = oFixedPlusPR.SaveStatus;

                        Session.Remove("CancelQuote");
                    }
                    else
                    {
                        Session.Remove("IsChildQuoteFixedPlus");
                        Session.Remove("ParentProductID");
                        Session.Remove("UnderlyingID");
                    }

                    if (IsQuotron == true)
                    {
                        objFixedPlusPR.IsQuotron = true;
                    }

                    if (Session["FixedPlusPRChildQuote"] == null && Session["FixedPlusPRCopyQuote"] == null)
                        objFixedPlusPR.SaveStatus = "";

                    if (Session["FixedPlusPRCopyQuote"] != null)
                        Session.Remove("FixedPlusPRCopyQuote");

                    if (Session["FixedPlusPRChildQuote"] != null)
                        Session.Remove("FixedPlusPRChildQuote");


                    if (ProductID == null)
                    {
                        objFixedPlusPR.isGraphActive = false;
                        return View(objFixedPlusPR);
                    }
                    else
                    {
                        return GenerateFixedPlusPRGraph(objFixedPlusPR);
                    }
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

                ClearFixedPlusPRSession();

                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "FixedPlusPR Get", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        private ActionResult GenerateFixedPlusPRGraph(FixedPlusPR objFixedPlusPR)
        {
            LoginController objLoginController = new LoginController();
            try
            {
                if (ValidateSession())
                {

                    var ParticipatoryRatio = TruncateDecimal(objFixedPlusPR.CallPaticipatoryRatio, 2);

                    var transactionCounts = new List<Graph>();
                    transactionCounts = GenerateGraphFixedPRCalculation(objFixedPlusPR.CallStrike1, objFixedPlusPR.CallStrike2, objFixedPlusPR.PutStrike1, objFixedPlusPR.PutStrike2, ParticipatoryRatio, objFixedPlusPR.PutPaticipatoryRatio, objFixedPlusPR.CallOptionTypeId, objFixedPlusPR.PutOptionTypeId, objFixedPlusPR.FixedCouponValue, "FixedPlusPR");

                    // FixedPlusPR obj = new FixedPlusPR();

                    #region Pie Chart For FixedPlusPR
                    var xDataMonths = transactionCounts.Select(i => i.Column1).ToArray();
                    var yDataCounts = transactionCounts.Select(i => new object[] { i.Column2 }).ToArray();
                    var yDataCounts1 = transactionCounts.Select(i => new object[] { i.Column3 }).ToArray();

                    var FixedPRChart = new Highcharts("pie")
                        //define the type of chart 
                                .InitChart(new Chart { DefaultSeriesType = ChartTypes.Line })
                        //overall Title of the chart 
                                .SetTitle(new Title { Text = "Fixed Plus PR" })
                        ////small label below the main Title
                        //        .SetSubtitle(new Subtitle { Text = "Accounting" })
                        //load the X values
                                .SetXAxis(new XAxis { Title = new XAxisTitle { Text = "Underlying Returns" }, Categories = xDataMonths, Labels = new XAxisLabels { Step = 2 } })
                        //set the Y title
                                .SetYAxis(new YAxis { Title = new YAxisTitle { Text = "Product Returns" } })
                                .SetTooltip(new Tooltip
                                {
                                    Enabled = true,
                                    Formatter = @"function() { return '<b>'+ this.series.name +'</b><br/>'+ this.x +': '+ this.y; }"
                                })
                                .SetPlotOptions(new PlotOptions
                                {
                                    Line = new PlotOptionsLine
                                    {
                                        DataLabels = new PlotOptionsLineDataLabels
                                        {
                                            Enabled = false
                                        },
                                        EnableMouseTracking = true
                                    }
                                })
                        //load the Y values 
                                .SetSeries(new[]
                    {
                        new Series {Name = "Coupon", Data = new Data(yDataCounts)},
                            //you can add more y data to create a second line
                             //new Series { Name = "Strike", Data = new Data(yDataCounts1) }
                    });
                    #endregion

                    if (Session["FixedPlusPRCopyQuote"] != null)
                    {
                        ViewBag.Message = "Quote is copied";
                        objFixedPlusPR = (FixedPlusPR)Session["FixedPlusPRCopyQuote"];
                        Session.Remove("FixedPlusPRCopyQuote");
                    }

                    objFixedPlusPR.FixedPlusPRChart = FixedPRChart;

                    return View(objFixedPlusPR);
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

                ClearFixedPlusPRSession();

                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "GenerateFixedPlusPRGraph", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        [HttpPost, ValidateInput(false)]
        public ActionResult FixedPlusPR(string Command, FixedPlusPR objFixedPlusPR, FormCollection objFormCollection)
        {
            LoginController objLoginController = new LoginController();
            try
            {

                if (ValidateSession())
                {
                    #region Bind Underlying List
                    List<Underlying> UnderlyingList = new List<Underlying>();

                    ObjectResult<UnderlyingListResult> objUnderlyingListResult;
                    List<UnderlyingListResult> UnderlyingListResultList = new List<UnderlyingListResult>();

                    objUnderlyingListResult = objSP_PRICINGEntities.SP_FETCH_UNDERLYING_DETAILS();
                    UnderlyingListResultList = objUnderlyingListResult.ToList();

                    foreach (UnderlyingListResult oUnderlyingListResult in UnderlyingListResultList)
                    {
                        Underlying objUnderlying = new Underlying();
                        General.ReflectSingleData(objUnderlying, oUnderlyingListResult);

                        UnderlyingList.Add(objUnderlying);
                    }
                    objFixedPlusPR.UnderlyingList = UnderlyingList;
                    #endregion

                    FixedPlusPR oFixedPlusPR = new FixedPlusPR();
                    if (objFixedPlusPR.ProductID != "" && objFixedPlusPR.ProductID != null)
                    {
                        ObjectResult<FixedPlusPREditResult> objFixedPlusPREditResult = objSP_PRICINGEntities.FETCH_FIXED_PLUS_PR_EDIT_DETAILS(objFixedPlusPR.ProductID);
                        List<FixedPlusPREditResult> FixedPlusPREditResultList = objFixedPlusPREditResult.ToList();

                        General.ReflectSingleData(oFixedPlusPR, FixedPlusPREditResultList[0]);
                        objFixedPlusPR.IsPrincipalProtected = oFixedPlusPR.IsPrincipalProtected;

                        objFixedPlusPR.CallOptionTypeId = oFixedPlusPR.CallOptionTypeId;
                        objFixedPlusPR.PutOptionTypeId = oFixedPlusPR.PutOptionTypeId;
                    }
                    if (Command == "ExportToExcel")
                    {
                        ExportFixedPlusPR(objFixedPlusPR, objFormCollection);

                        return RedirectToAction("FixedPlusPR");
                    }
                    else if (Command == "ExportCallStrike1Grid")
                    {
                        string StrikeHTML = objFormCollection["ExportCallStrike1Summary"];

                        ExportStrikeGrid(StrikeHTML);

                        return RedirectToAction("FixedPlusPR");
                    }
                    else if (Command == "ExportCallStrike2Grid")
                    {
                        string StrikeHTML = objFormCollection["ExportCallStrike2Summary"];

                        ExportStrikeGrid(StrikeHTML);

                        return RedirectToAction("FixedPlusPR");
                    }
                    else if (Command == "ExportPutStrike1Grid")
                    {
                        string StrikeHTML = objFormCollection["ExportPutStrike1Summary"];

                        ExportStrikeGrid(StrikeHTML);

                        return RedirectToAction("FixedPlusPR");
                    }
                    else if (Command == "ExportPutStrike2Grid")
                    {
                        string StrikeHTML = objFormCollection["ExportPutStrike2Summary"];

                        ExportStrikeGrid(StrikeHTML);

                        return RedirectToAction("FixedPlusPR");
                    }
                    else if (Command == "GenerateGraph")
                    {
                        objFixedPlusPR.isGraphActive = true;

                        TempData["FixedPlusGraph"] = objFixedPlusPR;
                        return RedirectToAction("FixedPlusPR", new { GenerateGraph = "GenerateGraph" });

                        //return GenerateFixedPlusPRGraph(objFixedPlusPR);
                    }
                    else if (Command == "CopyQuote")
                    {
                        Session["FixedPlusPRCopyQuote"] = objFixedPlusPR;
                        Session["UnderlyingID"] = objFixedPlusPR.UnderlyingID;

                        return RedirectToAction("FixedPlusPR");
                    }
                    else if (Command == "CreateChildQuote")
                    {
                        Session.Remove("ParentProductID");
                        Session.Remove("IsChildQuoteFixedPlus");
                        Session.Remove("UnderlyingID");

                        Session["ParentProductID"] = objFixedPlusPR.ProductID;
                        Session["UnderlyingID"] = objFixedPlusPR.UnderlyingID;

                        objFixedPlusPR.IsChildQuote = true;

                        Session["FixedPlusPRChildQuote"] = objFixedPlusPR;
                        Session["IsChildQuoteFixedPlus"] = objFixedPlusPR.IsChildQuote;

                        return RedirectToAction("FixedPlusPR");
                    }
                    else if (Command == "AddNewProduct")
                    {
                        var productID = objFixedPlusPR.ProductID;
                        UserMaster objUserMaster = new UserMaster();
                        objUserMaster = (UserMaster)Session["LoggedInUser"];

                        EncryptDecrypt obj = new EncryptDecrypt();
                        var encryptedpaswd = obj.Encrypt(objUserMaster.Password, "SPPricing", CryptographyEngine.AlgorithmType.DES);

                        var isPrincipalProtected = objFormCollection.Get("PrincipalProtected");
                        var ProductType = "PP";
                        if (isPrincipalProtected != "1")
                        {
                            ProductType = "NonPP";
                        }

                        var Url = "http://edemumnewuatvm4:63400/Login.aspx?UserId=" + objUserMaster.LoginName + "&Key=" + encryptedpaswd + "&ProductId=" + productID + "&ProductType=" + ProductType;
                        return Redirect(Url);

                    }
                    else if (Command == "Cancel")
                    {
                        Session["CancelQuote"] = objFixedPlusPR;

                        return RedirectToAction("FixedPlusPR");
                    }
                    else if (Command == "PricingInExcel")
                    {
                        objFixedPlusPR.IsWorkingFileExport = OpenWorkingExcelFile("FPP", objFixedPlusPR.ProductID);

                        if (!objFixedPlusPR.IsWorkingFileExport)
                            objFixedPlusPR.WorkingFileStatus = "File Not Found";

                        return View(objFixedPlusPR);
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

                ClearFixedPlusPRSession();

                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "FixedPlusPR Post", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        public List<Graph> GenerateGraphFixedPRCalculation(double Strike11, double Strike12, double Strike21, double Strike22, double BelowStrikeCoupon, double AfterStrikeCoupon, string CallOptionType, string PutOptionType, double FixedCouponValue, string GraphType)
        {
            var transactionCounts = new List<Graph>();
            try
            {
                DataTable dtGraph = new DataTable();

                dtGraph.Columns.Add("INITIAL");
                dtGraph.Columns.Add("FINAL");
                dtGraph.Columns.Add("NIFTY_PERFORMANCE");
                dtGraph.Columns.Add("FIXED");

                BelowStrikeCoupon = Math.Round(BelowStrikeCoupon, 2);

                int Count = 25;

                DataRow dr;
                var Initial = 100;
                var Nifty = -110;
                var Final = 0.0;
                var PrevFinal = 0;
                var NegSpot = -100;
                bool strike = false;
                bool strike2 = false;
                bool strike3 = false;
                bool strike4 = false;
                bool strike5 = false;


                double[] arrStrike = { Strike11, Strike12, Strike21, Strike22 };

                //var transactionCounts = new List<Graph>();
                for (int i = 0; i <= Count; i++)
                {
                    dr = dtGraph.NewRow();

                    Nifty = Nifty + 10;
                    Final = (int)(Initial + ((100 * Nifty) / 100));
                    if (Final == 210)
                    {
                        return transactionCounts;
                    }


                    if (arrStrike[0] > PrevFinal && arrStrike[0] < Final)
                    {
                        Nifty = Nifty - 10;

                        #region Previous Row
                        //previous row
                        dr["INITIAL"] = Initial;
                        var final = arrStrike[0] - 0.01;
                        dr["FINAL"] = final;
                        var nifty = final - Initial;//Nifty <= 0 ? (Nifty - (arrStrike[0] % 10)) : (Nifty + (arrStrike[0] % 10));
                        //nifty = nifty - 0.01;
                        dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);
                        double dblPR = 0;
                        double dblFixed = 10;

                        if (GraphType == "FixedPlusPR")
                        {
                            dblPR = CalculatePR(nifty, Strike11, Strike12, Strike21, Strike22, BelowStrikeCoupon, AfterStrikeCoupon, CallOptionType, PutOptionType, FixedCouponValue, GraphType);
                            if (dblPR < NegSpot)
                            {
                                dr["FIXED"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                            }
                            else
                            {
                                dr["FIXED"] = Math.Round(dblPR, 2);
                            }
                        }
                        else if (GraphType == "FixedOrPR")
                        {
                            dblPR = CalculatePR(nifty, Strike11, Strike12, Strike21, Strike22, BelowStrikeCoupon, AfterStrikeCoupon, CallOptionType, PutOptionType, FixedCouponValue, GraphType);
                            if (dblPR < NegSpot)
                            {
                                dr["FIXED"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                            }
                            else
                            {
                                dr["FIXED"] = Math.Round(dblPR, 2);
                            }
                        }

                        dtGraph.Rows.Add(dr);
                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["FIXED"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                        #endregion

                        #region Main row
                        dr = dtGraph.NewRow();
                        dr["INITIAL"] = Initial;
                        dr["FINAL"] = arrStrike[0];
                        var nifty1 = arrStrike[0] - Initial;//Nifty <= 0 ? (Nifty - (arrStrike[0] % 10)) : (Nifty + (arrStrike[0] % 10));
                        dr["NIFTY_PERFORMANCE"] = Math.Round(nifty1, 2);

                        if (GraphType == "FixedPlusPR")
                        {
                            dblPR = CalculatePR(nifty1, Strike11, Strike12, Strike21, Strike22, BelowStrikeCoupon, AfterStrikeCoupon, CallOptionType, PutOptionType, FixedCouponValue, GraphType);
                            if (dblPR < NegSpot)
                            {
                                dr["FIXED"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                            }
                            else
                            {
                                dr["FIXED"] = Math.Round(dblPR, 2);
                            }
                        }
                        else if (GraphType == "FixedOrPR")
                        {
                            dblPR = CalculatePR(nifty1, Strike11, Strike12, Strike21, Strike22, BelowStrikeCoupon, AfterStrikeCoupon, CallOptionType, PutOptionType, FixedCouponValue, GraphType);
                            if (dblPR < NegSpot)
                            {
                                dr["FIXED"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                            }
                            else
                            {
                                dr["FIXED"] = Math.Round(dblPR, 2);
                            }
                        }

                        dtGraph.Rows.Add(dr);
                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["FIXED"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                        #endregion

                        #region next row
                        dr = dtGraph.NewRow();
                        dr["INITIAL"] = Initial;
                        final = arrStrike[0] + 0.01;
                        dr["FINAL"] = final;
                        var nextnifty = final - Initial;//Nifty <= 0 ? (Nifty - (arrStrike[0] % 10)) : (Nifty + (arrStrike[0] % 10));
                        //nextnifty = nextnifty + 0.01;
                        dr["NIFTY_PERFORMANCE"] = Math.Round(nextnifty, 2);

                        if (GraphType == "FixedPlusPR")
                        {
                            dblPR = CalculatePR(nextnifty, Strike11, Strike12, Strike21, Strike22, BelowStrikeCoupon, AfterStrikeCoupon, CallOptionType, PutOptionType, FixedCouponValue, GraphType);
                            if (dblPR < NegSpot)
                            {
                                dr["FIXED"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                            }
                            else
                            {
                                dr["FIXED"] = Math.Round(dblPR, 2);
                            }

                        }
                        else if (GraphType == "FixedOrPR")
                        {
                            dblPR = CalculatePR(nextnifty, Strike11, Strike12, Strike21, Strike22, BelowStrikeCoupon, AfterStrikeCoupon, CallOptionType, PutOptionType, FixedCouponValue, GraphType);
                            if (dblPR < NegSpot)
                            {
                                dr["FIXED"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                            }
                            else
                            {
                                dr["FIXED"] = Math.Round(dblPR, 2);
                            }
                        }

                        dtGraph.Rows.Add(dr);
                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["FIXED"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                        #endregion
                    }
                    else if (arrStrike[1] > PrevFinal && arrStrike[1] < Final)
                    {
                        Nifty = Nifty - 10;

                        #region Previous Row
                        //previous row
                        dr["INITIAL"] = Initial;
                        var final = arrStrike[1] - 0.01;
                        dr["FINAL"] = final;
                        var nifty = final - Initial;//Nifty <= 0 ? (Nifty - (arrStrike[1] % 10)) : (Nifty + (arrStrike[1] % 10));
                        //nifty = nifty - 0.01;
                        dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);
                        double dblPR = 0;

                        if (GraphType == "FixedPlusPR")
                        {
                            dblPR = CalculatePR(nifty, Strike11, Strike12, Strike21, Strike22, BelowStrikeCoupon, AfterStrikeCoupon, CallOptionType, PutOptionType, FixedCouponValue, GraphType);
                            if (dblPR < NegSpot)
                            {
                                dr["FIXED"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                            }
                            else
                            {
                                dr["FIXED"] = Math.Round(dblPR, 2);
                            }
                        }
                        else if (GraphType == "FixedOrPR")
                        {
                            dblPR = CalculatePR(nifty, Strike11, Strike12, Strike21, Strike22, BelowStrikeCoupon, AfterStrikeCoupon, CallOptionType, PutOptionType, FixedCouponValue, GraphType);
                            if (dblPR < NegSpot)
                            {
                                dr["FIXED"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                            }
                            else
                            {
                                dr["FIXED"] = Math.Round(dblPR, 2);
                            }
                        }

                        dtGraph.Rows.Add(dr);
                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["FIXED"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                        #endregion

                        #region Main row
                        dr = dtGraph.NewRow();
                        dr["INITIAL"] = Initial;
                        dr["FINAL"] = arrStrike[1];
                        var nifty1 = arrStrike[1] - Initial;//Nifty <= 0 ? (Nifty - (arrStrike[1] % 10)) : (Nifty + (arrStrike[1] % 10));
                        dr["NIFTY_PERFORMANCE"] = Math.Round(nifty1, 2);

                        if (GraphType == "FixedPlusPR")
                        {
                            dblPR = CalculatePR(nifty1, Strike11, Strike12, Strike21, Strike22, BelowStrikeCoupon, AfterStrikeCoupon, CallOptionType, PutOptionType, FixedCouponValue, GraphType);
                            if (dblPR < NegSpot)
                            {
                                dr["FIXED"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                            }
                            else
                            {
                                dr["FIXED"] = Math.Round(dblPR, 2);
                            }
                        }
                        else if (GraphType == "FixedOrPR")
                        {
                            dblPR = CalculatePR(nifty1, Strike11, Strike12, Strike21, Strike22, BelowStrikeCoupon, AfterStrikeCoupon, CallOptionType, PutOptionType, FixedCouponValue, GraphType);
                            if (dblPR < NegSpot)
                            {
                                dr["FIXED"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                            }
                            else
                            {
                                dr["FIXED"] = Math.Round(dblPR, 2);
                            }
                        }

                        dtGraph.Rows.Add(dr);
                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["FIXED"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                        #endregion

                        #region next row
                        dr = dtGraph.NewRow();
                        dr["INITIAL"] = Initial;
                        final = arrStrike[1] + 0.01;
                        dr["FINAL"] = final;
                        var nextnifty = final - Initial;//Nifty <= 0 ? (Nifty - (arrStrike[1] % 10)) : (Nifty + (arrStrike[1] % 10));
                        //nextnifty = nextnifty + 0.01;
                        dr["NIFTY_PERFORMANCE"] = Math.Round(nextnifty, 2);

                        if (GraphType == "FixedPlusPR")
                        {
                            dblPR = CalculatePR(nextnifty, Strike11, Strike12, Strike21, Strike22, BelowStrikeCoupon, AfterStrikeCoupon, CallOptionType, PutOptionType, FixedCouponValue, GraphType);
                            if (dblPR < NegSpot)
                            {
                                dr["FIXED"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                            }
                            else
                            {
                                dr["FIXED"] = Math.Round(dblPR, 2);
                            }

                        }
                        else if (GraphType == "FixedOrPR")
                        {
                            dblPR = CalculatePR(nextnifty, Strike11, Strike12, Strike21, Strike22, BelowStrikeCoupon, AfterStrikeCoupon, CallOptionType, PutOptionType, FixedCouponValue, GraphType);
                            if (dblPR < NegSpot)
                            {
                                dr["FIXED"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                            }
                            else
                            {
                                dr["FIXED"] = Math.Round(dblPR, 2);
                            }
                        }

                        dtGraph.Rows.Add(dr);
                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["FIXED"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                        #endregion
                    }
                    else if (arrStrike[2] > PrevFinal && arrStrike[2] < Final)
                    {
                        Nifty = Nifty - 10;

                        #region Previous Row
                        //previous row
                        dr["INITIAL"] = Initial;
                        var final = arrStrike[2] - 0.01;
                        dr["FINAL"] = final;
                        var nifty = final - Initial;//Nifty <= 0 ? (Nifty - (arrStrike[2] % 10)) : (Nifty + (arrStrike[2] % 10));
                        // nifty = nifty - 0.01;
                        dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);
                        double dblPR = 0;

                        if (GraphType == "FixedPlusPR")
                        {
                            dblPR = CalculatePR(nifty, Strike11, Strike12, Strike21, Strike22, BelowStrikeCoupon, AfterStrikeCoupon, CallOptionType, PutOptionType, FixedCouponValue, GraphType);
                            if (dblPR < NegSpot)
                            {
                                dr["FIXED"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                            }
                            else
                            {
                                dr["FIXED"] = Math.Round(dblPR, 2);
                            }
                        }
                        else if (GraphType == "FixedOrPR")
                        {
                            dblPR = CalculatePR(nifty, Strike11, Strike12, Strike21, Strike22, BelowStrikeCoupon, AfterStrikeCoupon, CallOptionType, PutOptionType, FixedCouponValue, GraphType);
                            if (dblPR < NegSpot)
                            {
                                dr["FIXED"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                            }
                            else
                            {
                                dr["FIXED"] = Math.Round(dblPR, 2);
                            }
                        }

                        dtGraph.Rows.Add(dr);
                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["FIXED"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                        #endregion

                        #region Main row
                        dr = dtGraph.NewRow();
                        dr["INITIAL"] = Initial;
                        dr["FINAL"] = arrStrike[2];
                        var nifty1 = arrStrike[2] - Initial;//Nifty <= 0 ? (Nifty - (arrStrike[2] % 10)) : (Nifty + (arrStrike[2] % 10));
                        dr["NIFTY_PERFORMANCE"] = Math.Round(nifty1, 2);

                        if (GraphType == "FixedPlusPR")
                        {
                            dblPR = CalculatePR(nifty1, Strike11, Strike12, Strike21, Strike22, BelowStrikeCoupon, AfterStrikeCoupon, CallOptionType, PutOptionType, FixedCouponValue, GraphType);
                            if (dblPR < NegSpot)
                            {
                                dr["FIXED"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                            }
                            else
                            {
                                dr["FIXED"] = Math.Round(dblPR, 2);
                            }
                        }
                        else if (GraphType == "FixedOrPR")
                        {
                            dblPR = CalculatePR(nifty1, Strike11, Strike12, Strike21, Strike22, BelowStrikeCoupon, AfterStrikeCoupon, CallOptionType, PutOptionType, FixedCouponValue, GraphType);
                            if (dblPR < NegSpot)
                            {
                                dr["FIXED"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                            }
                            else
                            {
                                dr["FIXED"] = Math.Round(dblPR, 2);
                            }
                        }

                        dtGraph.Rows.Add(dr);
                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["FIXED"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                        #endregion

                        #region next row
                        dr = dtGraph.NewRow();
                        dr["INITIAL"] = Initial;
                        final = arrStrike[2] + 0.01;
                        dr["FINAL"] = final;
                        var nextnifty = final - Initial;//Nifty <= 0 ? (Nifty - (arrStrike[2] % 10)) : (Nifty + (arrStrike[2] % 10));
                        // nextnifty = nextnifty + 0.01;
                        dr["NIFTY_PERFORMANCE"] = Math.Round(nextnifty, 2);

                        if (GraphType == "FixedPlusPR")
                        {
                            dblPR = CalculatePR(nextnifty, Strike11, Strike12, Strike21, Strike22, BelowStrikeCoupon, AfterStrikeCoupon, CallOptionType, PutOptionType, FixedCouponValue, GraphType);
                            if (dblPR < NegSpot)
                            {
                                dr["FIXED"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                            }
                            else
                            {
                                dr["FIXED"] = Math.Round(dblPR, 2);
                            }

                        }
                        else if (GraphType == "FixedOrPR")
                        {
                            dblPR = CalculatePR(nextnifty, Strike11, Strike12, Strike21, Strike22, BelowStrikeCoupon, AfterStrikeCoupon, CallOptionType, PutOptionType, FixedCouponValue, GraphType);
                            if (dblPR < NegSpot)
                            {
                                dr["FIXED"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                            }
                            else
                            {
                                dr["FIXED"] = Math.Round(dblPR, 2);
                            }
                        }

                        dtGraph.Rows.Add(dr);
                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["FIXED"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                        #endregion
                    }
                    else if (arrStrike[3] > PrevFinal && arrStrike[3] < Final)
                    {
                        Nifty = Nifty - 10;

                        #region Previous Row
                        //previous row
                        dr["INITIAL"] = Initial;
                        var final = arrStrike[3] - 0.01;
                        dr["FINAL"] = final;
                        var nifty = final - Initial;//Nifty <= 0 ? (Nifty - (arrStrike[3] % 10)) : (Nifty + (arrStrike[3] % 10));
                        //nifty = nifty - 0.01;
                        dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);
                        double dblPR = 0;

                        if (GraphType == "FixedPlusPR")
                        {
                            dblPR = CalculatePR(nifty, Strike11, Strike12, Strike21, Strike22, BelowStrikeCoupon, AfterStrikeCoupon, CallOptionType, PutOptionType, FixedCouponValue, GraphType);
                            if (dblPR < NegSpot)
                            {
                                dr["FIXED"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                            }
                            else
                            {
                                dr["FIXED"] = Math.Round(dblPR, 2);
                            }
                        }
                        else if (GraphType == "FixedOrPR")
                        {
                            dblPR = CalculatePR(nifty, Strike11, Strike12, Strike21, Strike22, BelowStrikeCoupon, AfterStrikeCoupon, CallOptionType, PutOptionType, FixedCouponValue, GraphType);
                            if (dblPR < NegSpot)
                            {
                                dr["FIXED"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                            }
                            else
                            {
                                dr["FIXED"] = Math.Round(dblPR, 2);
                            }
                        }

                        dtGraph.Rows.Add(dr);
                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["FIXED"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                        #endregion

                        #region Main row
                        dr = dtGraph.NewRow();
                        dr["INITIAL"] = Initial;
                        dr["FINAL"] = arrStrike[3];
                        var nifty1 = arrStrike[3] - Initial;//Nifty <= 0 ? (Nifty - (arrStrike[3] % 10)) : (Nifty + (arrStrike[3] % 10));
                        dr["NIFTY_PERFORMANCE"] = Math.Round(nifty1, 2);

                        if (GraphType == "FixedPlusPR")
                        {
                            dblPR = CalculatePR(nifty1, Strike11, Strike12, Strike21, Strike22, BelowStrikeCoupon, AfterStrikeCoupon, CallOptionType, PutOptionType, FixedCouponValue, GraphType);
                            if (dblPR < NegSpot)
                            {
                                dr["FIXED"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                            }
                            else
                            {
                                dr["FIXED"] = Math.Round(dblPR, 2);
                            }
                        }
                        else if (GraphType == "FixedOrPR")
                        {
                            dblPR = CalculatePR(nifty1, Strike11, Strike12, Strike21, Strike22, BelowStrikeCoupon, AfterStrikeCoupon, CallOptionType, PutOptionType, FixedCouponValue, GraphType);
                            if (dblPR < NegSpot)
                            {
                                dr["FIXED"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                            }
                            else
                            {
                                dr["FIXED"] = Math.Round(dblPR, 2);
                            }
                        }

                        dtGraph.Rows.Add(dr);
                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["FIXED"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                        #endregion

                        #region next row
                        dr = dtGraph.NewRow();
                        dr["INITIAL"] = Initial;
                        final = arrStrike[3] + 0.01;
                        dr["FINAL"] = final;
                        var nextnifty = final - Initial;//Nifty <= 0 ? (Nifty - (arrStrike[3] % 10)) : (Nifty + (arrStrike[3] % 10));
                        //nextnifty = nextnifty + 0.01;
                        dr["NIFTY_PERFORMANCE"] = Math.Round(nextnifty, 2);

                        if (GraphType == "FixedPlusPR")
                        {
                            dblPR = CalculatePR(nextnifty, Strike11, Strike12, Strike21, Strike22, BelowStrikeCoupon, AfterStrikeCoupon, CallOptionType, PutOptionType, FixedCouponValue, GraphType);
                            if (dblPR < NegSpot)
                            {
                                dr["FIXED"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                            }
                            else
                            {
                                dr["FIXED"] = Math.Round(dblPR, 2);
                            }

                        }
                        else if (GraphType == "FixedOrPR")
                        {
                            dblPR = CalculatePR(nextnifty, Strike11, Strike12, Strike21, Strike22, BelowStrikeCoupon, AfterStrikeCoupon, CallOptionType, PutOptionType, FixedCouponValue, GraphType);
                            if (dblPR < NegSpot)
                            {
                                dr["FIXED"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                            }
                            else
                            {
                                dr["FIXED"] = Math.Round(dblPR, 2);
                            }
                        }

                        dtGraph.Rows.Add(dr);
                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["FIXED"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                        #endregion
                    }
                    else if (arrStrike.Contains(Final))
                    {
                        if (strike4)
                            strike5 = true;

                        if (strike3)
                            strike4 = true;

                        if (strike2)
                            strike3 = true;

                        if (strike)
                            strike2 = true;

                        double dblPR = 0;

                        if (Final == 0)
                        {
                            #region Main row
                            dr = dtGraph.NewRow();
                            dr["INITIAL"] = Initial;
                            dr["FINAL"] = (double)(Initial + ((100 * Nifty) / 100));
                            dr["NIFTY_PERFORMANCE"] = Nifty;

                            if (GraphType == "FixedPlusPR")
                            {
                                dblPR = CalculatePR(Nifty, Strike11, Strike12, Strike21, Strike22, BelowStrikeCoupon, AfterStrikeCoupon, CallOptionType, PutOptionType, FixedCouponValue, GraphType);
                                if (dblPR < NegSpot)
                                {
                                    dr["FIXED"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                }
                                else
                                {
                                    dr["FIXED"] = Math.Round(dblPR, 2);
                                }
                            }
                            else if (GraphType == "FixedOrPR")
                            {
                                dblPR = CalculatePR(Nifty, Strike11, Strike12, Strike21, Strike22, BelowStrikeCoupon, AfterStrikeCoupon, CallOptionType, PutOptionType, FixedCouponValue, GraphType);
                                if (dblPR < NegSpot)
                                {
                                    dr["FIXED"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                }
                                else
                                {
                                    dr["FIXED"] = Math.Round(dblPR, 2);
                                }
                            }

                            dtGraph.Rows.Add(dr);
                            transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["FIXED"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                            #endregion
                        }
                        else
                        {
                            #region Previous Row
                            //previous row
                            dr["INITIAL"] = Initial;
                            var nifty = Nifty - 0.01;
                            dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);
                            var final = Initial + nifty;//(double)(Initial + ((100 * Nifty) / 100)) - 0.01;
                            dr["FINAL"] = Math.Round(final, 2);
                            //double dblPR = 0;
                            double dblFixed = 10;

                            if (GraphType == "FixedPlusPR")
                            {
                                dblPR = CalculatePR(nifty, Strike11, Strike12, Strike21, Strike22, BelowStrikeCoupon, AfterStrikeCoupon, CallOptionType, PutOptionType, FixedCouponValue, GraphType);
                                if (dblPR < NegSpot)
                                {
                                    dr["FIXED"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                }
                                else
                                {
                                    dr["FIXED"] = Math.Round(dblPR, 2);
                                }
                            }
                            else if (GraphType == "FixedOrPR")
                            {
                                dblPR = CalculatePR(nifty, Strike11, Strike12, Strike21, Strike22, BelowStrikeCoupon, AfterStrikeCoupon, CallOptionType, PutOptionType, FixedCouponValue, GraphType);
                                if (dblPR < NegSpot)
                                {
                                    dr["FIXED"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                }
                                else
                                {
                                    dr["FIXED"] = Math.Round(dblPR, 2);
                                }
                            }

                            dtGraph.Rows.Add(dr);
                            transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["FIXED"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                            #endregion

                            #region Main row
                            dr = dtGraph.NewRow();
                            dr["INITIAL"] = Initial;
                            dr["NIFTY_PERFORMANCE"] = Nifty;
                            dr["FINAL"] = Initial + Nifty;//(double)(Initial + ((100 * Nifty) / 100));

                            if (GraphType == "FixedPlusPR")
                            {
                                dblPR = CalculatePR(Nifty, Strike11, Strike12, Strike21, Strike22, BelowStrikeCoupon, AfterStrikeCoupon, CallOptionType, PutOptionType, FixedCouponValue, GraphType);
                                if (dblPR < NegSpot)
                                {
                                    dr["FIXED"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                }
                                else
                                {
                                    dr["FIXED"] = Math.Round(dblPR, 2);
                                }
                            }
                            else if (GraphType == "FixedOrPR")
                            {
                                dblPR = CalculatePR(Nifty, Strike11, Strike12, Strike21, Strike22, BelowStrikeCoupon, AfterStrikeCoupon, CallOptionType, PutOptionType, FixedCouponValue, GraphType);
                                if (dblPR < NegSpot)
                                {
                                    dr["FIXED"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                }
                                else
                                {
                                    dr["FIXED"] = Math.Round(dblPR, 2);
                                }
                            }

                            dtGraph.Rows.Add(dr);
                            transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["FIXED"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                            #endregion

                            #region next row
                            dr = dtGraph.NewRow();
                            dr["INITIAL"] = Initial;
                            var nextnifty = Nifty + 0.01;
                            dr["NIFTY_PERFORMANCE"] = Math.Round(nextnifty, 2);
                            final = Initial + nextnifty;//(double)(Initial + ((100 * Nifty) / 100)) + 0.01;
                            dr["FINAL"] = Math.Round(final, 2);

                            if (GraphType == "FixedPlusPR")
                            {
                                dblPR = CalculatePR(nextnifty, Strike11, Strike12, Strike21, Strike22, BelowStrikeCoupon, AfterStrikeCoupon, CallOptionType, PutOptionType, FixedCouponValue, GraphType);
                                if (dblPR < NegSpot)
                                {
                                    dr["FIXED"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                }
                                else
                                {
                                    dr["FIXED"] = Math.Round(dblPR, 2);
                                }

                            }
                            else if (GraphType == "FixedOrPR")
                            {
                                dblPR = CalculatePR(nextnifty, Strike11, Strike12, Strike21, Strike22, BelowStrikeCoupon, AfterStrikeCoupon, CallOptionType, PutOptionType, FixedCouponValue, GraphType);
                                if (dblPR < NegSpot)
                                {
                                    dr["FIXED"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                }
                                else
                                {
                                    dr["FIXED"] = Math.Round(dblPR, 2);
                                }
                            }

                            dtGraph.Rows.Add(dr);
                            transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["FIXED"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                            #endregion
                        }

                        strike = true;
                    }
                    else
                    {
                        dr["INITIAL"] = Initial;
                        dr["FINAL"] = Final;
                        dr["NIFTY_PERFORMANCE"] = Nifty;

                        double dblPR = 0;
                        double dblFixed = 10;

                        dblPR = CalculatePR(Nifty, Strike11, Strike12, Strike21, Strike22, BelowStrikeCoupon, AfterStrikeCoupon, CallOptionType, PutOptionType, FixedCouponValue, GraphType);
                        if (dblPR < NegSpot)
                        {
                            dr["FIXED"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                        }
                        else
                        {
                            dr["FIXED"] = Math.Round(dblPR, 2);
                        }

                        dtGraph.Rows.Add(dr);
                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["FIXED"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                    }
                    PrevFinal = Convert.ToInt32(Final);
                }
                //}

                return transactionCounts;
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];

                ClearFixedPlusPRSession();

                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "GenerateGraphFixedPRCalculation", objUserMaster.UserID);
                return transactionCounts;
            }
        }
        #endregion

        public double CalculatePR(double UnderlyingPerformance, double Strike11, double Strike12, double Strike21, double Strike22, double BelowStrikeCoupon, double AfterStrikeCoupon, string CallOptionType, string PutOptionType, double FixedCouponValue, string GraphType)
        {
            double dblPR = 0;
            double dblPR1 = 0;
            double dblPR2 = 0;
            double dblSpot = 100;

            if (GraphType == "FixedPlusPR")
            {
                if (CallOptionType == "Call Long")
                {
                    if (Strike11 == dblSpot)
                    {
                        dblPR1 = FixedCouponValue / 100 + Math.Max(0, BelowStrikeCoupon * ((UnderlyingPerformance / 100)));
                    }
                    else if (Strike11 > dblSpot)
                    {
                        dblPR1 = FixedCouponValue / 100 + Math.Max(0, BelowStrikeCoupon * ((UnderlyingPerformance / 100) - (Strike11 - dblSpot) / 100));
                        //Max(0 , ' + CallParticipatoryRatio + ' * (Underlying Performance - (' + (parseFloat(Strike1) - parseFloat(Spot)) + ')%) )';
                    }
                }
                else if (CallOptionType == "Call Spread Long")
                {
                    if (Strike12 > dblSpot && Strike11 == dblSpot)
                    {
                        dblPR1 = FixedCouponValue / 100 + Math.Max(0, BelowStrikeCoupon * Math.Min((Strike12 - dblSpot) / 100, UnderlyingPerformance / 100));
                        //' + CallParticipatoryRatio + ' * Min((' + (parseFloat(Strike2) - parseFloat(Spot)) + ')%, Underlying Performance))';
                    }
                    else if (Strike12 > dblSpot && Strike11 > dblSpot && Strike12 > Strike11)
                    {
                        dblPR1 = FixedCouponValue / 100 + Math.Max(0, BelowStrikeCoupon * Math.Min((Strike12 - Strike11) / 100, (UnderlyingPerformance / 100) - (Strike11 - dblSpot) / 100));
                        //Max(0 , ' + CallParticipatoryRatio + ' * Min((' + (parseFloat(Strike2) - parseFloat(Strike1)) + ')%, Underlying Performance - (' + (parseFloat(Strike1) - parseFloat(Spot)) + ')%))';
                    }
                }

                if (PutOptionType == "Put Short")
                {
                    if (Strike21 < dblSpot)
                    {
                        dblPR2 = Math.Min(0, AfterStrikeCoupon * (UnderlyingPerformance / 100 + (dblSpot - Strike21) / 100));

                        //'OTM PUT: +Min (0, ' + PutParticipatoryRatio + ' * (Underlying performance + (' + (parseFloat(Spot) - parseFloat(Strike1)) + ')% ))';
                    }
                    else
                    {
                        dblPR2 = Math.Min(0, AfterStrikeCoupon * UnderlyingPerformance / 100);

                        //'ATM Put: +Min (0,' + PutParticipatoryRatio + ' * Underlying Performance)';
                    }
                }
                else if (PutOptionType == "Put Spread Short")
                {

                    if (Strike22 < dblSpot && Strike22 < Strike21 && Strike21 == dblSpot)
                    {
                        dblPR2 = Math.Min(0, AfterStrikeCoupon * Math.Max((Strike22 - dblSpot) / 100, UnderlyingPerformance / 100));
                        //'ATM Capped Put: +Min (0, ' + PutParticipatoryRatio + ' * Max ((' + (parseFloat(Strike2) - parseFloat(Spot)) + ')%,  Underlying Performance))';
                    }
                    else if (Strike22 < dblSpot && Strike21 < dblSpot && Strike22 < Strike21)
                    {
                        dblPR2 = Math.Min(0, AfterStrikeCoupon * Math.Max((Strike22 - Strike21) / 100, UnderlyingPerformance / 100 + (dblSpot - Strike21) / 100));
                        //'OTM Capped PUT: +Min (0, ' + PutParticipatoryRatio + ' * Max((' + (parseFloat(Strike2) - parseFloat(Strike1)) + ')%, (Underlying performance+ (' + (parseFloat(Strike2) - parseFloat(Strike1)) + ')%)))';
                    }
                }
            }
            else
            {
                if (CallOptionType == "Call Long")
                {
                    dblPR1 = Math.Max(FixedCouponValue / 100, (BelowStrikeCoupon * (UnderlyingPerformance / 100)));
                }
                else if (CallOptionType == "Call Spread Long")
                {
                    if (Strike12 > dblSpot && Strike11 < Strike12)
                    {
                        dblPR1 = +Math.Max(FixedCouponValue / 100, (BelowStrikeCoupon * Math.Min((Strike12 - dblSpot) / 100, UnderlyingPerformance / 100)));
                    }
                }

                if (PutOptionType == "Put Short")
                {
                    if (Strike21 < dblSpot)
                    {
                        dblPR2 = Math.Min(0, AfterStrikeCoupon * (UnderlyingPerformance / 100 + (dblSpot - Strike21) / 100));
                    }
                    else
                    {
                        dblPR2 = Math.Min(0, AfterStrikeCoupon * UnderlyingPerformance / 100);
                    }
                }
                else if (PutOptionType == "Put Spread Short")
                {

                    if (Strike22 < dblSpot && Strike22 < Strike21 && Strike21 == dblSpot)
                    {
                        dblPR2 = Math.Min(0, AfterStrikeCoupon * Math.Max((Strike22 - dblSpot) / 100, UnderlyingPerformance / 100));
                    }
                    else if (Strike22 < dblSpot && Strike21 < dblSpot && Strike22 < Strike21)
                    {
                        dblPR2 = Math.Min(0, AfterStrikeCoupon * Math.Max((Strike22 - Strike21) / 100, UnderlyingPerformance / 100 + (dblSpot - Strike21) / 100));
                    }
                }
            }


            dblPR = dblPR1 + dblPR2;

            return dblPR * 100;
        }

        #region Fixed Or PR
        [HttpGet]
        public ActionResult FixedOrPR(string ProductID, string GenerateGraph, bool IsQuotron = false)
        {
            LoginController objLoginController = new LoginController();
            try
            {
                if (ValidateSession())
                {
                    FixedOrPR objFixedOrPR = new FixedOrPR();

                    #region Menu Access By on Role

                    Int32 intResult = 0;
                    // bool PPorNonPP = false;

                    UserMaster objUserMaster = new UserMaster();
                    objUserMaster = (UserMaster)Session["LoggedInUser"];

                    var Result = objSP_PRICINGEntities.VALIDATE_MENU_BY_ROLE(objUserMaster.RoleID, "BFOP");
                    intResult = Convert.ToInt32(Result.SingleOrDefault());

                    if (intResult == 0)
                        return RedirectToAction("UserNotAuthorize", "Login");

                    #endregion

                    #region Bind Underlying List
                    List<Underlying> UnderlyingList = new List<Underlying>();

                    ObjectResult<UnderlyingListResult> objUnderlyingListResult;
                    List<UnderlyingListResult> UnderlyingListResultList = new List<UnderlyingListResult>();

                    objUnderlyingListResult = objSP_PRICINGEntities.SP_FETCH_UNDERLYING_DETAILS();
                    UnderlyingListResultList = objUnderlyingListResult.ToList();

                    foreach (UnderlyingListResult oUnderlyingListResult in UnderlyingListResultList)
                    {
                        Underlying objUnderlying = new Underlying();
                        General.ReflectSingleData(objUnderlying, oUnderlyingListResult);

                        UnderlyingList.Add(objUnderlying);
                    }
                    objFixedOrPR.UnderlyingList = UnderlyingList;

                    //--Set default underlying--Added by Shweta on 3rd May 2016------------START--------------------
                    string strDefaultUnderlying = System.Configuration.ConfigurationManager.AppSettings["DefaultUnderlying"].ToUpper();
                    Underlying objDefaulyUnderlying = objFixedOrPR.UnderlyingList.Find(delegate(Underlying oUnderlying) { return oUnderlying.UnderlyingShortName == strDefaultUnderlying; });
                    objFixedOrPR.UnderlyingID = objDefaulyUnderlying.UnderlyingID;

                    objFixedOrPR.EntityID = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["DefaultEntityID"]);
                    objFixedOrPR.IsSecuredID = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["DefaultIsSecuredID"]);
                    //--Set default underlying--Added by Shweta on 3rd May 2016------------END----------------------
                    #endregion

                    if (ProductID != "" && ProductID != null)
                    {
                        ObjectResult<FixedOrPREditResult> objFixedOrPREditResult = objSP_PRICINGEntities.FETCH_FIXED_OR_PR_EDIT_DETAILS(ProductID);
                        List<FixedOrPREditResult> FixedOrPREditResultList = objFixedOrPREditResult.ToList();

                        General.ReflectSingleData(objFixedOrPR, FixedOrPREditResultList[0]);

                        DataSet dsResult = new DataSet();
                        dsResult = General.ExecuteDataSet("FETCH_UNDERLYING_CREATION_BYID", objFixedOrPR.UnderlyingID);

                        if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                        {
                            ViewBag.UnderlyingShortName = Convert.ToString(dsResult.Tables[0].Rows[0]["UNDERLYING_SHORTNAME"]);
                        }
                    }
                    else
                    {
                        objFixedOrPR.IsIRR = true;
                        objFixedOrPR.IsRedemptionPeriodMonth = true;
                    }

                    if (GenerateGraph == "GenerateGraph")
                    {
                        objFixedOrPR = (FixedOrPR)TempData["FixedOrGraph"];

                        ObjectResult<FixedOrPREditResult> objFixedOrPREditResult = objSP_PRICINGEntities.FETCH_FIXED_OR_PR_EDIT_DETAILS(objFixedOrPR.ProductID);
                        List<FixedOrPREditResult> FixedOrPREditResultList = objFixedOrPREditResult.ToList();
                        FixedOrPR oFixedOrPR = new FixedOrPR();
                        General.ReflectSingleData(oFixedOrPR, FixedOrPREditResultList[0]);


                        objFixedOrPR.Status = oFixedOrPR.Status;
                        // objFixedOrPR.SaveStatus = oFixedOrPR.SaveStatus;

                        return GenerateFixedOrPRGraph(objFixedOrPR);
                    }

                    else if (Session["FixedOrPRCopyQuote"] != null)
                    {
                        ViewBag.Message = "Quote is copied";
                        objFixedOrPR = (FixedOrPR)Session["FixedOrPRCopyQuote"];
                        objFixedOrPR.UnderlyingID = Convert.ToInt32(Session["UnderlyingID"]);

                        ObjectResult<FixedOrPREditResult> objFixedOrPREditResult = objSP_PRICINGEntities.FETCH_FIXED_OR_PR_EDIT_DETAILS(objFixedOrPR.ProductID);
                        List<FixedOrPREditResult> FixedOrPREditResultList = objFixedOrPREditResult.ToList();
                        FixedOrPR oFixedOrPR = new FixedOrPR();
                        if (FixedOrPREditResultList != null && FixedOrPREditResultList.Count > 0)
                            General.ReflectSingleData(oFixedOrPR, FixedOrPREditResultList[0]);

                        objFixedOrPR.ParentProductID = objFixedOrPR.ProductID;
                        objFixedOrPR.ProductID = "";
                        objFixedOrPR.Status = "";
                        objFixedOrPR.SaveStatus = "";
                        objFixedOrPR.IsCopyQuote = true;
                        objFixedOrPR.IsIRR = oFixedOrPR.IsIRR;
                        objFixedOrPR.IsRedemptionPeriodMonth = oFixedOrPR.IsRedemptionPeriodMonth;

                        objFixedOrPR.EntityID = oFixedOrPR.EntityID;
                        objFixedOrPR.IsSecuredID = oFixedOrPR.IsSecuredID;

                        //Added by Shweta on 10th May---------------START-------
                        objFixedOrPR.CallCustomIV1 = 0;
                        objFixedOrPR.CallCustomIV2 = 0;
                        objFixedOrPR.CallCustomRF1 = 0;
                        objFixedOrPR.CallCustomRF2 = 0;

                        objFixedOrPR.PutCustomIV1 = 0;
                        objFixedOrPR.PutCustomIV2 = 0;
                        objFixedOrPR.PutCustomRF1 = 0;
                        objFixedOrPR.PutCustomRF2 = 0;
                        //Added by Shweta on 10th May---------------END---------

                        //-------------Added by Shweta on 22nd July 2016 to Fetch Latest Deployment Rate------------START--------
                        string strDeploymentRate = "";
                        var DeploymentRate = objSP_PRICINGEntities.SP_FETCH_PRICING_DEPLOYMENT_RATE(Convert.ToInt32(objFixedOrPR.RedemptionPeriodDays), objFixedOrPR.EntityID, objFixedOrPR.IsSecuredID);
                        strDeploymentRate = Convert.ToString(DeploymentRate.SingleOrDefault());
                        objFixedOrPR.DeploymentRate = Convert.ToDouble(strDeploymentRate);
                        //-------------Added by Shweta on 22nd July 2016 to Fetch Latest Deployment Rate------------END----------
                    }

                    else if (Session["FixedOrPRChildQuote"] != null)
                    {
                        ViewBag.Message = true;
                        objFixedOrPR = (FixedOrPR)Session["FixedOrPRChildQuote"];
                        objFixedOrPR.UnderlyingID = Convert.ToInt32(Session["UnderlyingID"]);

                        ObjectResult<FixedOrPREditResult> objFixedOrPREditResult = objSP_PRICINGEntities.FETCH_FIXED_OR_PR_EDIT_DETAILS("");
                        List<FixedOrPREditResult> FixedOrPREditResultList = objFixedOrPREditResult.ToList();
                        FixedOrPR oFixedOrPR = new FixedOrPR();
                        if (FixedOrPREditResultList != null && FixedOrPREditResultList.Count > 0)
                            General.ReflectSingleData(oFixedOrPR, FixedOrPREditResultList[0]);

                        objFixedOrPR.ParentProductID = objFixedOrPR.ProductID;
                        objFixedOrPR.ProductID = "";
                        objFixedOrPR.Status = oFixedOrPR.Status;
                        objFixedOrPR.SaveStatus = oFixedOrPR.SaveStatus;
                        objFixedOrPR.IsChildQuote = true;
                    }
                    else if (Session["CancelQuote"] != null)
                    {
                        objFixedOrPR = (FixedOrPR)Session["CancelQuote"];

                        ObjectResult<FixedOrPREditResult> objFixedOrPREditResult = objSP_PRICINGEntities.FETCH_FIXED_OR_PR_EDIT_DETAILS(objFixedOrPR.ProductID);
                        List<FixedOrPREditResult> FixedOrPREditResultList = objFixedOrPREditResult.ToList();
                        FixedOrPR oFixedOrPR = new FixedOrPR();
                        if (FixedOrPREditResultList != null && FixedOrPREditResultList.Count > 0)
                            General.ReflectSingleData(oFixedOrPR, FixedOrPREditResultList[0]);

                        objFixedOrPR.Status = oFixedOrPR.Status;
                        objFixedOrPR.SaveStatus = oFixedOrPR.SaveStatus;

                        Session.Remove("CancelQuote");
                    }
                    else
                    {
                        Session.Remove("IsChildQuoteFixedOr");
                        Session.Remove("ParentProductID");
                        Session.Remove("UnderlyingID");
                    }

                    if (IsQuotron == true)
                    {
                        objFixedOrPR.IsQuotron = true;
                    }

                    if (Session["FixedOrPRChildQuote"] == null && Session["FixedOrPRCopyQuote"] == null)
                        objFixedOrPR.SaveStatus = "";

                    if (Session["FixedOrPRCopyQuote"] != null)
                        Session.Remove("FixedOrPRCopyQuote");

                    if (Session["FixedOrPRChildQuote"] != null)
                        Session.Remove("FixedOrPRChildQuote");

                    if (ProductID == null)
                    {
                        objFixedOrPR.isGraphActive = false;
                        return View(objFixedOrPR);
                    }
                    else
                    {
                        return GenerateFixedOrPRGraph(objFixedOrPR);
                    }
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

                ClearFixedOrPRSession();

                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "FixedOrPR Get", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        private ActionResult GenerateFixedOrPRGraph(FixedOrPR objFixedOrPR)
        {
            LoginController objLoginController = new LoginController();
            try
            {
                if (ValidateSession())
                {
                    var ParticipatoryRatio = TruncateDecimal(objFixedOrPR.CallPaticipatoryRatio, 2);

                    var transactionCounts = new List<Graph>();
                    transactionCounts = GenerateGraphFixedPRCalculation(objFixedOrPR.CallStrike1, objFixedOrPR.CallStrike2, objFixedOrPR.PutStrike1, objFixedOrPR.PutStrike2, ParticipatoryRatio, objFixedOrPR.PutPaticipatoryRatio, objFixedOrPR.CallOptionTypeId, objFixedOrPR.PutOptionTypeId, objFixedOrPR.FixedCouponValue, "FixedOrPR");

                    //FixedOrPR obj = new FixedOrPR();

                    #region Pie Chart For FixedOrPR
                    var xDataMonths = transactionCounts.Select(i => i.Column1).ToArray();
                    var yDataCounts = transactionCounts.Select(i => new object[] { i.Column2 }).ToArray();
                    var yDataCounts1 = transactionCounts.Select(i => new object[] { i.Column3 }).ToArray();



                    var FixedPRChart = new Highcharts("pie")
                        //define the type of chart 
                                .InitChart(new Chart { DefaultSeriesType = ChartTypes.Line })
                        //overall Title of the chart 
                                .SetTitle(new Title { Text = "Fixed OR PR" })
                        ////small label below the main Title
                        //        .SetSubtitle(new Subtitle { Text = "Accounting" })
                        //load the X values
                                .SetXAxis(new XAxis { Title = new XAxisTitle { Text = "Underlying Returns" }, Categories = xDataMonths, Labels = new XAxisLabels { Step = 2 } })
                        //set the Y title
                                .SetYAxis(new YAxis { Title = new YAxisTitle { Text = "Product Returns" } })
                                .SetTooltip(new Tooltip
                                {
                                    Enabled = true,
                                    Formatter = @"function() { return '<b>'+ this.series.name +'</b><br/>'+ this.x +': '+ this.y; }"
                                })
                                .SetPlotOptions(new PlotOptions
                                {
                                    Line = new PlotOptionsLine
                                    {
                                        DataLabels = new PlotOptionsLineDataLabels
                                        {
                                            Enabled = false
                                        },
                                        EnableMouseTracking = true
                                    }
                                })
                        //load the Y values 
                                .SetSeries(new[]
                    {
                        new Series {Name = "Coupon", Data = new Data(yDataCounts)},
                            //you can add more y data to create a second line
                            // new Series { Name = "Strike", Data = new Data(yDataCounts1) }
                    });
                    #endregion

                    if (Session["FixedOrPRCopyQuote"] != null)
                    {
                        ViewBag.Message = "Quote is copied";
                        objFixedOrPR = (FixedOrPR)Session["FixedOrPRCopyQuote"];
                        Session.Remove("FixedOrPRCopyQuote");
                    }

                    objFixedOrPR.FixedOrPRChart = FixedPRChart;

                    return View(objFixedOrPR);
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

                ClearFixedOrPRSession();

                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "GenerateFixedOrPRGraph", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        [HttpPost, ValidateInput(false)]
        public ActionResult FixedOrPR(string Command, FixedOrPR objFixedOrPR, FormCollection objFormCollection)
        {
            LoginController objLoginController = new LoginController();
            try
            {

                if (ValidateSession())
                {
                    #region Bind Underlying List
                    List<Underlying> UnderlyingList = new List<Underlying>();

                    ObjectResult<UnderlyingListResult> objUnderlyingListResult;
                    List<UnderlyingListResult> UnderlyingListResultList = new List<UnderlyingListResult>();

                    objUnderlyingListResult = objSP_PRICINGEntities.SP_FETCH_UNDERLYING_DETAILS();
                    UnderlyingListResultList = objUnderlyingListResult.ToList();

                    foreach (UnderlyingListResult oUnderlyingListResult in UnderlyingListResultList)
                    {
                        Underlying objUnderlying = new Underlying();
                        General.ReflectSingleData(objUnderlying, oUnderlyingListResult);

                        UnderlyingList.Add(objUnderlying);
                    }
                    objFixedOrPR.UnderlyingList = UnderlyingList;
                    #endregion

                    FixedOrPR oFixedOrPR = new FixedOrPR();
                    if (objFixedOrPR.ProductID != "" && objFixedOrPR.ProductID != null)
                    {
                        ObjectResult<FixedOrPREditResult> objFixedOrPREditResult = objSP_PRICINGEntities.FETCH_FIXED_OR_PR_EDIT_DETAILS(objFixedOrPR.ProductID);
                        List<FixedOrPREditResult> FixedOrPREditResultList = objFixedOrPREditResult.ToList();

                        General.ReflectSingleData(oFixedOrPR, FixedOrPREditResultList[0]);
                        objFixedOrPR.IsPrincipalProtected = oFixedOrPR.IsPrincipalProtected;
                        objFixedOrPR.CallOptionTypeId = oFixedOrPR.CallOptionTypeId;
                        objFixedOrPR.PutOptionTypeId = oFixedOrPR.PutOptionTypeId;
                    }

                    if (Command == "ExportToExcel")
                    {
                        ExportFixedOrPR(objFixedOrPR, objFormCollection);

                        return RedirectToAction("FixedOrPR");
                    }
                    else if (Command == "ExportCallStrike1Grid")
                    {
                        string StrikeHTML = objFormCollection["ExportCallStrike1Summary"];

                        ExportStrikeGrid(StrikeHTML);

                        return RedirectToAction("FixedOrPR");
                    }
                    else if (Command == "ExportCallStrike2Grid")
                    {
                        string StrikeHTML = objFormCollection["ExportCallStrike2Summary"];

                        ExportStrikeGrid(StrikeHTML);

                        return RedirectToAction("FixedOrPR");
                    }
                    else if (Command == "ExportPutStrike1Grid")
                    {
                        string StrikeHTML = objFormCollection["ExportPutStrike1Summary"];

                        ExportStrikeGrid(StrikeHTML);

                        return RedirectToAction("FixedOrPR");
                    }
                    else if (Command == "ExportPutStrike2Grid")
                    {
                        string StrikeHTML = objFormCollection["ExportPutStrike2Summary"];

                        ExportStrikeGrid(StrikeHTML);

                        return RedirectToAction("FixedOrPR");
                    }
                    else if (Command == "GenerateGraph")
                    {
                        objFixedOrPR.isGraphActive = true;
                        TempData["FixedOrGraph"] = objFixedOrPR;
                        return RedirectToAction("FixedOrPR", new { GenerateGraph = "GenerateGraph" });
                        //return GenerateFixedOrPRGraph(objFixedOrPR);
                    }

                    else if (Command == "CopyQuote")
                    {
                        Session["FixedOrPRCopyQuote"] = objFixedOrPR;
                        Session["UnderlyingID"] = objFixedOrPR.UnderlyingID;

                        return RedirectToAction("FixedOrPR");
                    }
                    else if (Command == "CreateChildQuote")
                    {
                        Session.Remove("ParentProductID");
                        Session.Remove("IsChildQuoteFixedOr");
                        Session.Remove("UnderlyingID");

                        Session["ParentProductID"] = objFixedOrPR.ProductID;
                        Session["UnderlyingID"] = objFixedOrPR.UnderlyingID;

                        objFixedOrPR.IsChildQuote = true;

                        Session["FixedOrPRChildQuote"] = objFixedOrPR;
                        Session["IsChildQuoteFixedOr"] = objFixedOrPR.IsChildQuote;

                        return RedirectToAction("FixedOrPR");
                    }
                    else if (Command == "AddNewProduct")
                    {
                        var productID = objFixedOrPR.ProductID;
                        UserMaster objUserMaster = new UserMaster();
                        objUserMaster = (UserMaster)Session["LoggedInUser"];

                        EncryptDecrypt obj = new EncryptDecrypt();
                        var encryptedpaswd = obj.Encrypt(objUserMaster.Password, "SPPricing", CryptographyEngine.AlgorithmType.DES);

                        var isPrincipalProtected = objFormCollection.Get("PrincipalProtected");
                        var ProductType = "PP";
                        if (isPrincipalProtected != "1")
                        {
                            ProductType = "NonPP";
                        }

                        var Url = "http://edemumnewuatvm4:63400/Login.aspx?UserId=" + objUserMaster.LoginName + "&Key=" + encryptedpaswd + "&ProductId=" + productID + "&ProductType=" + ProductType;
                        return Redirect(Url);
                    }
                    else if (Command == "Cancel")
                    {
                        Session["CancelQuote"] = objFixedOrPR;

                        return RedirectToAction("FixedOrPR");
                    }
                    else if (Command == "PricingInExcel")
                    {
                        objFixedOrPR.IsWorkingFileExport = OpenWorkingExcelFile("FOP", objFixedOrPR.ProductID);

                        if (!objFixedOrPR.IsWorkingFileExport)
                            objFixedOrPR.WorkingFileStatus = "File Not Found";

                        return View(objFixedOrPR);
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

                ClearFixedOrPRSession();

                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "FixedOrPR Post", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        public virtual void ExportFixedOrPR(FixedOrPR objFixedOrPR, FormCollection objFormCollection)
        {
            try
            {
                string strTemplateFilePath = Server.MapPath("~/Templates");
                string strTemplateFileName = strTemplateFilePath + "\\FixedOrPRTemplate.xlsx";

                string strTargetFilePath = Server.MapPath("~/OutputFiles");
                string strTargetFileName = strTargetFilePath + "\\" + objFixedOrPR.ProductID + "_FixedOrPR.xlsx";

                string Role = Convert.ToString(Session["Role"]);

                if (System.IO.File.Exists(strTargetFileName))
                    System.IO.File.Delete(strTargetFileName);

                FileInfo objTemplateFileInfo = new FileInfo(strTemplateFileName);
                objTemplateFileInfo.CopyTo(strTargetFileName);

                FileInfo objTargetFileInfo = new FileInfo(strTargetFileName);

                Underlying objUnderlying = objFixedOrPR.UnderlyingList.Find(delegate(Underlying oUnderlying) { return oUnderlying.UnderlyingID == objFixedOrPR.UnderlyingID; });

                using (var xlPackage = new ExcelPackage(objTargetFileInfo))
                {
                    var worksheet = xlPackage.Workbook.Worksheets["FixedOrPR"];

                    worksheet.Cell(1, 2).Value = objFixedOrPR.ProductID.ToString();
                    worksheet.Cell(1, 4).Value = objFixedOrPR.Distributor.ToString().ToUpper();
                    worksheet.Cell(1, 6).Value = objUnderlying.UnderlyingShortName;
                    if (objFixedOrPR.IsPrincipalProtected)
                        worksheet.Cell(1, 8).Value = "Yes";
                    else
                        worksheet.Cell(1, 8).Value = "No";

                    worksheet.Cell(2, 2).Formula = "=" + objFixedOrPR.EdelweissBuiltIn.ToString() + "%";
                    worksheet.Cell(2, 4).Formula = "=" + objFixedOrPR.DistributorBuiltIn.ToString() + "%";
                    worksheet.Cell(2, 6).Formula = "=B2+D2+H2";
                    worksheet.Cell(2, 8).Formula = "=" + objFixedOrPR.BuiltInAdjustment.ToString() + "%";

                    worksheet.Cell(3, 2).Formula = "=((POWER((1+D3),(12/D4))-1)*100) %";
                    worksheet.Cell(3, 4).Formula = "=" + objFixedOrPR.FixedCouponValue.ToString() + "%";

                    worksheet.Cell(4, 2).Value = objFixedOrPR.OptionTenureMonth.ToString();
                    worksheet.Cell(4, 4).Formula = "=ROUND(F4/30.417,0)";
                    worksheet.Cell(4, 6).Value = objFixedOrPR.RedemptionPeriodDays.ToString();

                    #region Get Entity Name
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
                    LookupMaster objLookupMasterEntity = EntityList.Find(delegate(LookupMaster oLookupMaster) { return oLookupMaster.LookupID == objFixedOrPR.EntityID; });
                    worksheet.Cell(5, 2).Value = objLookupMasterEntity.LookupDescription;
                    #endregion

                    #region Get Is Secured
                    objLookupResult = null;
                    LookupResultList = null;
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
                    LookupMaster objLookupMasterIsSecured = IsSecuredList.Find(delegate(LookupMaster oLookupMaster) { return oLookupMaster.LookupID == objFixedOrPR.IsSecuredID; });
                    worksheet.Cell(5, 4).Value = objLookupMasterIsSecured.LookupDescription;
                    #endregion
                    worksheet.Cell(5, 6).Formula = "=" + objFixedOrPR.DeploymentRate.ToString() + "%";
                    worksheet.Cell(5, 8).Formula = "=" + objFixedOrPR.CustomDeploymentRate.ToString() + "%";

                    worksheet.Cell(6, 2).Formula = "=(100-(B2+D2)*100)-(100*(1+ROUND(D3,4)))/(POWER((1+(IF(H5>0,H5,F5))),(D4/12)))";
                    worksheet.Cell(6, 4).Formula = "=H11 + H12";
                    worksheet.Cell(6, 6).Formula = "=B6 + D6";
                    worksheet.Cell(6, 8).Value = objFixedOrPR.StrikeCalculation.ToString();

                    worksheet.Cell(8, 2).Formula = objFixedOrPR.InitialAveragingMonth.ToString();
                    worksheet.Cell(8, 4).Value = objFixedOrPR.InitialAveragingDaysDiff.ToString();
                    worksheet.Cell(8, 6).Formula = objFixedOrPR.FinalAveragingMonth.ToString();
                    worksheet.Cell(8, 8).Value = objFixedOrPR.FinalAveragingDaysDiff.ToString();

                    if (objFixedOrPR.CallOptionTypeId != null)
                        worksheet.Cell(11, 1).Value = objFixedOrPR.CallOptionTypeId.ToString();
                    else
                        worksheet.Cell(11, 1).Value = "";
                    worksheet.Cell(11, 2).Value = objUnderlying.UnderlyingShortName;
                    worksheet.Cell(11, 3).Value = objFixedOrPR.CallStrike1.ToString();
                    worksheet.Cell(11, 4).Value = objFixedOrPR.CallStrike2.ToString();
                    worksheet.Cell(11, 5).Value = objFixedOrPR.CallPaticipatoryRatio.ToString();
                    worksheet.Cell(11, 6).Value = objFixedOrPR.CallPrice.ToString();
                    worksheet.Cell(11, 7).Value = objFixedOrPR.CallDiscountedPrice.ToString();
                    worksheet.Cell(11, 8).Value = objFixedOrPR.CallPrAdjustmentPrice.ToString();

                    if (Role == "Sales")
                    {
                        worksheet.Cell(10, 9).Value = "";
                        worksheet.Cell(10, 10).Value = "";
                        worksheet.Cell(10, 11).Value = "";
                        worksheet.Cell(10, 12).Value = "";
                        worksheet.Cell(10, 13).Value = "";
                        worksheet.Cell(10, 14).Value = "";
                        worksheet.Cell(10, 15).Value = "";
                        worksheet.Cell(10, 16).Value = "";
                    }
                    else
                    {
                        worksheet.Cell(11, 9).Formula = "=" + objFixedOrPR.CallIV1.ToString() + "%";
                        worksheet.Cell(11, 10).Formula = "=" + objFixedOrPR.CallCustomIV1.ToString() + "%";
                        worksheet.Cell(11, 11).Formula = "=" + objFixedOrPR.CallRF1.ToString() + "%";
                        worksheet.Cell(11, 12).Formula = "=" + objFixedOrPR.CallCustomRF1.ToString() + "%";
                        worksheet.Cell(11, 13).Formula = "=" + objFixedOrPR.CallIV2.ToString() + "%";
                        worksheet.Cell(11, 14).Formula = "=" + objFixedOrPR.CallCustomIV2.ToString() + "%";
                        worksheet.Cell(11, 15).Formula = "=" + objFixedOrPR.CallRF2.ToString() + "%";
                        worksheet.Cell(11, 16).Formula = "=" + objFixedOrPR.CallCustomRF2.ToString() + "%";

                    }

                    if (!objFixedOrPR.IsPrincipalProtected)
                        if (objFixedOrPR.PutStrike1 != 0)
                        {
                            if (objFixedOrPR.PutOptionTypeId != null)
                                worksheet.Cell(12, 1).Value = objFixedOrPR.PutOptionTypeId.ToString();
                            else
                                worksheet.Cell(12, 1).Value = "";
                            worksheet.Cell(12, 2).Value = objUnderlying.UnderlyingShortName;
                            worksheet.Cell(12, 3).Value = objFixedOrPR.PutStrike1.ToString();
                            worksheet.Cell(12, 4).Value = objFixedOrPR.PutStrike2.ToString();
                            worksheet.Cell(12, 5).Value = objFixedOrPR.PutPaticipatoryRatio.ToString();
                            worksheet.Cell(12, 6).Value = objFixedOrPR.PutPrice.ToString();
                            worksheet.Cell(12, 7).Value = objFixedOrPR.PutDiscountedPrice.ToString();
                            worksheet.Cell(12, 8).Value = objFixedOrPR.PutPrAdjustmentPrice.ToString();

                            if (Role == "Sales")
                            {
                                worksheet.Cell(10, 9).Value = "";
                                worksheet.Cell(10, 10).Value = "";
                                worksheet.Cell(10, 11).Value = "";
                                worksheet.Cell(10, 12).Value = "";
                                worksheet.Cell(10, 13).Value = "";
                                worksheet.Cell(10, 14).Value = "";
                                worksheet.Cell(10, 15).Value = "";
                                worksheet.Cell(10, 16).Value = "";
                            }
                            else
                            {
                                worksheet.Cell(12, 9).Formula = "=" + objFixedOrPR.PutIV1.ToString() + "%";
                                worksheet.Cell(12, 10).Formula = "=" + objFixedOrPR.PutCustomIV1.ToString() + "%";
                                worksheet.Cell(12, 11).Formula = "=" + objFixedOrPR.PutRF1.ToString() + "%";
                                worksheet.Cell(12, 12).Formula = "=" + objFixedOrPR.PutCustomRF1.ToString() + "%";
                                worksheet.Cell(12, 13).Formula = "=" + objFixedOrPR.PutIV2.ToString() + "%";
                                worksheet.Cell(12, 14).Formula = "=" + objFixedOrPR.PutCustomIV2.ToString() + "%";
                                worksheet.Cell(12, 15).Formula = "=" + objFixedOrPR.PutRF2.ToString() + "%";
                                worksheet.Cell(12, 16).Formula = "=" + objFixedOrPR.PutCustomRF2.ToString() + "%";
                            }
                        }

                    if (objFixedOrPR.SalesComments != null)
                        worksheet.Cell(14, 2).Value = objFixedOrPR.SalesComments.ToString();
                    else
                        worksheet.Cell(14, 2).Value = "";

                    if (objFixedOrPR.TradingComments != null)
                        worksheet.Cell(15, 2).Value = objFixedOrPR.TradingComments.ToString();
                    else
                        worksheet.Cell(15, 2).Value = "";

                    if (objFixedOrPR.CouponScenario1 != null)
                        worksheet.Cell(16, 2).Value = objFixedOrPR.CouponScenario1.ToString();
                    else
                        worksheet.Cell(16, 2).Value = "";

                    if (objFixedOrPR.CouponScenario2 != null)
                        worksheet.Cell(17, 2).Value = objFixedOrPR.CouponScenario2.ToString();
                    else
                        worksheet.Cell(17, 2).Value = "";

                    xlPackage.Save();
                }

                if (System.IO.File.Exists(strTargetFileName))
                {
                    FileInfo TemplateFile = new FileInfo(strTargetFileName);

                    Response.Clear();
                    Response.ClearHeaders();
                    Response.ClearContent();
                    Response.AddHeader("content-disposition", "attachment; filename=" + Path.GetFileName(strTargetFileName));
                    Response.AddHeader("Content-Type", "application/Excel");
                    Response.ContentType = "application/vnd.xls";
                    Response.AddHeader("Content-Length", TemplateFile.Length.ToString());
                    Response.WriteFile(TemplateFile.FullName);
                    Response.End();
                }
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];

                ClearFixedOrPRSession();

                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "ExportFixedOrPR", objUserMaster.UserID);

            }
        }

        public JsonResult ExportFixedOrPRWorkingFile(string ProductID, string Distributor, string EdelweissBuiltIn, string DistributorBuiltIn, string BuiltInAdjustment, string TotalBuiltIn, string IRR, string IsIRR, string DeploymentRate, string CustomerDeploymentRate, string FixedCoupon, string OptionTenureMonth, string RedemptionPeriodMonth, string IsRedemptionPeriodMonth, string RedemptionPeriodDays, string InitialAveragingMonth, string InitialAveragingDaysDiff, string FinalAveragingMonth, string FinalAveragingDaysDiff, string Underlying, string Remaining, string TotalOptionPrice, string NetRemaining, string SalesComments, string TradingComments, string CouponScenario1, string CouponScenario2,
            string CallOptionType, string CallStrike1, string CallStrike2, string CallParticipatoryRatio, string CallPrice, string CallDiscountedPrice, string CallPRAdjustedPrice, string CallIV1, string CallCustomIV1, string CallRF1, string CallCustomRF1, string CallIV2, string CallCustomIV2, string CallRF2, string CallCustomRF2,
            string PutOptionType, string PutStrike1, string PutStrike2, string PutParticipatoryRatio, string PutPrice, string PutDiscountedPrice, string PutPRAdjustedPrice, string PutIV1, string PutCustomIV1, string PutRF1, string PutCustomRF1, string PutIV2, string PutCustomIV2, string PutRF2, string PutCustomRF2, string IsPrincipalProtected,
            string CallStrike1Summary, string CallStrike2Summary, string PutStrike1Summary, string PutStrike2Summary,
            string ExportCallStrike1Summary, string ExportCallStrike2Summary, string ExportPutStrike1Summary, string ExportPutStrike2Summary, string Entity, string IsSecured, string StrikeCalculation)
        {
            try
            {
                string strTemplateFilePath = Server.MapPath("~/Templates");
                string strTemplateFileName = strTemplateFilePath + "\\FixedOrPRTemplateWorkingFile.xlsx";

                string strTargetFilePath = System.Configuration.ConfigurationManager.AppSettings["WorkingFilePath"];
                string strTargetFileName = strTargetFilePath + "\\" + ProductID + "_FixedOrPR.xlsx";

                string Role = Convert.ToString(Session["Role"]);

                if (System.IO.File.Exists(strTargetFileName))
                    System.IO.File.Delete(strTargetFileName);

                FileInfo objTemplateFileInfo = new FileInfo(strTemplateFileName);
                objTemplateFileInfo.CopyTo(strTargetFileName);

                FileInfo objTargetFileInfo = new FileInfo(strTargetFileName);

                using (var xlPackage = new ExcelPackage(objTargetFileInfo))
                {
                    var worksheet = xlPackage.Workbook.Worksheets["FixedOrPR"];

                    worksheet.Cell(1, 2).Value = ProductID.ToString();
                    worksheet.Cell(1, 4).Value = Distributor.ToString().ToUpper();
                    worksheet.Cell(1, 6).Value = Underlying;

                    if (IsPrincipalProtected == "True")
                        worksheet.Cell(1, 8).Value = "Yes";
                    else
                        worksheet.Cell(1, 8).Value = "No";

                    worksheet.Cell(2, 2).Formula = "=" + EdelweissBuiltIn.ToString() + "%";
                    worksheet.Cell(2, 4).Formula = "=" + DistributorBuiltIn.ToString() + "%";
                    worksheet.Cell(2, 6).Formula = "=B2+D2+H2";
                    worksheet.Cell(2, 8).Formula = "=" + BuiltInAdjustment.ToString() + "%";

                    if (IsIRR.ToUpper() == "TRUE")
                    {
                        worksheet.Cell(3, 2).Formula = "=" + IRR.ToString() + "%";
                        worksheet.Cell(3, 4).Formula = "=(POWER((1 + B3), (F4 / 365)) - 1)*100%";
                    }
                    else
                    {
                        worksheet.Cell(3, 2).Formula = "=((POWER((1+D3),(12/D4))-1) * 100) %";
                        worksheet.Cell(3, 4).Formula = "=" + FixedCoupon.ToString() + "%";
                    }

                    worksheet.Cell(4, 2).Value = OptionTenureMonth.ToString();
                    if (IsRedemptionPeriodMonth.ToUpper() == "TRUE")
                    {
                        worksheet.Cell(4, 4).Formula = RedemptionPeriodMonth;
                        worksheet.Cell(4, 6).Formula = "=ROUND(D4*30.417, 0)";
                    }
                    else
                    {
                        worksheet.Cell(4, 4).Formula = "=ROUND(F4/30.417,2)";
                        worksheet.Cell(4, 6).Formula = RedemptionPeriodDays.ToString();
                    }

                    #region Get Entity Name
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
                    LookupMaster objLookupMasterEntity = EntityList.Find(delegate(LookupMaster oLookupMaster) { return oLookupMaster.LookupID == Convert.ToInt32(Entity); });
                    worksheet.Cell(5, 2).Value = objLookupMasterEntity.LookupDescription;
                    #endregion

                    #region Get Is Secured
                    objLookupResult = null;
                    LookupResultList = null;
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
                    LookupMaster objLookupMasterIsSecured = IsSecuredList.Find(delegate(LookupMaster oLookupMaster) { return oLookupMaster.LookupID == Convert.ToInt32(IsSecured); });
                    worksheet.Cell(5, 4).Value = objLookupMasterIsSecured.LookupDescription;
                    #endregion
                    worksheet.Cell(5, 6).Formula = "=" + DeploymentRate.ToString() + "%";

                    if (CustomerDeploymentRate == "")
                        CustomerDeploymentRate = "0";
                    worksheet.Cell(5, 8).Formula = "=" + CustomerDeploymentRate.ToString() + "%";

                    worksheet.Cell(6, 2).Formula = "=(100-(B2+D2)*100)-(100*(1+ROUND(D3,4)))/(POWER((1+(IF(H5>0,H5,F5))),(D4/12)))";
                    worksheet.Cell(6, 4).Formula = "=H11 + H12";
                    worksheet.Cell(6, 6).Formula = "=ROUND(D6,4) + ROUND(B6,4)";
                    worksheet.Cell(6, 6).Value = StrikeCalculation.ToString(); ;

                    worksheet.Cell(8, 2).Value = InitialAveragingMonth.ToString();
                    worksheet.Cell(8, 4).Value = InitialAveragingDaysDiff.ToString();
                    worksheet.Cell(8, 6).Value = FinalAveragingMonth.ToString();
                    worksheet.Cell(8, 8).Value = FinalAveragingDaysDiff.ToString();

                    if (CallOptionType != null)
                        worksheet.Cell(11, 1).Value = CallOptionType.ToString();
                    else
                        worksheet.Cell(11, 1).Value = "";

                    worksheet.Cell(11, 2).Value = Underlying;
                    worksheet.Cell(11, 3).Value = CallStrike1.ToString();
                    worksheet.Cell(11, 4).Value = CallStrike2.ToString();

                    if (PutStrike1 != "" && PutStrike1 != "")
                        worksheet.Cell(11, 5).Formula = "=(B6+H12)/F11";
                    else
                        worksheet.Cell(11, 5).Formula = "=(B6+0)/F11";

                    if (CallStrike2 != "0" && CallStrike2 != "")
                        worksheet.Cell(11, 6).Formula = "=(AVERAGE(INDIRECT(\"$S$2:\"&ADDRESS(1+$B$8,18+$F$8))))-(AVERAGE(INDIRECT(\"$AA$2:\"&ADDRESS(1+$B$8,26+$F$8))))";
                    else
                        worksheet.Cell(11, 6).Formula = "=AVERAGE(INDIRECT(\"$S$2:\"&ADDRESS(1+$B$8,18+$F$8)))";

                    worksheet.Cell(11, 7).Formula = "=F11*-1";
                    worksheet.Cell(11, 8).Formula = "=ROUND(G11,4)*ROUND(E11,4)";

                    worksheet.Cell(11, 9).Formula = "=" + CallIV1.ToString() + "%";

                    if (CallCustomIV1 == "")
                        CallCustomIV1 = "0";
                    worksheet.Cell(11, 10).Formula = "=" + CallCustomIV1.ToString() + "%";

                    worksheet.Cell(11, 11).Formula = "=" + CallRF1.ToString() + "%";

                    if (CallCustomRF1 == "")
                        CallCustomRF1 = "0";
                    worksheet.Cell(11, 12).Formula = "=" + CallCustomRF1.ToString() + "%";

                    worksheet.Cell(11, 13).Formula = "=" + CallIV2.ToString() + "%";

                    if (CallCustomIV2 == "")
                        CallCustomIV2 = "0";
                    worksheet.Cell(11, 14).Formula = "=" + CallCustomIV2.ToString() + "%";

                    worksheet.Cell(11, 15).Formula = "=" + CallRF2.ToString() + "%";

                    if (CallCustomRF2 == "")
                        CallCustomRF2 = "0";
                    worksheet.Cell(11, 16).Formula = "=" + CallCustomRF2.ToString() + "%";

                    if (IsPrincipalProtected.ToUpper() == "FALSE")
                    {
                        if (PutStrike1 != "0")
                        {
                            if (PutOptionType != null)
                                worksheet.Cell(12, 1).Value = PutOptionType.ToString();
                            else
                                worksheet.Cell(12, 1).Value = "";

                            worksheet.Cell(12, 2).Value = Underlying;
                            worksheet.Cell(12, 3).Value = PutStrike1.ToString();
                            worksheet.Cell(12, 4).Value = PutStrike2.ToString();
                            worksheet.Cell(12, 5).Value = PutParticipatoryRatio.ToString();

                            if (PutStrike2 != "0" && PutStrike2 != "")
                                worksheet.Cell(12, 6).Formula = "=(AVERAGE(INDIRECT(\"$S$10:\"&ADDRESS(9+$B$8,18+$F$8))))-(AVERAGE(INDIRECT(\"$AA$10:\"&ADDRESS(9+$B$8,26+$F$8))))";
                            else
                                worksheet.Cell(12, 6).Formula = "=AVERAGE(INDIRECT(\"$S$10:\"&ADDRESS(9+$B$8,18+$F$8)))";

                            worksheet.Cell(12, 7).Formula = "=F12";
                            worksheet.Cell(12, 8).Formula = "=ROUND(G12,4)*ROUND(E12,4)";

                            worksheet.Cell(12, 9).Formula = "=" + PutIV1.ToString() + "%";

                            if (PutCustomIV1 == "")
                                PutCustomIV1 = "0";
                            worksheet.Cell(12, 10).Formula = "=" + PutCustomIV1.ToString() + "%";

                            worksheet.Cell(12, 11).Formula = "=" + PutRF1.ToString() + "%";

                            if (PutCustomRF1 == "")
                                PutCustomRF1 = "0";
                            worksheet.Cell(12, 12).Formula = "=" + PutCustomRF1.ToString() + "%";

                            worksheet.Cell(12, 13).Formula = "=" + PutIV2.ToString() + "%";

                            if (PutCustomIV2 == "")
                                PutCustomIV2 = "0";
                            worksheet.Cell(12, 14).Formula = "=" + PutCustomIV2.ToString() + "%";

                            worksheet.Cell(12, 15).Formula = "=" + PutRF2.ToString() + "%";

                            if (PutCustomRF2 == "")
                                PutCustomRF2 = "0";
                            worksheet.Cell(12, 16).Formula = "=" + PutCustomRF2.ToString() + "%";
                        }
                    }

                    if (SalesComments != null)
                        worksheet.Cell(14, 2).Value = SalesComments.ToString();
                    else
                        worksheet.Cell(14, 2).Value = "";

                    if (TradingComments != null)
                        worksheet.Cell(15, 2).Value = TradingComments.ToString();
                    else
                        worksheet.Cell(15, 2).Value = "";

                    if (CouponScenario1 != null)
                        worksheet.Cell(16, 2).Value = CouponScenario1.ToString();
                    else
                        worksheet.Cell(16, 2).Value = "";

                    if (CouponScenario2 != null)
                        worksheet.Cell(17, 2).Value = CouponScenario2.ToString();
                    else
                        worksheet.Cell(17, 2).Value = "";

                    //---------------Write Put Spread Strike 1 IV Grid-----------------START------------
                    worksheet.Cell(1, 18).Formula = "=C11";
                    worksheet.Cell(1, 19).Formula = "=ROUND($B$4*30.417,0)";
                    worksheet.Cell(1, 20).Formula = "=ROUND($B$4*30.417,0) - (1*$H$8)";
                    worksheet.Cell(1, 21).Formula = "=ROUND($B$4*30.417,0) - (2*$H$8)";
                    worksheet.Cell(1, 22).Formula = "=ROUND($B$4*30.417,0) - (3*$H$8)";
                    worksheet.Cell(1, 23).Formula = "=ROUND($B$4*30.417,0) - (4*$H$8)";
                    worksheet.Cell(1, 24).Formula = "=ROUND($B$4*30.417,0) - (5*$H$8)";

                    worksheet.Cell(2, 18).Formula = "0";
                    worksheet.Cell(3, 18).Formula = "=$R$2+1*$D$8";
                    worksheet.Cell(4, 18).Formula = "=$R$2+2*$D$8";
                    worksheet.Cell(5, 18).Formula = "=$R$2+3*$D$8";
                    worksheet.Cell(6, 18).Formula = "=$R$2+4*$D$8";
                    worksheet.Cell(7, 18).Formula = "=$R$2+5*$D$8";

                    worksheet.Cell(2, 19).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,S1,IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(3, 19).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($S$1-R3),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(4, 19).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($S$1-R4),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(5, 19).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($S$1-R5),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(6, 19).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($S$1-R6),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(7, 19).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($S$1-R7),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";

                    worksheet.Cell(2, 20).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,T1,IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(3, 20).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($T$1-R3),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(4, 20).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($T$1-R4),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(5, 20).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($T$1-R5),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(6, 20).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($T$1-R6),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(7, 20).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($T$1-R7),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";

                    worksheet.Cell(2, 21).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,U1,IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(3, 21).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($U$1-R3),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(4, 21).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($U$1-R4),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(5, 21).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($U$1-R5),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(6, 21).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($U$1-R6),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(7, 21).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($U$1-R7),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";

                    worksheet.Cell(2, 22).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,V1,IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(3, 22).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($V$1-R3),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(4, 22).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($V$1-R4),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(5, 22).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($V$1-R5),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(6, 22).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($V$1-R6),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(7, 22).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($V$1-R7),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";

                    worksheet.Cell(2, 23).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,W1,IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(3, 23).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($W$1-R3),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(4, 23).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($W$1-R4),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(5, 23).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($W$1-R5),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(6, 23).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($W$1-R6),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(7, 23).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($W$1-R7),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";

                    worksheet.Cell(2, 24).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,X1,IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(3, 24).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($X$1-R3),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(4, 24).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($X$1-R4),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(5, 24).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($X$1-R5),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(6, 24).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($X$1-R6),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(7, 24).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$C$11,100,($X$1-R7),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    //---------------Write Put Spread Strike 1 IV Grid-----------------END--------------

                    //---------------Write Put Spread Strike 2 IV Grid-----------------START------------
                    worksheet.Cell(1, 26).Formula = "=D11";
                    worksheet.Cell(1, 27).Formula = "=ROUND($B$4*30.417,0)";
                    worksheet.Cell(1, 28).Formula = "=ROUND($B$4*30.417,0) - (1*$H$8)";
                    worksheet.Cell(1, 29).Formula = "=ROUND($B$4*30.417,0) - (2*$H$8)";
                    worksheet.Cell(1, 30).Formula = "=ROUND($B$4*30.417,0) - (3*$H$8)";
                    worksheet.Cell(1, 31).Formula = "=ROUND($B$4*30.417,0) - (4*$H$8)";
                    worksheet.Cell(1, 32).Formula = "=ROUND($B$4*30.417,0) - (5*$H$8)";

                    worksheet.Cell(2, 26).Formula = "0";
                    worksheet.Cell(3, 26).Formula = "=$Z$2+1*$D$8";
                    worksheet.Cell(4, 26).Formula = "=$Z$2+2*$D$8";
                    worksheet.Cell(5, 26).Formula = "=$Z$2+3*$D$8";
                    worksheet.Cell(6, 26).Formula = "=$Z$2+4*$D$8";
                    worksheet.Cell(7, 26).Formula = "=$Z$2+5*$D$8";

                    worksheet.Cell(2, 27).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,AA1,IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(3, 27).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AA$1-Z3),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(4, 27).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AA$1-Z4),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(5, 27).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AA$1-Z5),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(6, 27).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AA$1-Z6),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(7, 27).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AA$1-Z7),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";

                    worksheet.Cell(2, 28).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,AB1,IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(3, 28).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AB$1-Z3),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(4, 28).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AB$1-Z4),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(5, 28).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AB$1-Z5),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(6, 28).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AB$1-Z6),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(7, 28).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AB$1-Z7),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";

                    worksheet.Cell(2, 29).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,AC1,IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(3, 29).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AC$1-Z3),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(4, 29).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AC$1-Z4),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(5, 29).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AC$1-Z5),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(6, 29).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AC$1-Z6),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(7, 29).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AC$1-Z7),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";

                    worksheet.Cell(2, 30).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,AD1,IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(3, 30).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AD$1-Z3),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(4, 30).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AD$1-Z4),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(5, 30).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AD$1-Z5),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(6, 30).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AD$1-Z6),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(7, 30).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AD$1-Z7),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";

                    worksheet.Cell(2, 31).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,AE1,IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(3, 31).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AE$1-Z3),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(4, 31).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AE$1-Z4),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(5, 31).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AE$1-Z5),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(6, 31).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AE$1-Z6),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(7, 31).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AE$1-Z7),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";

                    worksheet.Cell(2, 32).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,AF1,IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(3, 32).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AF$1-Z3),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(4, 32).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AF$1-Z4),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(5, 32).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AF$1-Z5),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(6, 32).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AF$1-Z6),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(7, 32).Formula = "=HoadleyOptions2(\"p\",1,\"C\",$D$11,100,($AF$1-Z7),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    //---------------Write Put Spread Strike 2 IV Grid-----------------END--------------

                    //---------------Write Put Strike 1 IV Grid-----------------START------------
                    if (PutStrike1 != "")
                    {
                        worksheet.Cell(9, 18).Formula = "=C12";
                        worksheet.Cell(9, 19).Formula = "=ROUND($B$4*30.417,0)";
                        worksheet.Cell(9, 20).Formula = "=ROUND($B$4*30.417,0) - (1*$H$8)";
                        worksheet.Cell(9, 21).Formula = "=ROUND($B$4*30.417,0) - (2*$H$8)";
                        worksheet.Cell(9, 22).Formula = "=ROUND($B$4*30.417,0) - (3*$H$8)";
                        worksheet.Cell(9, 23).Formula = "=ROUND($B$4*30.417,0) - (4*$H$8)";
                        worksheet.Cell(9, 24).Formula = "=ROUND($B$4*30.417,0) - (5*$H$8)";

                        worksheet.Cell(10, 18).Formula = "0";
                        worksheet.Cell(11, 18).Formula = "=$R$10+1*$D$8";
                        worksheet.Cell(12, 18).Formula = "=$R$10+2*$D$8";
                        worksheet.Cell(13, 18).Formula = "=$R$10+3*$D$8";
                        worksheet.Cell(14, 18).Formula = "=$R$10+4*$D$8";
                        worksheet.Cell(15, 18).Formula = "=$R$10+5*$D$8";

                        worksheet.Cell(10, 19).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,S9,IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(11, 19).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($S$9-R11),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(12, 19).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($S$9-R12),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(13, 19).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($S$9-R13),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(14, 19).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($S$9-R14),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(15, 19).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($S$9-R15),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";

                        worksheet.Cell(10, 20).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,T9,IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(11, 20).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($T$9-R11),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(12, 20).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($T$9-R12),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(13, 20).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($T$9-R13),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(14, 20).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($T$9-R14),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(15, 20).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($T$9-R15),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";

                        worksheet.Cell(10, 21).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,U9,IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(11, 21).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($U$9-R11),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(12, 21).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($U$9-R12),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(13, 21).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($U$9-R13),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(14, 21).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($U$9-R14),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(15, 21).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($U$9-R15),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";

                        worksheet.Cell(10, 22).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,V9,IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(11, 22).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($V$9-R11),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(12, 22).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($V$9-R12),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(13, 22).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($V$9-R13),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(14, 22).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($V$9-R14),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(15, 22).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($V$9-R15),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";

                        worksheet.Cell(10, 23).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,W9,IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(11, 23).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($W$9-R11),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(12, 23).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($W$9-R12),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(13, 23).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($W$9-R13),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(14, 23).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($W$9-R14),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(15, 23).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($W$9-R15),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";

                        worksheet.Cell(10, 24).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,X9,IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(11, 24).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($X$9-R11),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(12, 24).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($X$9-R12),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(13, 24).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($X$9-R13),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(14, 24).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($X$9-R14),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(15, 24).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($X$9-R15),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                    }
                    //---------------Write Put Strike 1 IV Grid-----------------END--------------

                    //---------------Write Put Strike 2 IV Grid-----------------START------------
                    if (PutStrike2 != "")
                    {
                        worksheet.Cell(9, 26).Formula = "=D11";
                        worksheet.Cell(9, 27).Formula = "=ROUND($B$4*30.417,0)";
                        worksheet.Cell(9, 28).Formula = "=ROUND($B$4*30.417,0) - (1*$H$8)";
                        worksheet.Cell(9, 29).Formula = "=ROUND($B$4*30.417,0) - (2*$H$8)";
                        worksheet.Cell(9, 30).Formula = "=ROUND($B$4*30.417,0) - (3*$H$8)";
                        worksheet.Cell(9, 31).Formula = "=ROUND($B$4*30.417,0) - (4*$H$8)";
                        worksheet.Cell(9, 32).Formula = "=ROUND($B$4*30.417,0) - (5*$H$8)";

                        worksheet.Cell(10, 26).Formula = "0";
                        worksheet.Cell(11, 26).Formula = "=$Z$10+1*$D$8";
                        worksheet.Cell(12, 26).Formula = "=$Z$10+2*$D$8";
                        worksheet.Cell(13, 26).Formula = "=$Z$10+3*$D$8";
                        worksheet.Cell(14, 26).Formula = "=$Z$10+4*$D$8";
                        worksheet.Cell(15, 26).Formula = "=$Z$10+5*$D$8";

                        worksheet.Cell(10, 27).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,AA9,IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(11, 27).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AA$9-Z11),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(12, 27).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AA$9-Z12),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(13, 27).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AA$9-Z13),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(14, 27).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AA$9-Z14),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(15, 27).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AA$9-Z15),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";

                        worksheet.Cell(10, 28).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,AB9,IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(11, 28).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AB$9-Z11),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(12, 28).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AB$9-Z12),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(13, 28).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AB$9-Z13),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(14, 28).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AB$9-Z14),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(15, 28).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AB$9-Z15),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";

                        worksheet.Cell(10, 29).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,AC9,IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(11, 29).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AC$9-Z11),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(12, 29).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AC$9-Z12),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(13, 29).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AC$9-Z13),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(14, 29).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AC$9-Z14),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(15, 29).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AC$9-Z15),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";

                        worksheet.Cell(10, 30).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,AD9,IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(11, 30).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AD$9-Z11),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(12, 30).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AD$9-Z12),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(13, 30).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AD$9-Z13),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(14, 30).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AD$9-Z14),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(15, 30).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AD$9-Z15),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";

                        worksheet.Cell(10, 31).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,AE9,IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(11, 31).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AE$9-Z11),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(12, 31).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AE$9-Z12),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(13, 31).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AE$9-Z13),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(14, 31).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AE$9-Z14),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(15, 31).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AE$9-Z15),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";

                        worksheet.Cell(10, 32).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,AF9,IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(11, 32).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AF$9-Z11),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(12, 32).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AF$9-Z12),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(13, 32).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AF$9-Z13),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(14, 32).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AF$9-Z14),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(15, 32).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AF$9-Z15),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                    }
                    //---------------Write Put Strike 2 IV Grid-----------------END--------------

                    xlPackage.Save();

                    return Json("");
                }
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];

                ClearFixedOrPRSession();

                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "ExportFixedPlusPR", objUserMaster.UserID);

                return Json("");
            }
        }

        public JsonResult ManageFixedOrPR(string ProductID, string Distributor, string EdelweissBuiltIn, string DistributorBuiltIn, string BuiltInAdjustment, string TotalBuiltIn, string IRR, string IsIRR, string DeploymentRate, string CustomerDeploymentRate, string FixedCoupon, string OptionTenureMonth, string RedemptionPeriodMonth, string IsRedemptionPeriodMonth, string RedemptionPeriodDays, string InitialAveragingMonth, string InitialAveragingDaysDiff, string FinalAveragingMonth, string FinalAveragingDaysDiff, string Underlying, string Remaining, string TotalOptionPrice, string NetRemaining, string SalesComments, string TradingComments, string CouponScenario1, string CouponScenario2,
            string CallOptionType, string CallStrike1, string CallStrike2, string CallParticipatoryRatio, string CallPrice, string CallDiscountedPrice, string CallPRAdjustedPrice, string CallIV1, string CallCustomIV1, string CallRF1, string CallCustomRF1, string CallIV2, string CallCustomIV2, string CallRF2, string CallCustomRF2,
            string PutOptionType, string PutStrike1, string PutStrike2, string PutParticipatoryRatio, string PutPrice, string PutDiscountedPrice, string PutPRAdjustedPrice, string PutIV1, string PutCustomIV1, string PutRF1, string PutCustomRF1, string PutIV2, string PutCustomIV2, string PutRF2, string PutCustomRF2,
            string StrikeCalculation, string IsPrincipalProtected, string CopyProductID, string CallStrike1Summary, string CallStrike2Summary, string PutStrike1Summary, string PutStrike2Summary,
            string ExportCallStrike1Summary, string ExportCallStrike2Summary, string ExportPutStrike1Summary, string ExportPutStrike2Summary, string Entity, string IsSecured)
        {
            try
            {
                ExportCallStrike1Summary = System.Uri.UnescapeDataString(ExportCallStrike1Summary);
                ExportCallStrike2Summary = System.Uri.UnescapeDataString(ExportCallStrike2Summary);
                ExportPutStrike1Summary = System.Uri.UnescapeDataString(ExportPutStrike1Summary);
                ExportPutStrike2Summary = System.Uri.UnescapeDataString(ExportPutStrike2Summary);
                CallStrike1Summary = System.Uri.UnescapeDataString(CallStrike1Summary);
                CallStrike2Summary = System.Uri.UnescapeDataString(CallStrike2Summary);
                PutStrike1Summary = System.Uri.UnescapeDataString(PutStrike1Summary);
                PutStrike2Summary = System.Uri.UnescapeDataString(PutStrike2Summary);

                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];

                if (PutOptionType == null)
                    PutOptionType = "";

                if (CustomerDeploymentRate == "")
                    CustomerDeploymentRate = "0";

                if (TotalOptionPrice == "")
                    TotalOptionPrice = "0";

                if (NetRemaining == "")
                    NetRemaining = "0";

                if (CallStrike1 == "")
                    CallStrike1 = "0";

                if (CallStrike2 == "")
                    CallStrike2 = "0";

                if (CallParticipatoryRatio == "")
                    CallParticipatoryRatio = "0";

                if (CallPrice == "")
                    CallPrice = "0";

                if (CallDiscountedPrice == "")
                    CallDiscountedPrice = "0";

                if (CallPRAdjustedPrice == "")
                    CallPRAdjustedPrice = "0";

                if (CallIV1 == "")
                    CallIV1 = "0";

                if (CallCustomIV1 == "")
                    CallCustomIV1 = "0";

                if (CallRF1 == "")
                    CallRF1 = "0";

                if (CallCustomRF1 == "")
                    CallCustomRF1 = "0";

                if (CallIV2 == "")
                    CallIV2 = "0";

                if (CallCustomIV2 == "")
                    CallCustomIV2 = "0";

                if (CallRF2 == "")
                    CallRF2 = "0";

                if (CallCustomRF2 == "")
                    CallCustomRF2 = "0";

                if (PutStrike1 == "")
                    PutStrike1 = "0";

                if (PutStrike2 == "")
                    PutStrike2 = "0";

                if (PutParticipatoryRatio == "" || PutParticipatoryRatio == "NaN")
                    PutParticipatoryRatio = "0";

                if (PutPrice == "" || PutPrice == "NaN")
                    PutPrice = "0";

                if (PutDiscountedPrice == "" || PutDiscountedPrice == "NaN")
                    PutDiscountedPrice = "0";

                if (PutPRAdjustedPrice == "" || PutPRAdjustedPrice == "NaN")
                    PutPRAdjustedPrice = "0";

                if (PutIV1 == "" || PutIV1 == "NaN")
                    PutIV1 = "0";

                if (PutCustomIV1 == "")
                    PutCustomIV1 = "0";

                if (PutRF1 == "" || PutRF1 == "NaN")
                    PutRF1 = "0";

                if (PutCustomRF1 == "")
                    PutCustomRF1 = "0";

                if (PutIV2 == "" || PutIV2 == "NaN")
                    PutIV2 = "0";

                if (PutCustomIV2 == "")
                    PutCustomIV2 = "0";

                if (PutRF2 == "" || PutRF2 == "NaN")
                    PutRF2 = "0";

                if (PutCustomRF2 == "")
                    PutCustomRF2 = "0";

                if (StrikeCalculation == "" || StrikeCalculation == "NaN")
                    StrikeCalculation = "0";

                string ParentProductID = "";
                if (Session["ParentProductID"] != null)
                    ParentProductID = (string)Session["ParentProductID"];

                ObjectResult<ManageFixedOrPRResult> objManageFixedOrPRResult = objSP_PRICINGEntities.SP_MANAGE_FIXED_OR_PR_DETAILS(ProductID, ParentProductID, Distributor, Convert.ToDouble(EdelweissBuiltIn), Convert.ToDouble(DistributorBuiltIn), Convert.ToDouble(BuiltInAdjustment), Convert.ToDouble(TotalBuiltIn), Convert.ToDouble(IRR), Convert.ToBoolean(IsIRR), Convert.ToDouble(DeploymentRate), Convert.ToDouble(CustomerDeploymentRate), Convert.ToDouble(FixedCoupon), Convert.ToInt32(OptionTenureMonth), Convert.ToDouble(RedemptionPeriodMonth), Convert.ToBoolean(IsRedemptionPeriodMonth), Convert.ToInt32(RedemptionPeriodDays), Convert.ToInt32(InitialAveragingMonth), Convert.ToInt32(InitialAveragingDaysDiff), Convert.ToInt32(FinalAveragingMonth), Convert.ToInt32(FinalAveragingDaysDiff), Convert.ToInt32(Underlying), Convert.ToDouble(Remaining), Convert.ToDouble(StrikeCalculation), Convert.ToDouble(TotalOptionPrice), Convert.ToDouble(NetRemaining), SalesComments, TradingComments, CouponScenario1, CouponScenario2, Convert.ToInt32(Entity), Convert.ToInt32(IsSecured), objUserMaster.UserID,
                    CallOptionType, Convert.ToDouble(CallStrike1), Convert.ToDouble(CallStrike2), Convert.ToDouble(CallParticipatoryRatio), Convert.ToDouble(CallPrice), Convert.ToDouble(CallDiscountedPrice), Convert.ToDouble(CallPRAdjustedPrice), Convert.ToDouble(CallIV1), Convert.ToDouble(CallCustomIV1), Convert.ToDouble(CallRF1), Convert.ToDouble(CallCustomRF1), Convert.ToDouble(CallIV2), Convert.ToDouble(CallCustomIV2), Convert.ToDouble(CallRF2), Convert.ToDouble(CallCustomRF2),
                    PutOptionType, Convert.ToDouble(PutStrike1), Convert.ToDouble(PutStrike2), Convert.ToDouble(PutParticipatoryRatio), Convert.ToDouble(PutPrice), Convert.ToDouble(PutDiscountedPrice), Convert.ToDouble(PutPRAdjustedPrice), Convert.ToDouble(PutIV1), Convert.ToDouble(PutCustomIV1), Convert.ToDouble(PutRF1), Convert.ToDouble(PutCustomRF1), Convert.ToDouble(PutIV2), Convert.ToDouble(PutCustomIV2), Convert.ToDouble(PutRF2), Convert.ToDouble(PutCustomRF2), Convert.ToBoolean(IsPrincipalProtected), CopyProductID,
                    CallStrike1Summary, CallStrike2Summary, PutStrike1Summary, PutStrike2Summary,
                    ExportCallStrike1Summary, ExportCallStrike2Summary, ExportPutStrike1Summary, ExportPutStrike2Summary
                    );
                List<ManageFixedOrPRResult> ManageFixedOrPRResultList = objManageFixedOrPRResult.ToList();

                return Json(ManageFixedOrPRResultList[0].ProductID);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];

                ClearFixedOrPRSession();

                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "ManageFixedOrPR", objUserMaster.UserID);
                return Json("");
            }
        }
        #endregion

        public bool OpenWorkingExcelFile(string PricerType, string ProductID)
        {
            try
            {
                string strPricerType = "";
                string strFileName = "";
                //string strFilePath = Server.MapPath("~/WorkingFiles/");
                string strFilePath = System.Configuration.ConfigurationManager.AppSettings["WorkingFilePath"];

                if (PricerType == "FC")
                    strPricerType = "FixedCoupon";
                else if (PricerType == "FCM")
                    strPricerType = "FixedCouponMLD";
                else if (PricerType == "FPP")
                    strPricerType = "FixedPlusPR";
                else if (PricerType == "FOP")
                    strPricerType = "FixedOrPR";
                else if (PricerType == "GC")
                    strPricerType = "GoldenCushion";
                else if (PricerType == "CB")
                    strPricerType = "CallBinary";
                else if (PricerType == "PB")
                    strPricerType = "PutBinary";

                strFileName = ProductID + "_" + strPricerType;
                strFilePath = strFilePath + strFileName + ".xlsx";

                if (System.IO.File.Exists(strFilePath))
                {
                    FileInfo TemplateFile = new FileInfo(strFilePath);

                    Response.Clear();
                    Response.ClearHeaders();
                    Response.ClearContent();
                    Response.AddHeader("content-disposition", "attachment; filename=" + Path.GetFileName(strFilePath));
                    Response.AddHeader("Content-Type", "application/Excel");
                    Response.ContentType = "application/vnd.xls";
                    Response.AddHeader("Content-Length", TemplateFile.Length.ToString());
                    Response.WriteFile(TemplateFile.FullName);
                    Response.End();

                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "FetchFinalIVRFValue", objUserMaster.UserID);

                return false;
            }
        }

        //public JsonResult OpenWorkingExcelFile(string PricerType, string ProductID)
        //{
        //    try
        //    {
        //        string strPricerType = "";
        //        string strFileName = "";
        //        string strFilePath = Server.MapPath("~/WorkingFiles/");

        //        if (PricerType == "FC")
        //            strPricerType = "FixedCoupon";
        //        else if (PricerType == "FCM")
        //            strPricerType = "FixedCouponMLD";
        //        else if (PricerType == "FPP")
        //            strPricerType = "FixedPlusPR";
        //        else if (PricerType == "FOP")
        //            strPricerType = "FixedOrPR";
        //        else if (PricerType == "GC")
        //            strPricerType = "GoldenCushion";
        //        else if (PricerType == "CB")
        //            strPricerType = "CallBinary";
        //        else if (PricerType == "PB")
        //            strPricerType = "PutBinary";

        //        strFileName = ProductID + "_" + strPricerType;
        //        strFilePath = strFilePath + strFileName + ".xlsx";

        //        if (System.IO.File.Exists(strFilePath))
        //        {
        //            FileInfo TemplateFile = new FileInfo(strFilePath);

        //            Response.Clear();
        //            Response.ClearHeaders();
        //            Response.ClearContent();
        //            Response.AddHeader("content-disposition", "attachment; filename=" + Path.GetFileName(strFilePath));
        //            Response.AddHeader("Content-Type", "application/Excel");
        //            Response.ContentType = "application/vnd.xls";
        //            Response.AddHeader("Content-Length", TemplateFile.Length.ToString());
        //            Response.WriteFile(TemplateFile.FullName);
        //            Response.End();

        //            return Json("");
        //        }
        //        else
        //            return Json("File Not Found");
        //    }
        //    catch (Exception ex)
        //    {
        //        UserMaster objUserMaster = new UserMaster();
        //        objUserMaster = (UserMaster)Session["LoggedInUser"];
        //        LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "FetchFinalIVRFValue", objUserMaster.UserID);
        //        return Json("");
        //    }
        //}

        public JsonResult FetchFinalIVRFValue(string Strike1, string Strike2, string Instrument, string Tenure, string UnderlyingID)
        {
            try
            {
                ObjectResult<FinalIVRFResult> objFinalIVRFResult = objSP_PRICINGEntities.SP_FETCH_FINAL_IV_RF_VALUE(Convert.ToDouble(Strike1), Convert.ToDouble(Strike2), Instrument, Convert.ToInt32(Tenure), Convert.ToInt32(UnderlyingID));
                List<FinalIVRFResult> FinalIVRFResultList = objFinalIVRFResult.ToList();

                return Json(FinalIVRFResultList[0]);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "FetchFinalIVRFValue", objUserMaster.UserID);
                return Json("");
            }
        }

        public JsonResult FetchFinalIVRFValue1(string Strike1, string Strike2, string Tenure, string UnderlyingID, string Instrument, string Strike2Flag, string Strike1IV)
        {
            try
            {
                Int32 intTenure = 0;

                if (Convert.ToBoolean(Strike2Flag))
                    intTenure = Convert.ToInt32(Tenure);
                else
                    intTenure = Convert.ToInt32(Convert.ToInt32(Tenure) * 30.417);

                ObjectResult<IVRF1Result> objIVRF1Result = objSP_PRICINGEntities.SP_FETCH_FINAL_IV_RF_VALUE_1(Convert.ToDouble(Strike1), Convert.ToDouble(Strike2), intTenure, Convert.ToInt32(UnderlyingID), Instrument, Convert.ToBoolean(Strike2Flag), Convert.ToDouble(Strike1IV));
                List<IVRF1Result> IVRF1ResultList = objIVRF1Result.ToList();

                return Json(IVRF1ResultList[0]);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "FetchFinalIVRFValue", objUserMaster.UserID);
                return Json("");
            }
        }

        public JsonResult CalculateGoldenCushionIV2(string Strike1, string Strike2, string Tenure, string UnderlyingID, string InitialAveragingMonth, string FinalAveragingMonth)
        {
            try
            {
                Int32 intTenure = 0;
                intTenure = Convert.ToInt32(Tenure);

                ObjectResult<RevisedIVCalculation> objRevisedIVCalculation = objSP_PRICINGEntities.SP_CALCULATE_GC_IV_2(Convert.ToDouble(Strike1), Convert.ToDouble(Strike2), intTenure, Convert.ToInt32(UnderlyingID), Convert.ToInt32(InitialAveragingMonth), Convert.ToInt32(FinalAveragingMonth));
                List<RevisedIVCalculation> RevisedIVCalculationList = objRevisedIVCalculation.ToList();

                return Json(RevisedIVCalculationList[0]);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "FetchFinalIVRFValue", objUserMaster.UserID);
                return Json("");
            }
        }

        public JsonResult CalculatePutAdjustment(string IV1, string Strike, string Tenure, string UnderlyingID)
        {
            try
            {
                double dblIV1 = 0;
                Int32 intTenure = 0;
                intTenure = Convert.ToInt32(Convert.ToInt32(Tenure) * 30.417);

                var NewIV1 = objSP_PRICINGEntities.SP_CALCULATE_PUT_ADJUSTEMENT(Convert.ToDouble(IV1), Convert.ToDouble(Strike), intTenure, Convert.ToInt32(UnderlyingID));
                dblIV1 = Convert.ToDouble(NewIV1.SingleOrDefault().Value);

                return Json(dblIV1);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "FetchFinalIVRFValue", objUserMaster.UserID);
                return Json("");
            }
        }

        public JsonResult GenerateNextNumber(string Type)
        {
            try
            {
                string strProductID = "";

                System.Data.Objects.ObjectParameter output = new System.Data.Objects.ObjectParameter("OutputParameterName", typeof(int));

                var NewProductID = objSP_PRICINGEntities.SP_GENERATE_NEXT_NUMBER(Type, output);
                strProductID = Convert.ToString(NewProductID.SingleOrDefault());

                return Json(strProductID);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "GenerateNextNumber", objUserMaster.UserID);
                return Json("");
            }
        }

        public JsonResult FetchPricingDeploymentRate(string Tenure, string Entity, string IsSecured)
        {
            try
            {
                string strDeploymentRate = "";

                var DeploymentRate = objSP_PRICINGEntities.SP_FETCH_PRICING_DEPLOYMENT_RATE(Convert.ToInt32(Tenure), Convert.ToInt32(Entity), Convert.ToInt32(IsSecured));
                strDeploymentRate = Convert.ToString(DeploymentRate.SingleOrDefault());

                return Json(strDeploymentRate);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "FetchPricingDeploymentRate", objUserMaster.UserID);
                return Json("");
            }
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
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "FetchUnderlyingList", objUserMaster.UserID);
                return Json("");
            }

            //return Json(UnderlyingListData);
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

        public JsonResult ManagePricerStatusLog(string PricerType, string ProductID, string StatusCode, string IsNonPP)
        {
            try
            {
                Int32 intResult = 0;
                // bool PPorNonPP = false;

                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];

                var Result = objSP_PRICINGEntities.SP_MANAGE_PRICER_STATUS_LOG(PricerType, ProductID, objUserMaster.UserID, StatusCode);
                intResult = Convert.ToInt32(Result.SingleOrDefault());

                if (StatusCode == "AP")
                {
                    var productID = ProductID;

                    EncryptDecrypt obj = new EncryptDecrypt();
                    var encryptedpaswd = obj.Encrypt(objUserMaster.Password, "SPPricing", CryptographyEngine.AlgorithmType.DES);

                    // var isPrincipalProtected = objSP_PRICINGEntities.SP_FETCH_IS_PROTECTED(PricerType, ProductID);
                    // PPorNonPP = Convert.ToBoolean(isPrincipalProtected.SingleOrDefault());

                    string spPortalLink = "";

                    var Link = objSP_PRICINGEntities.SP_FETCH_SP_PORTAL_LINK("SPPortalLink");
                    spPortalLink = Convert.ToString(Link.SingleOrDefault());

                    spPortalLink = spPortalLink.Replace("[LoginName]", objUserMaster.LoginName);
                    spPortalLink = spPortalLink.Replace("[EncryptedPaswd]", encryptedpaswd);
                    spPortalLink = spPortalLink.Replace("[ProductID]", productID);


                    var ProductType = "NonPP";

                    if (IsNonPP != null)
                    {
                        if (IsNonPP.ToUpper() == "FALSE")
                            ProductType = "PP";
                    }
                    else
                        ProductType = "PP";

                    //var Url = "http://edemumnewuatvm4:63400/Login.aspx?UserId=" + objUserMaster.LoginName + "&Key=" + encryptedpaswd + "&ProductId=" + productID + "&ProductType=" + ProductType;

                    spPortalLink = spPortalLink.Replace("[ProductType]", ProductType);

                    var Url = spPortalLink;

                    // Process.Start("iexplore", Url);
                    return Json(Url);
                }


                return Json(intResult);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "ManagePricerStatusLog", objUserMaster.UserID);
                return Json("");
            }
        }

        public JsonResult ManageFavouriteQuotes(string PricerType, string ProductID, string IsFavourite)
        {
            try
            {
                Int32 intResult = 0;

                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];

                var Result = objSP_PRICINGEntities.SP_MANAGE_FAVOURITE_PRODUCT(PricerType, ProductID, Convert.ToBoolean(IsFavourite), objUserMaster.UserID);
                intResult = Convert.ToInt32(Result.SingleOrDefault());

                return Json(intResult);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "ManagePricerStatusLog", objUserMaster.UserID);
                return Json("");
            }
        }

        public JsonResult SendDistributorMail(string PricerType, string ProductID)
        {
            UserMaster objUserMaster = new UserMaster();
            objUserMaster = (UserMaster)Session["LoggedInUser"];

            try
            {
                Int32 intResult = 0;

                var Result = objSP_PRICINGEntities.SP_SEND_DISTRIBUTOR_MAIL(PricerType, ProductID, objUserMaster.Email);
                intResult = Convert.ToInt32(Result.SingleOrDefault());

                return Json(intResult);
            }
            catch (Exception ex)
            {
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "SendDistributorMail", objUserMaster.UserID);
                return Json("");
            }
        }

        public JsonResult FetchToleranceLevel(string PricerType)
        {
            try
            {
                ObjectResult<ToleranceLevelResult> objToleranceLevelResult = objSP_PRICINGEntities.SP_FETCH_TOLERANCE_LEVEL(PricerType);
                List<ToleranceLevelResult> ToleranceLevelResultList = objToleranceLevelResult.ToList();

                //var ToleranceLevelResultListData = JsonConvert.SerializeObject(ToleranceLevelResultList[0]);
                return Json(ToleranceLevelResultList, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "FetchToleranceLevel", objUserMaster.UserID);
                return Json("");
            }
        }

        public JsonResult FetchPricerStatus(string PricerType, string ProductID)
        {
            string strStatus = "";

            if (PricerType == "FC")
            {
                ObjectResult<FixedCouponEditResult> objFixedCouponEditResult = objSP_PRICINGEntities.FETCH_FIXED_COUPON_EDIT_DETAILS(ProductID);
                List<FixedCouponEditResult> FixedCouponEditResultList = objFixedCouponEditResult.ToList();

                FixedCoupon oFixedCoupon = new FixedCoupon();
                General.ReflectSingleData(oFixedCoupon, FixedCouponEditResultList[0]);

                strStatus = oFixedCoupon.Status;
            }
            else if (PricerType == "FCM")
            {
                ObjectResult<FixedMLDEditResult> objFixedMLDEditResult = objSP_PRICINGEntities.FETCH_FIXED_COUPON_MLD_EDIT_DETAILS(ProductID);
                List<FixedMLDEditResult> FixedMLDEditResultList = objFixedMLDEditResult.ToList();
                FixedCouponMLD oFixedCouponMLD = new FixedCouponMLD();
                General.ReflectSingleData(oFixedCouponMLD, FixedMLDEditResultList[0]);

                strStatus = oFixedCouponMLD.Status;
            }
            else if (PricerType == "FPP")
            {
                ObjectResult<FixedPlusPREditResult> objFixedPlusPREditResult = objSP_PRICINGEntities.FETCH_FIXED_PLUS_PR_EDIT_DETAILS(ProductID);
                List<FixedPlusPREditResult> FixedPlusPREditResultList = objFixedPlusPREditResult.ToList();
                FixedPlusPR oFixedPlusPR = new FixedPlusPR();
                General.ReflectSingleData(oFixedPlusPR, FixedPlusPREditResultList[0]);

                strStatus = oFixedPlusPR.Status;
            }
            else if (PricerType == "FOP")
            {
                ObjectResult<FixedOrPREditResult> objFixedOrPREditResult = objSP_PRICINGEntities.FETCH_FIXED_OR_PR_EDIT_DETAILS(ProductID);
                List<FixedOrPREditResult> FixedOrPREditResultList = objFixedOrPREditResult.ToList();
                FixedOrPR oFixedOrPR = new FixedOrPR();
                General.ReflectSingleData(oFixedOrPR, FixedOrPREditResultList[0]);

                strStatus = oFixedOrPR.Status;
            }
            else if (PricerType == "GC")
            {
                ObjectResult<GoldenCushionEditResult> objGoldenCushionEditResult = objSP_PRICINGEntities.FETCH_GOLDEN_CUSHION_EDIT_DETAILS(ProductID);
                List<GoldenCushionEditResult> GoldenCushionEditResultList = objGoldenCushionEditResult.ToList();
                GoldenCushion oGoldenCushion = new GoldenCushion();
                General.ReflectSingleData(oGoldenCushion, GoldenCushionEditResultList[0]);

                strStatus = oGoldenCushion.Status;
            }
            else if (PricerType == "CB")
            {
                ObjectResult<CallBinaryEditResult> objCallBinaryEditResult = objSP_PRICINGEntities.FETCH_CALL_BINARY_EDIT_DETAILS(ProductID);
                List<CallBinaryEditResult> CallBinaryEditResultList = objCallBinaryEditResult.ToList();
                CallBinary oCallBinary = new CallBinary();
                General.ReflectSingleData(oCallBinary, CallBinaryEditResultList[0]);

                strStatus = oCallBinary.Status;
            }
            else if (PricerType == "PB")
            {
                ObjectResult<PutBinaryEditResult> objPutBinaryEditResult = objSP_PRICINGEntities.FETCH_PUT_BINARY_EDIT_DETAILS(ProductID);
                List<PutBinaryEditResult> PutBinaryEditResultList = objPutBinaryEditResult.ToList();
                PutBinary oPutBinary = new PutBinary();
                General.ReflectSingleData(oPutBinary, PutBinaryEditResultList[0]);

                strStatus = oPutBinary.Status;
            }
            return Json(strStatus);
        }

        #region Golden Cushion
        [HttpGet]
        public ActionResult GoldenCushion(string ProductID, string GenerateGraph, bool IsQuotron = false)
        {
            LoginController objLoginController = new LoginController();
            try
            {
                if (ValidateSession())
                {
                    GoldenCushion objGoldenCushion = new GoldenCushion();

                    #region Menu Access By on Role

                    Int32 intResult = 0;
                    // bool PPorNonPP = false;

                    UserMaster objUserMaster = new UserMaster();
                    objUserMaster = (UserMaster)Session["LoggedInUser"];

                    var Result = objSP_PRICINGEntities.VALIDATE_MENU_BY_ROLE(objUserMaster.RoleID, "BGC");
                    intResult = Convert.ToInt32(Result.SingleOrDefault());

                    if (intResult == 0)
                        return RedirectToAction("UserNotAuthorize", "Login");

                    #endregion

                    #region Bind Underlying List
                    List<Underlying> UnderlyingList = new List<Underlying>();

                    ObjectResult<UnderlyingListResult> objUnderlyingListResult;
                    List<UnderlyingListResult> UnderlyingListResultList = new List<UnderlyingListResult>();

                    objUnderlyingListResult = objSP_PRICINGEntities.SP_FETCH_UNDERLYING_DETAILS();
                    UnderlyingListResultList = objUnderlyingListResult.ToList();

                    foreach (UnderlyingListResult oUnderlyingListResult in UnderlyingListResultList)
                    {
                        Underlying objUnderlying = new Underlying();
                        General.ReflectSingleData(objUnderlying, oUnderlyingListResult);

                        UnderlyingList.Add(objUnderlying);
                    }
                    objGoldenCushion.UnderlyingList = UnderlyingList;

                    //--Set default underlying--Added by Shweta on 3rd May 2016------------START--------------------
                    string strDefaultUnderlying = System.Configuration.ConfigurationManager.AppSettings["DefaultUnderlying"].ToUpper();
                    Underlying objDefaulyUnderlying = objGoldenCushion.UnderlyingList.Find(delegate(Underlying oUnderlying) { return oUnderlying.UnderlyingShortName == strDefaultUnderlying; });
                    objGoldenCushion.UnderlyingID = objDefaulyUnderlying.UnderlyingID;

                    objGoldenCushion.EntityID = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["DefaultEntityID"]);
                    objGoldenCushion.IsSecuredID = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["DefaultIsSecuredID"]);
                    //--Set default underlying--Added by Shweta on 3rd May 2016------------END----------------------
                    #endregion

                    if (ProductID != "" && ProductID != null)
                    {
                        ObjectResult<GoldenCushionEditResult> objGoldenCushionEditResult = objSP_PRICINGEntities.FETCH_GOLDEN_CUSHION_EDIT_DETAILS(ProductID);
                        List<GoldenCushionEditResult> GoldenCushionEditResultList = objGoldenCushionEditResult.ToList();

                        General.ReflectSingleData(objGoldenCushion, GoldenCushionEditResultList[0]);

                        DataSet dsResult = new DataSet();
                        dsResult = General.ExecuteDataSet("FETCH_UNDERLYING_CREATION_BYID", objGoldenCushion.UnderlyingID);

                        if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                        {
                            ViewBag.UnderlyingShortName = Convert.ToString(dsResult.Tables[0].Rows[0]["UNDERLYING_SHORTNAME"]);
                        }
                    }
                    else
                    {
                        objGoldenCushion.IsFixedCouponIRR = true;
                        objGoldenCushion.IsLowerCouponIRR = true;
                        objGoldenCushion.IsRedemptionPeriodMonth = true;
                    }

                    if (GenerateGraph == "GenerateGraph")
                    {
                        objGoldenCushion = (GoldenCushion)TempData["GoldenCushionGraph"];

                        ObjectResult<GoldenCushionEditResult> objGoldenCushionEditResult = objSP_PRICINGEntities.FETCH_GOLDEN_CUSHION_EDIT_DETAILS(objGoldenCushion.ProductID);
                        List<GoldenCushionEditResult> GoldenCushionEditResultList = objGoldenCushionEditResult.ToList();
                        GoldenCushion oGoldenCushion = new GoldenCushion();
                        General.ReflectSingleData(oGoldenCushion, GoldenCushionEditResultList[0]);

                        objGoldenCushion.Status = oGoldenCushion.Status;
                        //objGoldenCushion.SaveStatus = oGoldenCushion.SaveStatus;

                        return GenerateGoldenCushionGraph(objGoldenCushion);
                    }

                    else if (Session["GoldenCushionCopyQuote"] != null)
                    {
                        ViewBag.Message = "Quote is copied";
                        objGoldenCushion = (GoldenCushion)Session["GoldenCushionCopyQuote"];
                        objGoldenCushion.UnderlyingID = Convert.ToInt32(Session["UnderlyingID"]);

                        ObjectResult<GoldenCushionEditResult> objGoldenCushionEditResult = objSP_PRICINGEntities.FETCH_GOLDEN_CUSHION_EDIT_DETAILS(objGoldenCushion.ProductID);
                        List<GoldenCushionEditResult> GoldenCushionEditResultList = objGoldenCushionEditResult.ToList();
                        GoldenCushion oGoldenCushion = new GoldenCushion();
                        if (GoldenCushionEditResultList != null && GoldenCushionEditResultList.Count > 0)
                            General.ReflectSingleData(oGoldenCushion, GoldenCushionEditResultList[0]);

                        objGoldenCushion.ParentProductID = objGoldenCushion.ProductID;
                        objGoldenCushion.ProductID = "";
                        objGoldenCushion.Status = "";
                        objGoldenCushion.SaveStatus = "";
                        objGoldenCushion.IsCopyQuote = true;
                        objGoldenCushion.IsFixedCouponIRR = oGoldenCushion.IsFixedCouponIRR;
                        objGoldenCushion.IsLowerCouponIRR = oGoldenCushion.IsLowerCouponIRR;
                        objGoldenCushion.IsRedemptionPeriodMonth = oGoldenCushion.IsRedemptionPeriodMonth;

                        //Added by Shweta on 10th May---------------START-------
                        objGoldenCushion.PutSpreadCustomIV1 = 0;
                        objGoldenCushion.PutSpreadCustomIV2 = 0;
                        objGoldenCushion.PutSpreadCustomRF1 = 0;
                        objGoldenCushion.PutSpreadCustomRF2 = 0;

                        objGoldenCushion.PutCustomIV1 = 0;
                        objGoldenCushion.PutCustomIV2 = 0;
                        objGoldenCushion.PutCustomRF1 = 0;
                        objGoldenCushion.PutCustomRF2 = 0;
                        //Added by Shweta on 10th May---------------END---------

                        //-------------Added by Shweta on 22nd July 2016 to Fetch Latest Deployment Rate------------START--------
                        string strDeploymentRate = "";
                        var DeploymentRate = objSP_PRICINGEntities.SP_FETCH_PRICING_DEPLOYMENT_RATE(Convert.ToInt32(objGoldenCushion.RedemptionPeriodDays), objGoldenCushion.EntityID, objGoldenCushion.IsSecuredID);
                        strDeploymentRate = Convert.ToString(DeploymentRate.SingleOrDefault());
                        objGoldenCushion.DeploymentRate = Convert.ToDouble(strDeploymentRate);
                        //-------------Added by Shweta on 22nd July 2016 to Fetch Latest Deployment Rate------------END----------
                    }

                    else if (Session["GoldenCushionChildQuote"] != null)
                    {
                        ViewBag.Message = true;
                        objGoldenCushion = (GoldenCushion)Session["GoldenCushionChildQuote"];
                        objGoldenCushion.UnderlyingID = Convert.ToInt32(Session["UnderlyingID"]);

                        ObjectResult<GoldenCushionEditResult> objGoldenCushionEditResult = objSP_PRICINGEntities.FETCH_GOLDEN_CUSHION_EDIT_DETAILS("");
                        List<GoldenCushionEditResult> GoldenCushionEditResultList = objGoldenCushionEditResult.ToList();
                        GoldenCushion oGoldenCushion = new GoldenCushion();
                        if (GoldenCushionEditResultList != null && GoldenCushionEditResultList.Count > 0)
                            General.ReflectSingleData(oGoldenCushion, GoldenCushionEditResultList[0]);

                        objGoldenCushion.ParentProductID = objGoldenCushion.ProductID;
                        objGoldenCushion.ProductID = "";
                        objGoldenCushion.Status = oGoldenCushion.Status;
                        objGoldenCushion.SaveStatus = oGoldenCushion.SaveStatus;
                        objGoldenCushion.IsChildQuote = true;
                    }
                    else if (Session["CancelQuote"] != null)
                    {
                        objGoldenCushion = (GoldenCushion)Session["CancelQuote"];

                        ObjectResult<GoldenCushionEditResult> objGoldenCushionEditResult = objSP_PRICINGEntities.FETCH_GOLDEN_CUSHION_EDIT_DETAILS(objGoldenCushion.ProductID);
                        List<GoldenCushionEditResult> GoldenCushionEditResultList = objGoldenCushionEditResult.ToList();
                        GoldenCushion oGoldenCushion = new GoldenCushion();
                        if (GoldenCushionEditResultList != null && GoldenCushionEditResultList.Count > 0)
                            General.ReflectSingleData(oGoldenCushion, GoldenCushionEditResultList[0]);

                        objGoldenCushion.Status = oGoldenCushion.Status;
                        objGoldenCushion.SaveStatus = oGoldenCushion.SaveStatus;

                        Session.Remove("CancelQuote");
                    }
                    else
                    {
                        Session.Remove("IsChildQuoteGoldenCushion");
                        Session.Remove("ParentProductID");
                        Session.Remove("UnderlyingID");
                    }

                    if (IsQuotron == true)
                    {
                        objGoldenCushion.IsQuotron = true;
                    }

                    if (Session["GoldenCushionChildQuote"] == null && Session["GoldenCushionCopyQuote"] == null)
                        objGoldenCushion.SaveStatus = "";

                    if (Session["GoldenCushionCopyQuote"] != null)
                        Session.Remove("GoldenCushionCopyQuote");

                    if (Session["GoldenCushionChildQuote"] != null)
                        Session.Remove("GoldenCushionChildQuote");

                    if (ProductID == null)
                    {
                        objGoldenCushion.isGraphActive = false;
                        return View(objGoldenCushion);
                    }
                    else
                    {
                        return GenerateGoldenCushionGraph(objGoldenCushion);
                    }
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

                ClearGoldenCushionSession();

                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "GoldenCushion Get", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        private ActionResult GenerateGoldenCushionGraph(GoldenCushion objGoldenCushion)
        {
            LoginController objLoginController = new LoginController();
            try
            {
                if (ValidateSession())
                {
                    //var ParticipatoryRatio = TruncateDecimal(objGoldenCushion.PutSpreadPaticipatoryRatio, 2);

                    var transactionCounts = new List<Graph>();
                    transactionCounts = GenerateGraphGoldenCushion(objGoldenCushion.PutSpreadStrike1, objGoldenCushion.PutSpreadStrike2, objGoldenCushion.PutStrike1, objGoldenCushion.PutStrike2, objGoldenCushion.LowerCoupon, objGoldenCushion.FixedCoupon, objGoldenCushion.PutSpreadPaticipatoryRatio, objGoldenCushion.PutPaticipatoryRatio, objGoldenCushion.PutSpreadOptionTypeId, objGoldenCushion.PutOptionTypeId);

                    //GoldenCushion obj = new GoldenCushion();

                    #region Pie Chart For FixedOrPR
                    var xDataMonths = transactionCounts.Select(i => i.Column1).ToArray();
                    var yDataCounts = transactionCounts.Select(i => new object[] { i.Column2 }).ToArray();
                    var yDataCounts1 = transactionCounts.Select(i => new object[] { i.Column3 }).ToArray();

                    var GoldenCushionChart = new Highcharts("pie")
                        //define the type of chart 
                                .InitChart(new Chart { DefaultSeriesType = ChartTypes.Line })
                        //overall Title of the chart 
                                .SetTitle(new Title { Text = "Golden Cushion" })
                        ////small label below the main Title
                        //        .SetSubtitle(new Subtitle { Text = "Accounting" })
                        //load the X values
                                .SetXAxis(new XAxis { Title = new XAxisTitle { Text = "Underlying Returns" }, Categories = xDataMonths, Labels = new XAxisLabels { Step = 2 } })
                        //set the Y title
                                .SetYAxis(new YAxis { Title = new YAxisTitle { Text = "Product Returns" } })
                                .SetTooltip(new Tooltip
                                {
                                    Enabled = true,
                                    Formatter = @"function() { return '<b>'+ this.series.name +'</b><br/>'+ this.x +': '+ this.y; }"
                                })
                                .SetPlotOptions(new PlotOptions
                                {
                                    Line = new PlotOptionsLine
                                    {
                                        DataLabels = new PlotOptionsLineDataLabels
                                        {
                                            Enabled = false
                                        },
                                        EnableMouseTracking = true
                                    }
                                })
                        //load the Y values 
                                .SetSeries(new[]
                    {
                        new Series {Name = "Coupon", Data = new Data(yDataCounts)},
                            //you can add more y data to create a second line
                            // new Series { Name = "Strike", Data = new Data(yDataCounts1) }
                    });
                    #endregion

                    if (Session["GoldenCushionCopyQuote"] != null)
                    {
                        ViewBag.Message = "Quote is copied";
                        objGoldenCushion = (GoldenCushion)Session["GoldenCushionCopyQuote"];
                        Session.Remove("GoldenCushionCopyQuote");
                    }

                    objGoldenCushion.GoldenCushionChart = GoldenCushionChart;

                    return View(objGoldenCushion);
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

                ClearGoldenCushionSession();

                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "GenerateGoldenCushionGraph", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        public JsonResult ManageGoldenCushion(string ProductID, string Distributor, string EdelweissBuiltIn, string DistributorBuiltIn, string BuiltInAdjustment, string TotalBuiltIn, string DeploymentRate, string CustomerDeploymentRate,
            string FixedCoupon, string IRR, string IsFixedCouponIRR, string LowerCoupon, string LowerCouponIRR, string IsLowerCouponIRR, string OptionTenureMonth, string RedemptionPeriodMonth, string IsRedemptionPeriodMonth, string RedemptionPeriodDays,
            string InitialAveragingMonth, string InitialAveragingDaysDiff, string FinalAveragingMonth, string FinalAveragingDaysDiff, string IsPrincipalProtected,
            string Underlying, string Remaining, string TotalOptionPrice, string NetRemaining, string SalesComments, string TradingComments, string CouponScenario1,
            string CouponScenario2, string PutSpreadOptionType, string PutSpreadStrike1, string PutSpreadStrike2, string PutSpreadParticipatoryRatio, string PutSpreadPrice,
            string PutSpreadDiscountedPrice, string PutSpreadPRAdjustedPrice, string PutSpreadIV1, string PutSpreadCustomIV1, string PutSpreadRF1, string PutSpreadCustomRF1,
            string PutSpreadIV2, string PutSpreadCustomIV2, string PutSpreadRF2, string PutSpreadCustomRF2, string PutOptionType, string PutStrike1, string PutStrike2,
            string PutParticipatoryRatio, string PutPrice, string PutDiscountedPrice, string PutPRAdjustedPrice, string PutIV1, string PutCustomIV1, string PutRF1,
            string PutCustomRF1, string PutIV2, string PutCustomIV2, string PutRF2, string PutCustomRF2, string CopyProductID,
            string PutSpreadStrike1Summary, string PutSpreadStrike2Summary, string PutStrike1Summary, string PutStrike2Summary,
            string ExportPutSpreadStrike1Summary, string ExportPutSpreadStrike2Summary, string ExportPutStrike1Summary, string ExportPutStrike2Summary, string Entity, string IsSecured)
        {
            try
            {
                ExportPutSpreadStrike1Summary = System.Uri.UnescapeDataString(ExportPutSpreadStrike1Summary);
                ExportPutSpreadStrike2Summary = System.Uri.UnescapeDataString(ExportPutSpreadStrike2Summary);
                ExportPutStrike1Summary = System.Uri.UnescapeDataString(ExportPutStrike1Summary);
                ExportPutStrike2Summary = System.Uri.UnescapeDataString(ExportPutStrike2Summary);

                PutSpreadStrike1Summary = System.Uri.UnescapeDataString(PutSpreadStrike1Summary);
                PutSpreadStrike2Summary = System.Uri.UnescapeDataString(PutSpreadStrike2Summary);
                PutStrike1Summary = System.Uri.UnescapeDataString(PutStrike1Summary);
                PutStrike2Summary = System.Uri.UnescapeDataString(PutStrike2Summary);

                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];

                if (CustomerDeploymentRate == "")
                    CustomerDeploymentRate = "0";

                if (TotalOptionPrice == "" || TotalOptionPrice == "NaN")
                    TotalOptionPrice = "0";

                if (NetRemaining == "" || NetRemaining == "NaN")
                    NetRemaining = "0";

                if (PutSpreadStrike1 == "")
                    PutSpreadStrike1 = "0";

                if (PutSpreadParticipatoryRatio == "" || PutSpreadParticipatoryRatio == "NaN")
                    PutSpreadParticipatoryRatio = "0";

                if (PutSpreadPrice == "" || PutSpreadPrice == "NaN")
                    PutSpreadPrice = "0";

                if (PutSpreadDiscountedPrice == "" || PutSpreadDiscountedPrice == "NaN")
                    PutSpreadDiscountedPrice = "0";

                if (PutSpreadPRAdjustedPrice == "" || PutSpreadPRAdjustedPrice == "NaN")
                    PutSpreadPRAdjustedPrice = "0";

                if (PutSpreadIV1 == "" || PutSpreadIV1 == "NaN")
                    PutSpreadIV1 = "0";

                if (PutSpreadCustomIV1 == "" || PutSpreadCustomIV1 == "NaN")
                    PutSpreadCustomIV1 = "0";

                if (PutSpreadRF1 == "" || PutSpreadRF1 == "NaN")
                    PutSpreadRF1 = "0";

                if (PutSpreadCustomRF1 == "" || PutSpreadCustomRF1 == "NaN")
                    PutSpreadCustomRF1 = "0";

                if (PutSpreadIV2 == "" || PutSpreadIV2 == "NaN")
                    PutSpreadIV2 = "0";

                if (PutSpreadCustomIV2 == "" || PutSpreadCustomIV2 == "NaN")
                    PutSpreadCustomIV2 = "0";

                if (PutSpreadRF2 == "" || PutSpreadRF2 == "NaN")
                    PutSpreadRF2 = "0";

                if (PutSpreadCustomRF2 == "" || PutSpreadCustomRF2 == "NaN")
                    PutSpreadCustomRF2 = "0";

                if (PutOptionType == null)
                    PutOptionType = "";

                if (PutStrike1 == "" || PutStrike1 == "NaN")
                    PutStrike1 = "0";

                if (PutStrike2 == "" || PutStrike2 == "NaN")
                    PutStrike2 = "0";

                if (PutParticipatoryRatio == "" || PutParticipatoryRatio == null || PutParticipatoryRatio == "NaN")
                    PutParticipatoryRatio = "0";

                if (PutPrice == "" || PutPrice == "NaN")
                    PutPrice = "0";

                if (PutDiscountedPrice == "" || PutDiscountedPrice == "NaN")
                    PutDiscountedPrice = "0";

                if (PutPRAdjustedPrice == "" || PutPRAdjustedPrice == "NaN")
                    PutPRAdjustedPrice = "0";

                if (PutIV1 == "")
                    PutIV1 = "0";

                if (PutCustomIV1 == "")
                    PutCustomIV1 = "0";

                if (PutRF1 == "")
                    PutRF1 = "0";

                if (PutCustomRF1 == "")
                    PutCustomRF1 = "0";

                if (PutIV2 == "")
                    PutIV2 = "0";

                if (PutCustomIV2 == "")
                    PutCustomIV2 = "0";

                if (PutRF2 == "")
                    PutRF2 = "0";

                if (LowerCouponIRR == "")
                    LowerCouponIRR = "0";

                if (PutCustomRF2 == "")
                    PutCustomRF2 = "0";

                string ParentProductID = "";
                if (Session["ParentProductID"] != null)
                    ParentProductID = (string)Session["ParentProductID"];

                ObjectResult<ManageGoldenCushionResult> objManageGoldenCushionResult = objSP_PRICINGEntities.SP_MANAGE_GOLDEN_CUSHION_DETAILS(ProductID, ParentProductID, Distributor, Convert.ToDouble(EdelweissBuiltIn),
                        Convert.ToDouble(DistributorBuiltIn), Convert.ToDouble(BuiltInAdjustment), Convert.ToDouble(TotalBuiltIn), Convert.ToDouble(DeploymentRate), Convert.ToDouble(CustomerDeploymentRate), Convert.ToDouble(FixedCoupon), Convert.ToDouble(IRR), Convert.ToBoolean(IsFixedCouponIRR),
                        Convert.ToDouble(LowerCoupon), Convert.ToDouble(LowerCouponIRR), Convert.ToBoolean(IsLowerCouponIRR), Convert.ToInt32(OptionTenureMonth), Convert.ToDouble(RedemptionPeriodMonth), Convert.ToBoolean(IsRedemptionPeriodMonth), Convert.ToInt32(RedemptionPeriodDays),
                        Convert.ToInt32(InitialAveragingMonth), Convert.ToInt32(InitialAveragingDaysDiff), Convert.ToInt32(FinalAveragingMonth), Convert.ToInt32(FinalAveragingDaysDiff),
                        Convert.ToBoolean(IsPrincipalProtected), Convert.ToInt32(Underlying), Convert.ToDouble(Remaining), Convert.ToDouble(TotalOptionPrice), Convert.ToDouble(NetRemaining),
                        SalesComments, TradingComments, CouponScenario1, CouponScenario2, Convert.ToInt32(Entity), Convert.ToInt32(IsSecured), objUserMaster.UserID, PutSpreadOptionType, Convert.ToDouble(PutSpreadStrike1), Convert.ToDouble(PutSpreadStrike2), Convert.ToDouble(PutSpreadParticipatoryRatio), Convert.ToDouble(PutSpreadPrice),
                        Convert.ToDouble(PutSpreadDiscountedPrice), Convert.ToDouble(PutSpreadPRAdjustedPrice), Convert.ToDouble(PutSpreadIV1), Convert.ToDouble(PutSpreadCustomIV1),
                        Convert.ToDouble(PutSpreadRF1), Convert.ToDouble(PutSpreadCustomRF1), Convert.ToDouble(PutSpreadIV2), Convert.ToDouble(PutSpreadCustomIV2),
                        Convert.ToDouble(PutSpreadRF2), Convert.ToDouble(PutSpreadCustomRF2), PutOptionType, Convert.ToDouble(PutStrike1), Convert.ToDouble(PutStrike2),
                        Convert.ToDouble(PutParticipatoryRatio), Convert.ToDouble(PutPrice), Convert.ToDouble(PutDiscountedPrice), Convert.ToDouble(PutPRAdjustedPrice),
                        Convert.ToDouble(PutIV1), Convert.ToDouble(PutCustomIV1), Convert.ToDouble(PutRF1), Convert.ToDouble(PutCustomRF1), Convert.ToDouble(PutIV2),
                        Convert.ToDouble(PutCustomIV2), Convert.ToDouble(PutRF2), Convert.ToDouble(PutCustomRF2), CopyProductID,
                        PutSpreadStrike1Summary, PutSpreadStrike2Summary, PutStrike1Summary, PutStrike2Summary,
                        ExportPutSpreadStrike1Summary, ExportPutSpreadStrike2Summary, ExportPutStrike1Summary, ExportPutStrike2Summary);
                List<ManageGoldenCushionResult> ManageGoldenCushionResultList = objManageGoldenCushionResult.ToList();

                Session.Remove("ParentProductID");

                return Json(ManageGoldenCushionResultList[0].ProductID);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];

                ClearGoldenCushionSession();

                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "ManageGoldenCushion", objUserMaster.UserID);
                return Json("");
            }
        }

        [HttpPost, ValidateInput(false)]
        public ActionResult GoldenCushion(string Command, GoldenCushion objGoldenCushion, FormCollection objFormCollection)
        {
            LoginController objLoginController = new LoginController();
            try
            {
                if (ValidateSession())
                {
                    #region Bind Underlying List
                    List<Underlying> UnderlyingList = new List<Underlying>();

                    ObjectResult<UnderlyingListResult> objUnderlyingListResult;
                    List<UnderlyingListResult> UnderlyingListResultList = new List<UnderlyingListResult>();

                    objUnderlyingListResult = objSP_PRICINGEntities.SP_FETCH_UNDERLYING_DETAILS();
                    UnderlyingListResultList = objUnderlyingListResult.ToList();

                    foreach (UnderlyingListResult oUnderlyingListResult in UnderlyingListResultList)
                    {
                        Underlying objUnderlying = new Underlying();
                        General.ReflectSingleData(objUnderlying, oUnderlyingListResult);

                        UnderlyingList.Add(objUnderlying);
                    }
                    objGoldenCushion.UnderlyingList = UnderlyingList;
                    #endregion

                    GoldenCushion oGoldenCushion = new GoldenCushion();
                    if (objGoldenCushion.ProductID != "" && objGoldenCushion.ProductID != null)
                    {
                        ObjectResult<GoldenCushionEditResult> objGoldenCushionEditResult = objSP_PRICINGEntities.FETCH_GOLDEN_CUSHION_EDIT_DETAILS(objGoldenCushion.ProductID);
                        List<GoldenCushionEditResult> GoldenCushionEditResultList = objGoldenCushionEditResult.ToList();

                        General.ReflectSingleData(oGoldenCushion, GoldenCushionEditResultList[0]);
                        objGoldenCushion.IsPrincipalProtected = oGoldenCushion.IsPrincipalProtected;

                        objGoldenCushion.PutSpreadOptionTypeId = oGoldenCushion.PutSpreadOptionTypeId;
                        objGoldenCushion.PutOptionTypeId = oGoldenCushion.PutOptionTypeId;
                        objGoldenCushion.PutStrike2 = oGoldenCushion.PutStrike2;
                    }

                    if (Command == "ExportToExcel")
                    {
                        ExportGoldenCushion(objGoldenCushion, objFormCollection);

                        return RedirectToAction("GoldenCushion");
                    }
                    else if (Command == "ExportPutSpreadStrike1Grid")
                    {
                        string StrikeHTML = objFormCollection["ExportPutSpreadStrike1Summary"];

                        ExportStrikeGrid(StrikeHTML);

                        return RedirectToAction("GoldenCushion");
                    }
                    else if (Command == "ExportPutSpreadStrike2Grid")
                    {
                        string StrikeHTML = objFormCollection["ExportPutSpreadStrike2Summary"];

                        ExportStrikeGrid(StrikeHTML);

                        return RedirectToAction("GoldenCushion");
                    }
                    else if (Command == "ExportPutStrike1Grid")
                    {
                        string StrikeHTML = objFormCollection["ExportPutStrike1Summary"];

                        ExportStrikeGrid(StrikeHTML);

                        return RedirectToAction("GoldenCushion");
                    }
                    else if (Command == "ExportPutStrike2Grid")
                    {
                        string StrikeHTML = objFormCollection["ExportPutStrike2Summary"];

                        ExportStrikeGrid(StrikeHTML);

                        return RedirectToAction("GoldenCushion");
                    }
                    else if (Command == "CopyQuote")
                    {
                        Session["GoldenCushionCopyQuote"] = objGoldenCushion;
                        Session["UnderlyingID"] = objGoldenCushion.UnderlyingID;

                        return RedirectToAction("GoldenCushion");
                    }
                    else if (Command == "CreateChildQuote")
                    {
                        Session.Remove("ParentProductID");
                        Session.Remove("IsChildQuoteGoldenCushion");
                        Session.Remove("UnderlyingID");

                        Session["ParentProductID"] = objGoldenCushion.ProductID;
                        Session["UnderlyingID"] = objGoldenCushion.UnderlyingID;

                        objGoldenCushion.IsChildQuote = true;

                        Session["GoldenCushionChildQuote"] = objGoldenCushion;
                        Session["IsChildQuoteGoldenCushion"] = objGoldenCushion.IsChildQuote;

                        return RedirectToAction("GoldenCushion");
                    }
                    else if (Command == "GenerateGraph")
                    {
                        objGoldenCushion.isGraphActive = true;

                        TempData["GoldenCushionGraph"] = objGoldenCushion;
                        return RedirectToAction("GoldenCushion", new { GenerateGraph = "GenerateGraph" });
                        // return GenerateGoldenCushionGraph(objGoldenCushion);
                    }
                    else if (Command == "AddNewProduct")
                    {
                        var productID = objGoldenCushion.ProductID;
                        UserMaster objUserMaster = new UserMaster();
                        objUserMaster = (UserMaster)Session["LoggedInUser"];

                        EncryptDecrypt obj = new EncryptDecrypt();
                        var encryptedpaswd = obj.Encrypt(objUserMaster.Password, "SPPricing", CryptographyEngine.AlgorithmType.DES);

                        var isPrincipalProtected = objFormCollection.Get("PrincipalProtected");
                        var ProductType = "PP";
                        if (isPrincipalProtected != "1")
                        {
                            ProductType = "NonPP";
                        }

                        var Url = "http://edemumnewuatvm4:63400/Login.aspx?UserId=" + objUserMaster.LoginName + "&Key=" + encryptedpaswd + "&ProductId=" + productID + "&ProductType=" + ProductType;
                        return Redirect(Url);
                    }
                    else if (Command == "Cancel")
                    {
                        Session["CancelQuote"] = objGoldenCushion;

                        return RedirectToAction("GoldenCushion");
                    }
                    else if (Command == "PricingInExcel")
                    {
                        objGoldenCushion.IsWorkingFileExport = OpenWorkingExcelFile("GC", objGoldenCushion.ProductID);

                        if (!objGoldenCushion.IsWorkingFileExport)
                            objGoldenCushion.WorkingFileStatus = "File Not Found";

                        return View(objGoldenCushion);
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

                ClearGoldenCushionSession();

                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "GoldenCushion Post", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        public virtual void ExportGoldenCushion(GoldenCushion objGoldenCushion, FormCollection objFormCollection)
        {
            try
            {
                string strTemplateFilePath = Server.MapPath("~/Templates");
                string strTemplateFileName = strTemplateFilePath + "//GoldenCushionTemplate.xlsx";

                string strTargetFilePath = Server.MapPath("~/OutputFiles");
                string strTargetFileName = strTargetFilePath + "//" + objGoldenCushion.ProductID + "_GoldenCushion.xlsx";

                string Role = Convert.ToString(Session["Role"]);

                if (System.IO.File.Exists(strTargetFileName))
                    System.IO.File.Delete(strTargetFileName);

                FileInfo objTemplateFileInfo = new FileInfo(strTemplateFileName);
                objTemplateFileInfo.CopyTo(strTargetFileName);

                FileInfo objTargetFileInfo = new FileInfo(strTargetFileName);

                Underlying objUnderlying = objGoldenCushion.UnderlyingList.Find(delegate(Underlying oUnderlying) { return oUnderlying.UnderlyingID == objGoldenCushion.UnderlyingID; });

                using (var xlPackage = new ExcelPackage(objTargetFileInfo))
                {
                    var worksheet = xlPackage.Workbook.Worksheets["GoldenCushion"];

                    worksheet.Cell(1, 2).Value = objGoldenCushion.ProductID.ToString();
                    worksheet.Cell(1, 4).Value = objGoldenCushion.Distributor.ToString().ToUpper();
                    worksheet.Cell(1, 6).Value = objUnderlying.UnderlyingShortName;
                    if (objGoldenCushion.IsPrincipalProtected)
                        worksheet.Cell(1, 8).Value = "Yes";
                    else
                        worksheet.Cell(1, 8).Value = "No";

                    worksheet.Cell(2, 2).Formula = "=" + objGoldenCushion.EdelweissBuiltIn.ToString() + "%";
                    worksheet.Cell(2, 4).Formula = "=" + objGoldenCushion.DistributorBuiltIn.ToString() + "%";
                    worksheet.Cell(2, 6).Formula = "=B2+D2+H2";
                    worksheet.Cell(2, 8).Formula = "=" + objGoldenCushion.BuiltInAdjustment.ToString() + "%";

                    worksheet.Cell(3, 2).Formula = "=((POWER((1+D3),(12/D4))-1)*100) %";
                    worksheet.Cell(3, 4).Formula = "=" + objGoldenCushion.FixedCoupon.ToString() + "%";
                    worksheet.Cell(3, 6).Formula = "=((POWER((1+H3),(12/D4))-1)*100) %";
                    worksheet.Cell(3, 8).Formula = "=" + objGoldenCushion.LowerCoupon.ToString() + "%";

                    worksheet.Cell(4, 2).Value = objGoldenCushion.OptionTenure.ToString();
                    worksheet.Cell(4, 4).Formula = "=ROUND(F4/30.417,0)";
                    worksheet.Cell(4, 6).Value = objGoldenCushion.RedemptionPeriodDays.ToString();

                    #region Get Entity Name
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
                    LookupMaster objLookupMasterEntity = EntityList.Find(delegate(LookupMaster oLookupMaster) { return oLookupMaster.LookupID == objGoldenCushion.EntityID; });
                    worksheet.Cell(5, 2).Value = objLookupMasterEntity.LookupDescription;
                    #endregion

                    #region Get Is Secured
                    objLookupResult = null;
                    LookupResultList = null;
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
                    LookupMaster objLookupMasterIsSecured = IsSecuredList.Find(delegate(LookupMaster oLookupMaster) { return oLookupMaster.LookupID == objGoldenCushion.IsSecuredID; });
                    worksheet.Cell(5, 4).Value = objLookupMasterIsSecured.LookupDescription;
                    #endregion
                    worksheet.Cell(5, 6).Formula = "=" + objGoldenCushion.DeploymentRate.ToString() + "%";
                    worksheet.Cell(5, 8).Formula = "=" + objGoldenCushion.CustomDeploymentRate.ToString() + "%";

                    worksheet.Cell(6, 2).Formula = "=(100-(B2+D2)*100)-(100*(1+ROUND(D3,4)))/(POWER((1+(IF(H5>0,H5,F5))),(D4/12)))";
                    worksheet.Cell(6, 4).Formula = "=H11 + H12";
                    worksheet.Cell(6, 6).Formula = "=D6 + B6";

                    worksheet.Cell(8, 2).Value = objGoldenCushion.InitialAveragingMonth.ToString();
                    worksheet.Cell(8, 4).Value = objGoldenCushion.InitialAveragingDaysDiff.ToString();
                    worksheet.Cell(8, 6).Value = objGoldenCushion.FinalAveragingMonth.ToString();
                    worksheet.Cell(8, 8).Value = objGoldenCushion.FinalAveragingDaysDiff.ToString();

                    if (objGoldenCushion.PutSpreadOptionTypeId != null)
                        worksheet.Cell(11, 1).Value = objGoldenCushion.PutSpreadOptionTypeId.ToString();
                    else
                        worksheet.Cell(11, 1).Value = "Put Spread Short";
                    worksheet.Cell(11, 2).Value = objUnderlying.UnderlyingShortName;
                    worksheet.Cell(11, 3).Value = objGoldenCushion.PutSpreadStrike1.ToString();
                    worksheet.Cell(11, 4).Value = objGoldenCushion.PutSpreadStrike2.ToString();
                    worksheet.Cell(11, 5).Value = objGoldenCushion.PutSpreadPaticipatoryRatio.ToString();
                    worksheet.Cell(11, 6).Value = objGoldenCushion.PutSpreadPrice.ToString();
                    worksheet.Cell(11, 7).Value = objGoldenCushion.PutSpreadDiscountedPrice.ToString();
                    worksheet.Cell(11, 8).Value = objGoldenCushion.PutSpreadPrAdjustmentPrice.ToString();


                    if (Role == "Sales")
                    {
                        worksheet.Cell(10, 9).Value = "";
                        worksheet.Cell(10, 10).Value = "";
                        worksheet.Cell(10, 11).Value = "";
                        worksheet.Cell(10, 12).Value = "";
                        worksheet.Cell(10, 13).Value = "";
                        worksheet.Cell(10, 14).Value = "";
                        worksheet.Cell(10, 15).Value = "";
                        worksheet.Cell(10, 16).Value = "";
                    }
                    else
                    {
                        worksheet.Cell(11, 9).Formula = "=" + objGoldenCushion.PutSpreadIV1.ToString() + "%";
                        worksheet.Cell(11, 10).Formula = "=" + objGoldenCushion.PutSpreadCustomIV1.ToString() + "%";
                        worksheet.Cell(11, 11).Formula = "=" + objGoldenCushion.PutSpreadRF1.ToString() + "%";
                        worksheet.Cell(11, 12).Formula = "=" + objGoldenCushion.PutSpreadCustomRF1.ToString() + "%";
                        worksheet.Cell(11, 13).Formula = "=" + objGoldenCushion.PutSpreadIV2.ToString() + "%";
                        worksheet.Cell(11, 14).Formula = "=" + objGoldenCushion.PutSpreadCustomIV2.ToString() + "%";
                        worksheet.Cell(11, 15).Formula = "=" + objGoldenCushion.PutSpreadRF2.ToString() + "%";
                        worksheet.Cell(11, 16).Formula = "=" + objGoldenCushion.PutSpreadCustomRF2.ToString() + "%";
                    }

                    if (!objGoldenCushion.IsPrincipalProtected)
                        if (objGoldenCushion.PutStrike1 != 0)
                        {
                            if (objGoldenCushion.PutOptionTypeId != null)
                                worksheet.Cell(12, 1).Value = objGoldenCushion.PutOptionTypeId.ToString();
                            else
                                worksheet.Cell(12, 1).Value = "Put Short";
                            worksheet.Cell(12, 2).Value = objUnderlying.UnderlyingShortName;
                            worksheet.Cell(12, 3).Value = objGoldenCushion.PutStrike1.ToString();
                            worksheet.Cell(12, 4).Value = objGoldenCushion.PutStrike2.ToString();
                            worksheet.Cell(12, 5).Value = objGoldenCushion.PutPaticipatoryRatio.ToString();
                            worksheet.Cell(12, 6).Value = objGoldenCushion.PutPrice.ToString();
                            worksheet.Cell(12, 7).Value = objGoldenCushion.PutDiscountedPrice.ToString();
                            worksheet.Cell(12, 8).Value = objGoldenCushion.PutPrAdjustmentPrice.ToString();

                            if (Role == "Sales")
                            {
                                worksheet.Cell(10, 9).Value = "";
                                worksheet.Cell(10, 10).Value = "";
                                worksheet.Cell(10, 11).Value = "";
                                worksheet.Cell(10, 12).Value = "";
                                worksheet.Cell(10, 13).Value = "";
                                worksheet.Cell(10, 14).Value = "";
                                worksheet.Cell(10, 15).Value = "";
                                worksheet.Cell(10, 16).Value = "";
                            }
                            else
                            {
                                worksheet.Cell(12, 9).Formula = "=" + objGoldenCushion.PutIV1.ToString() + "%";
                                worksheet.Cell(12, 10).Formula = "=" + objGoldenCushion.PutCustomIV1.ToString() + "%";
                                worksheet.Cell(12, 11).Formula = "=" + objGoldenCushion.PutRF1.ToString() + "%";
                                worksheet.Cell(12, 12).Formula = "=" + objGoldenCushion.PutCustomRF1.ToString() + "%";
                                worksheet.Cell(12, 13).Formula = "=" + objGoldenCushion.PutIV2.ToString() + "%";
                                worksheet.Cell(12, 14).Formula = "=" + objGoldenCushion.PutCustomIV2.ToString() + "%";
                                worksheet.Cell(12, 15).Formula = "=" + objGoldenCushion.PutRF2.ToString() + "%";
                                worksheet.Cell(12, 16).Formula = "=" + objGoldenCushion.PutCustomRF2.ToString() + "%";
                            }
                        }

                    if (objGoldenCushion.SalesComments != null)
                        worksheet.Cell(14, 2).Value = objGoldenCushion.SalesComments.ToString();
                    else
                        worksheet.Cell(14, 2).Value = "";

                    if (objGoldenCushion.TradingComments != null)
                        worksheet.Cell(15, 2).Value = objGoldenCushion.TradingComments.ToString();
                    else
                        worksheet.Cell(15, 2).Value = "";

                    if (objGoldenCushion.CouponScenario1 != null)
                        worksheet.Cell(16, 2).Value = objGoldenCushion.CouponScenario1.ToString();
                    else
                        worksheet.Cell(16, 2).Value = "";

                    if (objGoldenCushion.CouponScenario2 != null)
                        worksheet.Cell(17, 2).Value = objGoldenCushion.CouponScenario2.ToString();
                    else
                        worksheet.Cell(17, 2).Value = "";

                    xlPackage.Save();
                }

                if (System.IO.File.Exists(strTargetFileName))
                {
                    FileInfo TemplateFile = new FileInfo(strTargetFileName);

                    Response.Clear();
                    Response.ClearHeaders();
                    Response.ClearContent();
                    Response.AddHeader("content-disposition", "attachment; filename=" + Path.GetFileName(strTargetFileName));
                    Response.AddHeader("Content-Type", "application/Excel");
                    Response.ContentType = "application/vnd.xls";
                    Response.AddHeader("Content-Length", TemplateFile.Length.ToString());
                    Response.WriteFile(TemplateFile.FullName);
                    Response.End();
                }
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];

                ClearGoldenCushionSession();

                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "ExportGoldenCushion", objUserMaster.UserID);
            }
        }

        public JsonResult ExportGoldenCushionWorkingFile(string ProductID, string Distributor, string EdelweissBuiltIn, string DistributorBuiltIn, string BuiltInAdjustment, string TotalBuiltIn, string DeploymentRate, string CustomerDeploymentRate,
            string FixedCoupon, string IRR, string IsFixedCouponIRR, string LowerCoupon, string LowerCouponIRR, string IsLowerCouponIRR, string OptionTenureMonth, string RedemptionPeriodMonth, string IsRedemptionPeriodMonth, string RedemptionPeriodDays,
            string InitialAveragingMonth, string InitialAveragingDaysDiff, string FinalAveragingMonth, string FinalAveragingDaysDiff, string IsPrincipalProtected,
            string Underlying, string Remaining, string TotalOptionPrice, string NetRemaining, string SalesComments, string TradingComments, string CouponScenario1,
            string CouponScenario2, string PutSpreadOptionType, string PutSpreadStrike1, string PutSpreadStrike2, string PutSpreadParticipatoryRatio, string PutSpreadPrice,
            string PutSpreadDiscountedPrice, string PutSpreadPRAdjustedPrice, string PutSpreadIV1, string PutSpreadCustomIV1, string PutSpreadRF1, string PutSpreadCustomRF1,
            string PutSpreadIV2, string PutSpreadCustomIV2, string PutSpreadRF2, string PutSpreadCustomRF2, string PutOptionType, string PutStrike1, string PutStrike2,
            string PutParticipatoryRatio, string PutPrice, string PutDiscountedPrice, string PutPRAdjustedPrice, string PutIV1, string PutCustomIV1, string PutRF1,
            string PutCustomRF1, string PutIV2, string PutCustomIV2, string PutRF2, string PutCustomRF2,
            string PutSpreadStrike1Summary, string PutSpreadStrike2Summary, string PutStrike1Summary, string PutStrike2Summary,
            string ExportPutSpreadStrike1Summary, string ExportPutSpreadStrike2Summary, string ExportPutStrike1Summary, string ExportPutStrike2Summary, string Entity, string IsSecured)
        {
            try
            {
                string strTemplateFilePath = Server.MapPath("~/Templates");
                string strTemplateFileName = strTemplateFilePath + "//GoldenCushionTemplateWorkingFile.xlsx";

                string strTargetFilePath = System.Configuration.ConfigurationManager.AppSettings["WorkingFilePath"];
                string strTargetFileName = strTargetFilePath + "//" + ProductID + "_GoldenCushion.xlsx";

                string Role = Convert.ToString(Session["Role"]);

                if (System.IO.File.Exists(strTargetFileName))
                    System.IO.File.Delete(strTargetFileName);

                FileInfo objTemplateFileInfo = new FileInfo(strTemplateFileName);
                objTemplateFileInfo.CopyTo(strTargetFileName);

                FileInfo objTargetFileInfo = new FileInfo(strTargetFileName);

                using (var xlPackage = new ExcelPackage(objTargetFileInfo))
                {
                    var worksheet = xlPackage.Workbook.Worksheets["GoldenCushion"];

                    worksheet.Cell(1, 2).Value = ProductID.ToString();
                    worksheet.Cell(1, 4).Value = Distributor.ToString().ToUpper();
                    worksheet.Cell(1, 6).Value = Underlying;
                    if (IsPrincipalProtected.ToUpper() == "TRUE")
                        worksheet.Cell(1, 8).Value = "Yes";
                    else
                        worksheet.Cell(1, 8).Value = "No";

                    worksheet.Cell(2, 2).Formula = "=" + EdelweissBuiltIn.ToString() + "%";
                    worksheet.Cell(2, 4).Formula = "=" + DistributorBuiltIn.ToString() + "%";
                    worksheet.Cell(2, 6).Formula = "=B2+D2+H2";
                    worksheet.Cell(2, 8).Formula = "=" + BuiltInAdjustment.ToString() + "%";

                    if (IsFixedCouponIRR.ToUpper() == "TRUE")
                    {
                        worksheet.Cell(3, 2).Formula = "=" + IRR.ToString() + "%";
                        worksheet.Cell(3, 4).Formula = "=(POWER((1 + B3), (F4 / 365)) - 1)*100%";
                    }
                    else
                    {
                        worksheet.Cell(3, 2).Formula = "=((POWER((1+D3),(12/D4))-1) * 100) %";
                        worksheet.Cell(3, 4).Formula = "=" + FixedCoupon.ToString() + "%";
                    }

                    if (IsLowerCouponIRR.ToUpper() == "TRUE")
                    {
                        worksheet.Cell(3, 6).Formula = "=" + LowerCouponIRR.ToString() + "%";
                        worksheet.Cell(3, 8).Formula = "=(POWER((1 + F3), (F4 / 365)) - 1)*100%";
                    }
                    else
                    {
                        worksheet.Cell(3, 6).Formula = "=((POWER((1+H3),(12/D4))-1) * 100) %";
                        worksheet.Cell(3, 8).Formula = "=" + LowerCoupon.ToString() + "%";
                    }

                    worksheet.Cell(4, 2).Value = OptionTenureMonth.ToString();

                    if (IsRedemptionPeriodMonth.ToUpper() == "TRUE")
                    {
                        worksheet.Cell(4, 4).Formula = RedemptionPeriodMonth;
                        worksheet.Cell(4, 6).Formula = "=ROUND(D4*30.417, 0)";
                    }
                    else
                    {
                        worksheet.Cell(4, 4).Formula = "=ROUND(F4/30.417,2)";
                        worksheet.Cell(4, 6).Formula = RedemptionPeriodDays.ToString();
                    }

                    #region Get Entity Name
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
                    LookupMaster objLookupMasterEntity = EntityList.Find(delegate(LookupMaster oLookupMaster) { return oLookupMaster.LookupID == Convert.ToInt32(Entity); });
                    worksheet.Cell(5, 2).Value = objLookupMasterEntity.LookupDescription;
                    #endregion

                    #region Get Is Secured
                    objLookupResult = null;
                    LookupResultList = null;
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
                    LookupMaster objLookupMasterIsSecured = IsSecuredList.Find(delegate(LookupMaster oLookupMaster) { return oLookupMaster.LookupID == Convert.ToInt32(IsSecured); });
                    worksheet.Cell(5, 4).Value = objLookupMasterIsSecured.LookupDescription;

                    #endregion
                    worksheet.Cell(5, 6).Formula = "=" + DeploymentRate.ToString() + "%";

                    if (CustomerDeploymentRate == "")
                        CustomerDeploymentRate = "0";
                    worksheet.Cell(5, 8).Formula = "=" + CustomerDeploymentRate.ToString() + "%";

                    worksheet.Cell(6, 2).Formula = "=(100-(B2+D2)*100)-(100*(1+ROUND(D3,4)))/(POWER((1+(IF(H5>0,H5,F5))),(D4/12)))";
                    worksheet.Cell(6, 4).Formula = "=H11 + H12";
                    worksheet.Cell(6, 6).Formula = "=D6 + B6";

                    worksheet.Cell(8, 2).Value = InitialAveragingMonth.ToString();
                    worksheet.Cell(8, 4).Value = InitialAveragingDaysDiff.ToString();
                    worksheet.Cell(8, 6).Value = FinalAveragingMonth.ToString();
                    worksheet.Cell(8, 8).Value = FinalAveragingDaysDiff.ToString();

                    if (PutSpreadOptionType != null)
                        worksheet.Cell(11, 1).Value = PutSpreadOptionType.ToString();
                    else
                        worksheet.Cell(11, 1).Value = "Put Spread Short";

                    worksheet.Cell(11, 2).Value = Underlying;
                    worksheet.Cell(11, 3).Value = PutSpreadStrike1.ToString();
                    worksheet.Cell(11, 4).Value = PutSpreadStrike2.ToString();
                    worksheet.Cell(11, 5).Formula = "=(D3 - H3) * 100 / (C11 - D11)";
                    worksheet.Cell(11, 6).Formula = "=(AVERAGE(INDIRECT(\"$S$2:\"&ADDRESS(1+$B$8,18+$F$8))))-(AVERAGE(INDIRECT(\"$AA$2:\"&ADDRESS(1+$B$8,26+$F$8))))";
                    worksheet.Cell(11, 7).Formula = "=F11 * (1 / POWER((1 + ((1 + IF(H5>0,H5,F5)) / (1 + K11) - 1)), (INT(D4) / 12)))";
                    worksheet.Cell(11, 8).Formula = "=ROUND(G11,4)*ROUND(E11,4)";

                    worksheet.Cell(11, 9).Formula = "=" + PutSpreadIV1.ToString() + "%";

                    if (PutSpreadCustomIV1 == "")
                        PutSpreadCustomIV1 = "0";
                    worksheet.Cell(11, 10).Formula = "=" + PutSpreadCustomIV1.ToString() + "%";

                    worksheet.Cell(11, 11).Formula = "=" + PutSpreadRF1.ToString() + "%";

                    if (PutSpreadCustomRF1 == "")
                        PutSpreadCustomRF1 = "0";
                    worksheet.Cell(11, 12).Formula = "=" + PutSpreadCustomRF1.ToString() + "%";

                    worksheet.Cell(11, 13).Formula = "=" + PutSpreadIV2.ToString() + "%";

                    if (PutSpreadCustomIV2 == "")
                        PutSpreadCustomIV2 = "0";
                    worksheet.Cell(11, 14).Formula = "=" + PutSpreadCustomIV2.ToString() + "%";

                    worksheet.Cell(11, 15).Formula = "=" + PutSpreadRF2.ToString() + "%";

                    if (PutSpreadCustomRF2 == "")
                        PutSpreadCustomRF2 = "0";
                    worksheet.Cell(11, 16).Formula = "=" + PutSpreadCustomRF2.ToString() + "%";

                    if (IsPrincipalProtected.ToUpper() == "FALSE")
                    {
                        if (Convert.ToInt32(PutStrike1) != 0)
                        {
                            if (PutOptionType != null)
                                worksheet.Cell(12, 1).Value = PutOptionType;
                            else
                                worksheet.Cell(12, 1).Value = "Put Short";
                            worksheet.Cell(12, 2).Value = Underlying;
                            worksheet.Cell(12, 3).Value = PutStrike1.ToString();
                            worksheet.Cell(12, 4).Value = PutStrike2.ToString();
                            worksheet.Cell(12, 5).Value = PutParticipatoryRatio;

                            if (PutStrike2 != "" && Convert.ToInt32(PutStrike2) != 0)
                                worksheet.Cell(12, 6).Formula = "=(AVERAGE(INDIRECT(\"$S$10:\"&ADDRESS(9+B8,18+F8))))-(AVERAGE(INDIRECT(\"$AA$10:\"&ADDRESS(9+$B$8,26+$F$8))))";
                            else
                                worksheet.Cell(12, 6).Formula = "=AVERAGE(INDIRECT(\"$S$10:\"&ADDRESS(9+B8,18+F8)))";

                            worksheet.Cell(12, 7).Formula = "=F12 * (1 / POWER((1 + ((1 + IF(H5>0,H5,F5)) / (1 + K12) - 1)), (INT(D4) / 12)))";
                            worksheet.Cell(12, 8).Formula = "=ROUND(G12,4)*ROUND(E12,4)";

                            if (PutIV1 == "")
                                PutIV1 = "0";
                            worksheet.Cell(12, 9).Formula = "=" + PutIV1.ToString() + "%";

                            if (PutCustomIV1 == "")
                                PutCustomIV1 = "0";
                            worksheet.Cell(12, 10).Formula = "=" + PutCustomIV1.ToString() + "%";

                            if (PutRF1 == "")
                                PutRF1 = "0";
                            worksheet.Cell(12, 11).Formula = "=" + PutRF1.ToString() + "%";

                            if (PutCustomRF1 == "")
                                PutCustomRF1 = "0";
                            worksheet.Cell(12, 12).Formula = "=" + PutCustomRF1.ToString() + "%";

                            if (PutIV2 == "")
                                PutIV2 = "0";
                            worksheet.Cell(12, 13).Formula = "=" + PutIV2.ToString() + "%";

                            if (PutCustomIV2 == "")
                                PutCustomIV2 = "0";
                            worksheet.Cell(12, 14).Formula = "=" + PutCustomIV2.ToString() + "%";

                            if (PutRF2 == "")
                                PutRF2 = "0";
                            worksheet.Cell(12, 15).Formula = "=" + PutRF2.ToString() + "%";

                            if (PutCustomRF2 == "")
                                PutCustomRF2 = "0";
                            worksheet.Cell(12, 16).Formula = "=" + PutCustomRF2.ToString() + "%";
                        }
                    }

                    if (SalesComments != null)
                        worksheet.Cell(14, 2).Value = SalesComments.ToString();
                    else
                        worksheet.Cell(14, 2).Value = "";

                    if (TradingComments != null)
                        worksheet.Cell(15, 2).Value = TradingComments.ToString();
                    else
                        worksheet.Cell(15, 2).Value = "";

                    if (CouponScenario1 != null)
                        worksheet.Cell(16, 2).Value = CouponScenario1.ToString();
                    else
                        worksheet.Cell(16, 2).Value = "";

                    if (CouponScenario2 != null)
                        worksheet.Cell(17, 2).Value = CouponScenario2.ToString();
                    else
                        worksheet.Cell(17, 2).Value = "";

                    //---------------Write Put Spread Strike 1 IV Grid-----------------START------------
                    worksheet.Cell(1, 18).Formula = "=C11";
                    worksheet.Cell(1, 19).Formula = "=ROUND($B$4*30.417,0)";
                    worksheet.Cell(1, 20).Formula = "=ROUND($B$4*30.417,0) - (1*$H$8)";
                    worksheet.Cell(1, 21).Formula = "=ROUND($B$4*30.417,0) - (2*$H$8)";
                    worksheet.Cell(1, 22).Formula = "=ROUND($B$4*30.417,0) - (3*$H$8)";
                    worksheet.Cell(1, 23).Formula = "=ROUND($B$4*30.417,0) - (4*$H$8)";
                    worksheet.Cell(1, 24).Formula = "=ROUND($B$4*30.417,0) - (5*$H$8)";

                    worksheet.Cell(2, 18).Formula = "0";
                    worksheet.Cell(3, 18).Formula = "=$R$2+1*$D$8";
                    worksheet.Cell(4, 18).Formula = "=$R$2+2*$D$8";
                    worksheet.Cell(5, 18).Formula = "=$R$2+3*$D$8";
                    worksheet.Cell(6, 18).Formula = "=$R$2+4*$D$8";
                    worksheet.Cell(7, 18).Formula = "=$R$2+5*$D$8";

                    worksheet.Cell(2, 19).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,S1,IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(3, 19).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($S$1-R3),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(4, 19).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($S$1-R4),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(5, 19).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($S$1-R5),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(6, 19).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($S$1-R6),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(7, 19).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($S$1-R7),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";

                    worksheet.Cell(2, 20).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,T1,IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(3, 20).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($T$1-R3),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(4, 20).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($T$1-R4),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(5, 20).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($T$1-R5),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(6, 20).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($T$1-R6),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(7, 20).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($T$1-R7),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";

                    worksheet.Cell(2, 21).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,U1,IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(3, 21).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($U$1-R3),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(4, 21).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($U$1-R4),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(5, 21).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($U$1-R5),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(6, 21).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($U$1-R6),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(7, 21).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($U$1-R7),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";

                    worksheet.Cell(2, 22).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,V1,IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(3, 22).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($V$1-R3),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(4, 22).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($V$1-R4),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(5, 22).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($V$1-R5),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(6, 22).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($V$1-R6),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(7, 22).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($V$1-R7),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";

                    worksheet.Cell(2, 23).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,W1,IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(3, 23).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($W$1-R3),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(4, 23).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($W$1-R4),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(5, 23).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($W$1-R5),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(6, 23).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($W$1-R6),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(7, 23).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($W$1-R7),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";

                    worksheet.Cell(2, 24).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,X1,IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(3, 24).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($X$1-R3),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(4, 24).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($X$1-R4),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(5, 24).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($X$1-R5),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(6, 24).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($X$1-R6),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    worksheet.Cell(7, 24).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($X$1-R7),IF(J11>0,J11,I11),IF(L11>0,L11,K11))";
                    //---------------Write Put Spread Strike 1 IV Grid-----------------END--------------

                    //---------------Write Put Spread Strike 2 IV Grid-----------------START------------
                    worksheet.Cell(1, 26).Formula = "=D11";
                    worksheet.Cell(1, 27).Formula = "=ROUND($B$4*30.417,0)";
                    worksheet.Cell(1, 28).Formula = "=ROUND($B$4*30.417,0) - (1*$H$8)";
                    worksheet.Cell(1, 29).Formula = "=ROUND($B$4*30.417,0) - (2*$H$8)";
                    worksheet.Cell(1, 30).Formula = "=ROUND($B$4*30.417,0) - (3*$H$8)";
                    worksheet.Cell(1, 31).Formula = "=ROUND($B$4*30.417,0) - (4*$H$8)";
                    worksheet.Cell(1, 32).Formula = "=ROUND($B$4*30.417,0) - (5*$H$8)";

                    worksheet.Cell(2, 26).Formula = "0";
                    worksheet.Cell(3, 26).Formula = "=$Z$2+1*$D$8";
                    worksheet.Cell(4, 26).Formula = "=$Z$2+2*$D$8";
                    worksheet.Cell(5, 26).Formula = "=$Z$2+3*$D$8";
                    worksheet.Cell(6, 26).Formula = "=$Z$2+4*$D$8";
                    worksheet.Cell(7, 26).Formula = "=$Z$2+5*$D$8";

                    worksheet.Cell(2, 27).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$11,100,AA1,IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(3, 27).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$11,100,($AA$1-Z3),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(4, 27).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$11,100,($AA$1-Z4),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(5, 27).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$11,100,($AA$1-Z5),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(6, 27).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$11,100,($AA$1-Z6),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(7, 27).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$11,100,($AA$1-Z7),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";

                    worksheet.Cell(2, 28).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$11,100,AB1,IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(3, 28).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$11,100,($AB$1-Z3),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(4, 28).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$11,100,($AB$1-Z4),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(5, 28).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$11,100,($AB$1-Z5),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(6, 28).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$11,100,($AB$1-Z6),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(7, 28).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$11,100,($AB$1-Z7),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";

                    worksheet.Cell(2, 29).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$11,100,AC1,IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(3, 29).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$11,100,($AC$1-Z3),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(4, 29).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$11,100,($AC$1-Z4),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(5, 29).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$11,100,($AC$1-Z5),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(6, 29).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$11,100,($AC$1-Z6),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(7, 29).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$11,100,($AC$1-Z7),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";

                    worksheet.Cell(2, 30).Formula = "=HoadleyOptions2(\"p\",1,\"p\",$D$11,100,AD1,IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(3, 30).Formula = "=HoadleyOptions2(\"p\",1,\"p\",$D$11,100,($AD$1-Z3),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(4, 30).Formula = "=HoadleyOptions2(\"p\",1,\"p\",$D$11,100,($AD$1-Z4),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(5, 30).Formula = "=HoadleyOptions2(\"p\",1,\"p\",$D$11,100,($AD$1-Z5),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(6, 30).Formula = "=HoadleyOptions2(\"p\",1,\"p\",$D$11,100,($AD$1-Z6),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(7, 30).Formula = "=HoadleyOptions2(\"p\",1,\"p\",$D$11,100,($AD$1-Z7),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";

                    worksheet.Cell(2, 31).Formula = "=HoadleyOptions2(\"p\",1,\"p\",$D$11,100,AE1,IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(3, 31).Formula = "=HoadleyOptions2(\"p\",1,\"p\",$D$11,100,($AE$1-Z3),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(4, 31).Formula = "=HoadleyOptions2(\"p\",1,\"p\",$D$11,100,($AE$1-Z4),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(5, 31).Formula = "=HoadleyOptions2(\"p\",1,\"p\",$D$11,100,($AE$1-Z5),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(6, 31).Formula = "=HoadleyOptions2(\"p\",1,\"p\",$D$11,100,($AE$1-Z6),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(7, 31).Formula = "=HoadleyOptions2(\"p\",1,\"p\",$D$11,100,($AE$1-Z7),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";

                    worksheet.Cell(2, 32).Formula = "=HoadleyOptions2(\"p\",1,\"p\",$D$11,100,AF1,IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(3, 32).Formula = "=HoadleyOptions2(\"p\",1,\"p\",$D$11,100,($AF$1-Z3),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(4, 32).Formula = "=HoadleyOptions2(\"p\",1,\"p\",$D$11,100,($AF$1-Z4),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(5, 32).Formula = "=HoadleyOptions2(\"p\",1,\"p\",$D$11,100,($AF$1-Z5),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(6, 32).Formula = "=HoadleyOptions2(\"p\",1,\"p\",$D$11,100,($AF$1-Z6),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    worksheet.Cell(7, 32).Formula = "=HoadleyOptions2(\"p\",1,\"p\",$D$11,100,($AF$1-Z7),IF(N11>0,N11,M11),IF(P11>0,P11,O11))";
                    //---------------Write Put Spread Strike 2 IV Grid-----------------END--------------

                    //---------------Write Put Strike 1 IV Grid-----------------START------------
                    if (PutStrike1 != "")
                    {
                        worksheet.Cell(9, 18).Formula = "=C12";
                        worksheet.Cell(9, 19).Formula = "=ROUND($B$4*30.417,0)";
                        worksheet.Cell(9, 20).Formula = "=ROUND($B$4*30.417,0) - (1*$H$8)";
                        worksheet.Cell(9, 21).Formula = "=ROUND($B$4*30.417,0) - (2*$H$8)";
                        worksheet.Cell(9, 22).Formula = "=ROUND($B$4*30.417,0) - (3*$H$8)";
                        worksheet.Cell(9, 23).Formula = "=ROUND($B$4*30.417,0) - (4*$H$8)";
                        worksheet.Cell(9, 24).Formula = "=ROUND($B$4*30.417,0) - (5*$H$8)";

                        worksheet.Cell(10, 18).Formula = "0";
                        worksheet.Cell(11, 18).Formula = "=$R$10+1*$D$8";
                        worksheet.Cell(12, 18).Formula = "=$R$10+2*$D$8";
                        worksheet.Cell(13, 18).Formula = "=$R$10+3*$D$8";
                        worksheet.Cell(14, 18).Formula = "=$R$10+4*$D$8";
                        worksheet.Cell(15, 18).Formula = "=$R$10+5*$D$8";

                        worksheet.Cell(10, 19).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,S9,IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(11, 19).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($S$9-R11),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(12, 19).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($S$9-R12),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(13, 19).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($S$9-R13),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(14, 19).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($S$9-R14),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(15, 19).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($S$9-R15),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";

                        worksheet.Cell(10, 20).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,T9,IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(11, 20).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($T$9-R11),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(12, 20).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($T$9-R12),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(13, 20).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($T$9-R13),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(14, 20).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($T$9-R14),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(15, 20).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($T$9-R15),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";

                        worksheet.Cell(10, 21).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,U9,IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(11, 21).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($U$9-R11),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(12, 21).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($U$9-R12),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(13, 21).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($U$9-R13),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(14, 21).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($U$9-R14),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(15, 21).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($U$9-R15),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";

                        worksheet.Cell(10, 22).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,V9,IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(11, 22).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($V$9-R11),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(12, 22).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($V$9-R12),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(13, 22).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($V$9-R13),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(14, 22).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($V$9-R14),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(15, 22).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($V$9-R15),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";

                        worksheet.Cell(10, 23).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,W9,IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(11, 23).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($W$9-R11),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(12, 23).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($W$9-R12),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(13, 23).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($W$9-R13),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(14, 23).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($W$9-R14),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(15, 23).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($W$9-R15),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";

                        worksheet.Cell(10, 24).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,X9,IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(11, 24).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($X$9-R11),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(12, 24).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($X$9-R12),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(13, 24).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($X$9-R13),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(14, 24).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($X$9-R14),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                        worksheet.Cell(15, 24).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$12,100,($X$9-R15),IF(J12>0,J12,I12),IF(L12>0,L12,K12))";
                    }
                    //---------------Write Put Strike 1 IV Grid-----------------END--------------

                    //---------------Write Put Strike 2 IV Grid-----------------START------------
                    if (PutStrike2 != "")
                    {
                        worksheet.Cell(9, 26).Formula = "=D11";
                        worksheet.Cell(9, 27).Formula = "=ROUND($B$4*30.417,0)";
                        worksheet.Cell(9, 28).Formula = "=ROUND($B$4*30.417,0) - (1*$H$8)";
                        worksheet.Cell(9, 29).Formula = "=ROUND($B$4*30.417,0) - (2*$H$8)";
                        worksheet.Cell(9, 30).Formula = "=ROUND($B$4*30.417,0) - (3*$H$8)";
                        worksheet.Cell(9, 31).Formula = "=ROUND($B$4*30.417,0) - (4*$H$8)";
                        worksheet.Cell(9, 32).Formula = "=ROUND($B$4*30.417,0) - (5*$H$8)";

                        worksheet.Cell(10, 26).Formula = "0";
                        worksheet.Cell(11, 26).Formula = "=$Z$10+1*$D$8";
                        worksheet.Cell(12, 26).Formula = "=$Z$10+2*$D$8";
                        worksheet.Cell(13, 26).Formula = "=$Z$10+3*$D$8";
                        worksheet.Cell(14, 26).Formula = "=$Z$10+4*$D$8";
                        worksheet.Cell(15, 26).Formula = "=$Z$10+5*$D$8";

                        worksheet.Cell(10, 27).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,AA9,IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(11, 27).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AA$9-Z11),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(12, 27).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AA$9-Z12),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(13, 27).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AA$9-Z13),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(14, 27).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AA$9-Z14),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(15, 27).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AA$9-Z15),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";

                        worksheet.Cell(10, 28).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,AB9,IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(11, 28).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AB$9-Z11),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(12, 28).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AB$9-Z12),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(13, 28).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AB$9-Z13),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(14, 28).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AB$9-Z14),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(15, 28).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AB$9-Z15),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";

                        worksheet.Cell(10, 29).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,AC9,IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(11, 29).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AC$9-Z11),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(12, 29).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AC$9-Z12),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(13, 29).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AC$9-Z13),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(14, 29).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AC$9-Z14),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(15, 29).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AC$9-Z15),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";

                        worksheet.Cell(10, 30).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,AD9,IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(11, 30).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AD$9-Z11),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(12, 30).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AD$9-Z12),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(13, 30).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AD$9-Z13),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(14, 30).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AD$9-Z14),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(15, 30).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AD$9-Z15),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";

                        worksheet.Cell(10, 31).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,AE9,IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(11, 31).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AE$9-Z11),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(12, 31).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AE$9-Z12),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(13, 31).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AE$9-Z13),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(14, 31).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AE$9-Z14),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(15, 31).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AE$9-Z15),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";

                        worksheet.Cell(10, 32).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,AF9,IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(11, 32).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AF$9-Z11),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(12, 32).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AF$9-Z12),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(13, 32).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AF$9-Z13),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(14, 32).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AF$9-Z14),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                        worksheet.Cell(15, 32).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$D$12,100,($AF$9-Z15),IF(N12>0,N12,M12),IF(P12>0,P12,O12))";
                    }
                    //---------------Write Put Strike 2 IV Grid-----------------END--------------

                    xlPackage.Save();
                }

                return Json("");
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];

                ClearGoldenCushionSession();

                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "ExportGoldenCushion", objUserMaster.UserID);

                return Json("");
            }
        }

        public List<Graph> GenerateGraphGoldenCushion(double PutSpreadStrike1, double PutSpreadStrike2, double PutShortStrike1, double PutShortStrike2, double LowerCoupon, double FixedCoupon, double PutSpreadPR, double PutPR, string PutSpreadOptionType, string PutOptionType)
        {
            var transactionCounts = new List<Graph>();
            try
            {
                DataTable dtGraph = new DataTable();


                dtGraph.Columns.Add("INITIAL");
                dtGraph.Columns.Add("FINAL");
                dtGraph.Columns.Add("NIFTY_PERFORMANCE");
                dtGraph.Columns.Add("PR");

                int Count = 25;

                DataRow dr;
                DataRow lastRow;
                var Initial = 100;
                var Nifty = -100.0;
                var Final = 0.0;
                var Spot = 100;
                var PrevFinal = 0.0;
                var NegSpot = -100;
                var BaseDiffrence = 0.0;
                var val = 0.0;
                var val1 = 0.0;
                var final = 0.0;
                var nifty = 0.0;
                var CoupanDiff = FixedCoupon - LowerCoupon;
                var StrikeDiff = Math.Abs(PutSpreadStrike2 - PutSpreadStrike1);
                var StrikeDiffrencePut = Math.Abs(PutShortStrike2 - PutShortStrike1);
                var Diffrence = Spot - PutShortStrike1;
                bool strike = false;
                bool strike1 = false;
                bool strike2 = false;
                bool strike3 = false;
                bool strike4 = false;
                bool setFlag = false;
                double[] arrStrike;
                double[] finalArray;

                finalArray = new double[21];
                finalArray[0] = 0.0;
                finalArray[1] = 10.0;
                finalArray[2] = 20.0;
                finalArray[3] = 30.0;
                finalArray[4] = 40.0;
                finalArray[5] = 50.0;
                finalArray[6] = 60.0;
                finalArray[7] = 70.0;
                finalArray[8] = 80.0;
                finalArray[9] = 90.0;
                finalArray[10] = 100.0;
                finalArray[11] = 110.0;
                finalArray[12] = 120.0;
                finalArray[13] = 130.0;
                finalArray[14] = 140.0;
                finalArray[15] = 150.0;
                finalArray[16] = 160.0;
                finalArray[17] = 170.0;
                finalArray[18] = 180.0;
                finalArray[19] = 190.0;
                finalArray[20] = 200.0;

                if (PutOptionType == "Put Spread Short")
                {
                    arrStrike = new double[4];
                    arrStrike[0] = PutSpreadStrike1;
                    arrStrike[1] = PutSpreadStrike2;
                    arrStrike[2] = PutShortStrike1;
                    arrStrike[3] = PutShortStrike2;
                    // Array.Sort(arrStrike);

                }
                else
                {
                    arrStrike = new double[4];
                    arrStrike[0] = PutSpreadStrike1;
                    arrStrike[1] = PutSpreadStrike2;
                    arrStrike[2] = PutShortStrike1;
                    arrStrike[3] = 0;
                    // Array.Sort(arrStrike);
                }

                //var transactionCounts = new List<Graph>();
                for (int i = 1; i <= Count; i++)
                {
                    dr = dtGraph.NewRow();
                    lastRow = dtGraph.NewRow();

                    if (strike)
                    {
                        lastRow = dtGraph.Rows[dtGraph.Rows.Count - 1];
                        Nifty = Convert.ToDouble(lastRow[2]);
                        PrevFinal = Convert.ToDouble(lastRow[1]);
                        if (Nifty % 10 == 0)
                        {
                            Nifty = Nifty + 10;
                            Final = Initial + Nifty;
                        }
                        else
                        {
                            for (int l = 0; l < finalArray.Length; l++)
                            {
                                if (finalArray[l] >= PrevFinal)
                                {
                                    Final = finalArray[l];
                                    Nifty = Final - Initial;
                                    break;
                                }
                            }
                        }
                    }


                    dr = dtGraph.NewRow();
                    if (Final == 210)
                    {
                        return transactionCounts;
                    }

                    BaseDiffrence = PutSpreadStrike2 - Spot;
                    var value = (double)(CoupanDiff + (PutSpreadPR * (Nifty - BaseDiffrence)));
                    var value1 = Math.Min(CoupanDiff, value);
                    var value2 = Math.Max(0, value1);
                    var FinalValue = 0.0;
                    if (LowerCoupon > 0)
                    {
                        val = Nifty + Diffrence;
                        val1 = Math.Min(0, val);
                    }
                    if (PutShortStrike1 > 0)
                    {
                        FinalValue = value2 + LowerCoupon + val1;
                    }
                    else
                    {
                        FinalValue = value2 + val1;
                    }
                    //if (Final != 0)
                    //{

                    if (arrStrike[1] < arrStrike[2])
                        setFlag = true;

                    if (arrStrike[0] > PrevFinal && arrStrike[0] < Final && arrStrike[0] % 10 != 0)
                    {
                        #region Strike 1
                        if (Final == 0)
                        {

                            dr["INITIAL"] = Initial;
                            dr["FINAL"] = Final;
                            dr["NIFTY_PERFORMANCE"] = Nifty;

                            FinalValue = CalculateGoldenCushion(Nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);
                            if (FinalValue < NegSpot)
                            {
                                dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                            }
                            else
                            {
                                dr["PR"] = Math.Round(FinalValue, 2);
                            }

                            dtGraph.Rows.Add(dr);
                            transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                        }
                        else
                        {

                            #region Main Row
                            dr = dtGraph.NewRow();
                            dr["INITIAL"] = Initial;
                            var final1 = arrStrike[0];
                            dr["FINAL"] = final1;
                            var nifty1 = Math.Round(final1 - Initial, 2);//Nifty <= 0 ? (Nifty - (arrStrike[0] % 10)) : (Nifty + (arrStrike[0] % 10));
                            dr["NIFTY_PERFORMANCE"] = nifty1;


                            FinalValue = CalculateGoldenCushion(nifty1, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);
                            //dr["PR"] = FinalValue;
                            if (FinalValue < NegSpot)
                            {
                                dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                            }
                            else
                            {
                                dr["PR"] = Math.Round(FinalValue, 2);
                            }

                            dtGraph.Rows.Add(dr);
                            transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                            #endregion

                            #region Next Row
                            dr = dtGraph.NewRow();
                            dr["INITIAL"] = Initial;
                            final = arrStrike[0] + 0.01;//(double)(Initial + ((100 * Nifty) / 100)) + 0.01;
                            dr["FINAL"] = final;
                            nifty = Math.Round(final - Initial, 2);//Nifty <= 0 ? (Nifty - (arrStrike[0] % 10)) : (Nifty + (arrStrike[0] % 10));
                            // nifty = nifty + 0.01;
                            dr["NIFTY_PERFORMANCE"] = nifty;

                            FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);

                            if (FinalValue < NegSpot)
                            {
                                dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                            }
                            else
                            {
                                dr["PR"] = Math.Round(FinalValue, 2);
                            }
                            dtGraph.Rows.Add(dr);
                            transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });

                            #endregion

                            Nifty = Nifty - 10;

                            strike1 = true;
                        }
                        #endregion
                    }
                    else if (arrStrike[1] > PrevFinal && arrStrike[1] < Final && arrStrike[1] % 10 != 0 && setFlag == true)
                    {
                        #region Strike 2
                        //  Nifty = Nifty - 10;
                        if (Final == 0)
                        {

                            dr["INITIAL"] = Initial;
                            dr["FINAL"] = Final;
                            dr["NIFTY_PERFORMANCE"] = Nifty;

                            FinalValue = CalculateGoldenCushion(Nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);
                            if (FinalValue < NegSpot)
                            {
                                dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                            }
                            else
                            {
                                dr["PR"] = Math.Round(FinalValue, 2);
                            }

                            dtGraph.Rows.Add(dr);
                            transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                        }
                        else
                        {
                            var Val = 0;
                            #region Less Than 15
                            if (StrikeDiff < 15)
                            {
                                for (int a = 0; a < 10; a++)
                                {
                                    dr = dtGraph.NewRow();
                                    dr["INITIAL"] = Initial;
                                    final = arrStrike[1] + a;
                                    dr["FINAL"] = final;

                                    nifty = Math.Round(final - Initial, 2);
                                    dr["NIFTY_PERFORMANCE"] = nifty;

                                    FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);
                                    if (FinalValue < NegSpot)
                                    {
                                        dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                    }
                                    else
                                    {
                                        dr["PR"] = Math.Round(FinalValue, 2);
                                    }
                                    if (final < PutSpreadStrike1)
                                    {
                                        dtGraph.Rows.Add(dr);
                                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                                    }
                                    if (PutOptionType != "0")
                                    {
                                        if (final == arrStrike[3] && arrStrike[3] < arrStrike[0])
                                        {
                                            if (arrStrike[2] >= (final + a))
                                                a = a + Convert.ToInt32(arrStrike[2] - final);
                                            #region Strike 4
                                            #region Less than 15
                                            if (StrikeDiffrencePut < 15)
                                            {
                                                for (int b = 1; b < 10; b++)
                                                {
                                                    dr = dtGraph.NewRow();
                                                    dr["INITIAL"] = Initial;
                                                    nifty = (arrStrike[3] + b) - Initial;
                                                    final = arrStrike[3] + b;
                                                    dr["FINAL"] = final;

                                                    dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);

                                                    FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);
                                                    if (FinalValue < NegSpot)
                                                    {
                                                        dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                                    }
                                                    else
                                                    {
                                                        dr["PR"] = Math.Round(FinalValue, 2);
                                                    }

                                                    if (final < PutShortStrike1)
                                                    {
                                                        dtGraph.Rows.Add(dr);
                                                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                                                    }
                                                    if (final == arrStrike[2] && arrStrike[2] < arrStrike[0])
                                                    {
                                                        #region Main Row
                                                        dr = dtGraph.NewRow();
                                                        dr["INITIAL"] = Initial;
                                                        dr["FINAL"] = final;
                                                        nifty = final - Initial;
                                                        dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);
                                                        FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);
                                                        //dr["PR"] = FinalValue;
                                                        if (FinalValue < NegSpot)
                                                        {
                                                            dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                                        }
                                                        else
                                                        {
                                                            dr["PR"] = Math.Round(FinalValue, 2);
                                                        }

                                                        dtGraph.Rows.Add(dr);
                                                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                                                        #endregion

                                                        #region Next Row
                                                        dr = dtGraph.NewRow();
                                                        dr["INITIAL"] = Initial;
                                                        final = final + 0.01;
                                                        dr["FINAL"] = final;
                                                        nifty = final - Initial;
                                                        dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);

                                                        FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);

                                                        if (FinalValue < NegSpot)
                                                        {
                                                            dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                                        }
                                                        else
                                                        {
                                                            dr["PR"] = Math.Round(FinalValue, 2);
                                                        }
                                                        dtGraph.Rows.Add(dr);
                                                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });

                                                        #endregion
                                                    }
                                                }
                                                //Nifty = Nifty - 10;
                                            }
                                            #endregion

                                            #region Greater than 15
                                            else
                                            {
                                                int e = 0;
                                                e = Convert.ToInt32(StrikeDiffrencePut / 5);
                                                e = e + 1;
                                                bool flag1 = false;
                                                int addedValue1 = 0;
                                                for (int c = 0; c < e; c++)
                                                {
                                                    if (flag1)
                                                        addedValue1 = addedValue1 + 5;

                                                    dr = dtGraph.NewRow();
                                                    dr["INITIAL"] = Initial;
                                                    nifty = arrStrike[3] - Initial;
                                                    final = arrStrike[3] + addedValue1;
                                                    dr["FINAL"] = final;

                                                    nifty = nifty + addedValue1;
                                                    dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);

                                                    FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);
                                                    if (FinalValue < NegSpot)
                                                    {
                                                        dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                                    }
                                                    else
                                                    {
                                                        dr["PR"] = Math.Round(FinalValue, 2);
                                                    }

                                                    if (final < PutShortStrike1)
                                                    {
                                                        dtGraph.Rows.Add(dr);
                                                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                                                    }
                                                    if (final == arrStrike[2] && arrStrike[2] < arrStrike[0])
                                                    {
                                                        #region Main Row
                                                        dr = dtGraph.NewRow();
                                                        dr["INITIAL"] = Initial;
                                                        dr["FINAL"] = final;
                                                        nifty = final - Initial;
                                                        dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);
                                                        FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);
                                                        //dr["PR"] = FinalValue;
                                                        if (FinalValue < NegSpot)
                                                        {
                                                            dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                                        }
                                                        else
                                                        {
                                                            dr["PR"] = Math.Round(FinalValue, 2);
                                                        }

                                                        dtGraph.Rows.Add(dr);
                                                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                                                        #endregion

                                                        #region Next Row
                                                        dr = dtGraph.NewRow();
                                                        dr["INITIAL"] = Initial;
                                                        final = final + 0.01;
                                                        dr["FINAL"] = final;
                                                        nifty = final - Initial;
                                                        dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);

                                                        FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);

                                                        if (FinalValue < NegSpot)
                                                        {
                                                            dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                                        }
                                                        else
                                                        {
                                                            dr["PR"] = Math.Round(FinalValue, 2);
                                                        }
                                                        dtGraph.Rows.Add(dr);
                                                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });

                                                        #endregion
                                                    }
                                                }

                                            }
                                            #endregion
                                            strike4 = true;

                                            #endregion
                                        }
                                        if (final == arrStrike[2])
                                        {
                                            #region Strike3
                                            // Nifty = Nifty - 10;
                                            if (Final == 0)
                                            {
                                                dr["INITIAL"] = Initial;
                                                dr["FINAL"] = Final;
                                                dr["NIFTY_PERFORMANCE"] = Nifty;

                                                FinalValue = CalculateGoldenCushion(Nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);
                                                if (FinalValue < NegSpot)
                                                {
                                                    dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                                }
                                                else
                                                {
                                                    dr["PR"] = Math.Round(FinalValue, 2);
                                                }

                                                dtGraph.Rows.Add(dr);
                                                transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                                            }
                                            else
                                            {

                                                #region Next Row
                                                dr = dtGraph.NewRow();
                                                dr["INITIAL"] = Initial;
                                                final = arrStrike[2];//(double)(Initial + ((100 * Nifty) / 100)) + 0.01;
                                                dr["FINAL"] = final + 0.01;
                                                nifty = final - Initial;//Nifty <= 0 ? (Nifty - (arrStrike[2] % 10)) : (Nifty + (arrStrike[2] % 10));
                                                nifty = nifty + 0.01;
                                                dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);

                                                FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);

                                                if (FinalValue < NegSpot)
                                                {
                                                    dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                                }
                                                else
                                                {
                                                    dr["PR"] = Math.Round(FinalValue, 2);
                                                }
                                                dtGraph.Rows.Add(dr);
                                                transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });

                                                #endregion

                                                strike3 = true;
                                            }
                                            #endregion
                                        }
                                    }
                                }
                                ////Nifty = Nifty - 10;

                                //var n = Convert.ToInt32(final - Initial);
                                //Nifty = n - 20;
                                //Nifty = Nifty + 10;
                            }
                            #endregion

                            #region Greater than 15
                            else
                            {
                                int d = 0;
                                d = Convert.ToInt32(StrikeDiff / 5);
                                d = d + 1;
                                bool flag = false;
                                int addedValue = 0;
                                for (int a = 0; a < d; a++)
                                {
                                    if (flag)
                                        addedValue = addedValue + 5;

                                    dr = dtGraph.NewRow();
                                    dr["INITIAL"] = Initial;
                                    final = arrStrike[1] + addedValue;
                                    dr["FINAL"] = final;

                                    nifty = final - Initial;
                                    dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);

                                    FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);
                                    if (FinalValue < NegSpot)
                                    {
                                        dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                    }
                                    else
                                    {
                                        dr["PR"] = Math.Round(FinalValue, 2);
                                    }
                                    if (final < PutSpreadStrike1)
                                    {
                                        dtGraph.Rows.Add(dr);
                                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                                    }
                                    if (PutOptionType != "0")
                                    {
                                        if (arrStrike[2] >= (final + addedValue))
                                            addedValue = addedValue + Convert.ToInt32(arrStrike[2] - final);
                                        if (final == arrStrike[3] && arrStrike[3] < arrStrike[0])
                                        {

                                            #region Strike 4
                                            #region Less than 15
                                            if (StrikeDiffrencePut < 15)
                                            {
                                                for (int b = 1; b < 10; b++)
                                                {
                                                    dr = dtGraph.NewRow();
                                                    dr["INITIAL"] = Initial;
                                                    nifty = (arrStrike[3] + b) - Initial;
                                                    final = arrStrike[3] + b;
                                                    dr["FINAL"] = final;

                                                    dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);

                                                    FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);
                                                    if (FinalValue < NegSpot)
                                                    {
                                                        dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                                    }
                                                    else
                                                    {
                                                        dr["PR"] = Math.Round(FinalValue, 2);
                                                    }

                                                    if (final < PutShortStrike1)
                                                    {
                                                        dtGraph.Rows.Add(dr);
                                                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                                                    }
                                                    if (final == arrStrike[2] && arrStrike[2] < arrStrike[0])
                                                    {
                                                        #region Main Row
                                                        dr = dtGraph.NewRow();
                                                        dr["INITIAL"] = Initial;
                                                        dr["FINAL"] = final;
                                                        nifty = final - Initial;
                                                        dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);
                                                        FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);
                                                        //dr["PR"] = FinalValue;
                                                        if (FinalValue < NegSpot)
                                                        {
                                                            dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                                        }
                                                        else
                                                        {
                                                            dr["PR"] = Math.Round(FinalValue, 2);
                                                        }

                                                        dtGraph.Rows.Add(dr);
                                                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                                                        #endregion

                                                        #region Next Row
                                                        dr = dtGraph.NewRow();
                                                        dr["INITIAL"] = Initial;
                                                        final = final + 0.01;
                                                        dr["FINAL"] = final;
                                                        nifty = final - Initial;
                                                        dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);

                                                        FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);

                                                        if (FinalValue < NegSpot)
                                                        {
                                                            dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                                        }
                                                        else
                                                        {
                                                            dr["PR"] = Math.Round(FinalValue, 2);
                                                        }
                                                        dtGraph.Rows.Add(dr);
                                                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });

                                                        #endregion
                                                    }
                                                }
                                                //Nifty = Nifty - 10;
                                            }
                                            #endregion

                                            #region Greater than 15
                                            else
                                            {
                                                int e = 0;
                                                e = Convert.ToInt32(StrikeDiffrencePut / 5);
                                                e = e + 1;
                                                bool flag1 = false;
                                                int addedValue1 = 0;
                                                for (int c = 0; c < e; c++)
                                                {
                                                    if (flag1)
                                                        addedValue1 = addedValue1 + 5;

                                                    dr = dtGraph.NewRow();
                                                    dr["INITIAL"] = Initial;
                                                    nifty = arrStrike[3] - Initial;
                                                    final = arrStrike[3] + addedValue1;
                                                    dr["FINAL"] = final;

                                                    nifty = nifty + addedValue1;
                                                    dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);

                                                    FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);
                                                    if (FinalValue < NegSpot)
                                                    {
                                                        dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                                    }
                                                    else
                                                    {
                                                        dr["PR"] = Math.Round(FinalValue, 2);
                                                    }

                                                    if (final < PutShortStrike1)
                                                    {
                                                        dtGraph.Rows.Add(dr);
                                                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                                                    }
                                                    if (final == arrStrike[2] && arrStrike[2] < arrStrike[0])
                                                    {
                                                        #region Main Row
                                                        dr = dtGraph.NewRow();
                                                        dr["INITIAL"] = Initial;
                                                        dr["FINAL"] = final;
                                                        nifty = final - Initial;
                                                        dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);
                                                        FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);
                                                        //dr["PR"] = FinalValue;
                                                        if (FinalValue < NegSpot)
                                                        {
                                                            dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                                        }
                                                        else
                                                        {
                                                            dr["PR"] = Math.Round(FinalValue, 2);
                                                        }

                                                        dtGraph.Rows.Add(dr);
                                                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                                                        #endregion

                                                        #region Next Row
                                                        dr = dtGraph.NewRow();
                                                        dr["INITIAL"] = Initial;
                                                        final = final + 0.01;
                                                        dr["FINAL"] = final;
                                                        nifty = final - Initial;
                                                        dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);

                                                        FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);

                                                        if (FinalValue < NegSpot)
                                                        {
                                                            dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                                        }
                                                        else
                                                        {
                                                            dr["PR"] = Math.Round(FinalValue, 2);
                                                        }
                                                        dtGraph.Rows.Add(dr);
                                                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });

                                                        #endregion
                                                    }
                                                    flag = true;
                                                }

                                            }
                                            #endregion
                                            strike4 = true;

                                            #endregion
                                        }
                                        if (arrStrike[3] > Val && arrStrike[3] < Final)
                                        {

                                            #region Strike 4
                                            #region Less than 15
                                            if (StrikeDiffrencePut < 15)
                                            {
                                                for (int b = 1; b < 10; b++)
                                                {
                                                    dr = dtGraph.NewRow();
                                                    dr["INITIAL"] = Initial;
                                                    nifty = (arrStrike[3] + b) - Initial;
                                                    final = arrStrike[3] + b;
                                                    dr["FINAL"] = final;

                                                    dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);

                                                    FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);
                                                    if (FinalValue < NegSpot)
                                                    {
                                                        dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                                    }
                                                    else
                                                    {
                                                        dr["PR"] = Math.Round(FinalValue, 2);
                                                    }

                                                    if (final < PutShortStrike1)
                                                    {
                                                        dtGraph.Rows.Add(dr);
                                                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                                                    }
                                                    if (final == arrStrike[2] && arrStrike[2] < arrStrike[0])
                                                    {
                                                        #region Main Row
                                                        dr = dtGraph.NewRow();
                                                        dr["INITIAL"] = Initial;
                                                        dr["FINAL"] = final;
                                                        nifty = final - Initial;
                                                        dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);
                                                        FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);
                                                        //dr["PR"] = FinalValue;
                                                        if (FinalValue < NegSpot)
                                                        {
                                                            dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                                        }
                                                        else
                                                        {
                                                            dr["PR"] = Math.Round(FinalValue, 2);
                                                        }

                                                        dtGraph.Rows.Add(dr);
                                                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                                                        #endregion

                                                        #region Next Row
                                                        dr = dtGraph.NewRow();
                                                        dr["INITIAL"] = Initial;
                                                        final = final + 0.01;
                                                        dr["FINAL"] = final;
                                                        nifty = final - Initial;
                                                        dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);

                                                        FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);

                                                        if (FinalValue < NegSpot)
                                                        {
                                                            dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                                        }
                                                        else
                                                        {
                                                            dr["PR"] = Math.Round(FinalValue, 2);
                                                        }
                                                        dtGraph.Rows.Add(dr);
                                                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });

                                                        #endregion
                                                    }
                                                }
                                                //Nifty = Nifty - 10;
                                            }
                                            #endregion

                                            #region Greater than 15
                                            else
                                            {
                                                int e = 0;
                                                e = Convert.ToInt32(StrikeDiffrencePut / 5);
                                                e = e + 1;
                                                bool flag1 = false;
                                                int addedValue1 = 0;
                                                for (int c = 0; c < e; c++)
                                                {
                                                    if (flag1)
                                                        addedValue1 = addedValue1 + 5;

                                                    dr = dtGraph.NewRow();
                                                    dr["INITIAL"] = Initial;
                                                    nifty = arrStrike[3] - Initial;
                                                    final = arrStrike[3] + addedValue1;
                                                    dr["FINAL"] = final;

                                                    nifty = nifty + addedValue1;
                                                    dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);

                                                    FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);
                                                    if (FinalValue < NegSpot)
                                                    {
                                                        dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                                    }
                                                    else
                                                    {
                                                        dr["PR"] = Math.Round(FinalValue, 2);
                                                    }

                                                    if (final < PutShortStrike1)
                                                    {
                                                        dtGraph.Rows.Add(dr);
                                                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                                                    }
                                                    if (final == arrStrike[2] && arrStrike[2] < arrStrike[0])
                                                    {
                                                        #region Main Row
                                                        dr = dtGraph.NewRow();
                                                        dr["INITIAL"] = Initial;
                                                        dr["FINAL"] = final;
                                                        nifty = final - Initial;
                                                        dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);
                                                        FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);
                                                        //dr["PR"] = FinalValue;
                                                        if (FinalValue < NegSpot)
                                                        {
                                                            dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                                        }
                                                        else
                                                        {
                                                            dr["PR"] = Math.Round(FinalValue, 2);
                                                        }

                                                        dtGraph.Rows.Add(dr);
                                                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                                                        #endregion

                                                        #region Next Row
                                                        dr = dtGraph.NewRow();
                                                        dr["INITIAL"] = Initial;
                                                        final = final + 0.01;
                                                        dr["FINAL"] = final;
                                                        nifty = final - Initial;
                                                        dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);

                                                        FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);

                                                        if (FinalValue < NegSpot)
                                                        {
                                                            dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                                        }
                                                        else
                                                        {
                                                            dr["PR"] = Math.Round(FinalValue, 2);
                                                        }
                                                        dtGraph.Rows.Add(dr);
                                                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });

                                                        #endregion
                                                    }
                                                    flag = true;
                                                }

                                            }
                                            #endregion
                                            strike4 = true;

                                            #endregion
                                        }
                                        if (final == arrStrike[2])
                                        {
                                            #region Strike3
                                            // Nifty = Nifty - 10;
                                            if (Final == 0)
                                            {
                                                dr["INITIAL"] = Initial;
                                                dr["FINAL"] = Final;
                                                dr["NIFTY_PERFORMANCE"] = Nifty;

                                                FinalValue = CalculateGoldenCushion(Nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);
                                                if (FinalValue < NegSpot)
                                                {
                                                    dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                                }
                                                else
                                                {
                                                    dr["PR"] = Math.Round(FinalValue, 2);
                                                }

                                                dtGraph.Rows.Add(dr);
                                                transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                                            }
                                            else
                                            {
                                                #region Main Row
                                                dr = dtGraph.NewRow();
                                                dr["INITIAL"] = Initial;
                                                var final1 = arrStrike[2];
                                                dr["FINAL"] = final1;
                                                var nifty1 = final1 - Initial;//Nifty <= 0 ? (Nifty - (arrStrike[2] % 10)) : (Nifty + (arrStrike[2] % 10));
                                                dr["NIFTY_PERFORMANCE"] = Math.Round(nifty1, 2);


                                                FinalValue = CalculateGoldenCushion(nifty1, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);
                                                //dr["PR"] = FinalValue;
                                                if (FinalValue < NegSpot)
                                                {
                                                    dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                                }
                                                else
                                                {
                                                    dr["PR"] = Math.Round(FinalValue, 2);
                                                }

                                                dtGraph.Rows.Add(dr);
                                                transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                                                #endregion

                                                #region Next Row
                                                dr = dtGraph.NewRow();
                                                dr["INITIAL"] = Initial;
                                                final = arrStrike[2];//(double)(Initial + ((100 * Nifty) / 100)) + 0.01;
                                                dr["FINAL"] = final + 0.01;
                                                nifty = final - Initial;//Nifty <= 0 ? (Nifty - (arrStrike[2] % 10)) : (Nifty + (arrStrike[2] % 10));
                                                nifty = nifty + 0.01;
                                                dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);

                                                FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);

                                                if (FinalValue < NegSpot)
                                                {
                                                    dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                                }
                                                else
                                                {
                                                    dr["PR"] = Math.Round(FinalValue, 2);
                                                }
                                                dtGraph.Rows.Add(dr);
                                                transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });

                                                #endregion

                                                strike3 = true;
                                            }
                                            #endregion
                                        }

                                        Val = Convert.ToInt32(Final);
                                    }
                                    flag = true;
                                }
                                //if (PutOptionType == "0")
                                //    Nifty = Nifty + 10;
                            }
                            #endregion
                            strike2 = true;

                        }
                        #endregion
                    }
                    else if (arrStrike[2] > PrevFinal && arrStrike[2] < Final && arrStrike[2] % 10 != 0)
                    {
                        #region Strike3
                        // Nifty = Nifty - 10;
                        if (Final == 0)
                        {
                            dr["INITIAL"] = Initial;
                            dr["FINAL"] = Final;
                            dr["NIFTY_PERFORMANCE"] = Nifty;

                            FinalValue = CalculateGoldenCushion(Nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);
                            if (FinalValue < NegSpot)
                            {
                                dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                            }
                            else
                            {
                                dr["PR"] = Math.Round(FinalValue, 2);
                            }

                            dtGraph.Rows.Add(dr);
                            transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                        }
                        else
                        {
                            #region Main Row
                            dr = dtGraph.NewRow();
                            dr["INITIAL"] = Initial;
                            var final1 = arrStrike[2];
                            dr["FINAL"] = final1;
                            var nifty1 = final1 - Initial;//Nifty <= 0 ? (Nifty - (arrStrike[2] % 10)) : (Nifty + (arrStrike[2] % 10));
                            dr["NIFTY_PERFORMANCE"] = Math.Round(nifty1, 2);


                            FinalValue = CalculateGoldenCushion(nifty1, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);
                            //dr["PR"] = FinalValue;
                            if (FinalValue < NegSpot)
                            {
                                dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                            }
                            else
                            {
                                dr["PR"] = Math.Round(FinalValue, 2);
                            }

                            dtGraph.Rows.Add(dr);
                            transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                            #endregion

                            #region Next Row
                            dr = dtGraph.NewRow();
                            dr["INITIAL"] = Initial;
                            final = arrStrike[2];//(double)(Initial + ((100 * Nifty) / 100)) + 0.01;
                            dr["FINAL"] = final + 0.01;
                            nifty = final - Initial;//Nifty <= 0 ? (Nifty - (arrStrike[2] % 10)) : (Nifty + (arrStrike[2] % 10));
                            nifty = nifty + 0.01;
                            dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);

                            FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);

                            if (FinalValue < NegSpot)
                            {
                                dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                            }
                            else
                            {
                                dr["PR"] = Math.Round(FinalValue, 2);
                            }
                            dtGraph.Rows.Add(dr);
                            transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });

                            #endregion

                            strike3 = true;
                        }
                        #endregion
                        setFlag = true;
                    }
                    else if (arrStrike.Length == 4 && arrStrike[3] > PrevFinal && arrStrike[3] < Final && arrStrike[3] % 10 != 0)
                    {
                        #region Strike 4
                        // Nifty = Nifty - 10;
                        if (Final == 0)
                        {

                            dr["INITIAL"] = Initial;
                            dr["FINAL"] = Final;
                            dr["NIFTY_PERFORMANCE"] = Nifty;

                            FinalValue = CalculateGoldenCushion(Nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);
                            if (FinalValue < NegSpot)
                            {
                                dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                            }
                            else
                            {
                                dr["PR"] = Math.Round(FinalValue, 2);
                            }

                            dtGraph.Rows.Add(dr);
                            transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                        }
                        else
                        {
                            if (StrikeDiffrencePut < 15)
                            {
                                for (int a = 0; a < StrikeDiffrencePut; a++)
                                {
                                    dr = dtGraph.NewRow();
                                    dr["INITIAL"] = Initial;
                                    nifty = arrStrike[3] - Initial;
                                    final = arrStrike[3] + a;
                                    dr["FINAL"] = final;

                                    nifty = nifty + a;
                                    dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);

                                    FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);
                                    if (FinalValue < NegSpot)
                                    {
                                        dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                    }
                                    else
                                    {
                                        dr["PR"] = Math.Round(FinalValue, 2);
                                    }

                                    if (final < PutShortStrike1)
                                    {
                                        dtGraph.Rows.Add(dr);
                                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                                    }
                                }
                                //Nifty = Nifty - 10;
                            }
                            else
                            {
                                int d = 0;
                                d = Convert.ToInt32(StrikeDiffrencePut / 5);
                                d = d + 1;
                                bool flag = false;
                                int addedValue = 0;
                                for (int a = 0; a < d; a++)
                                {
                                    if (flag)
                                        addedValue = addedValue + 5;

                                    dr = dtGraph.NewRow();
                                    dr["INITIAL"] = Initial;
                                    nifty = arrStrike[3] - Initial;
                                    final = arrStrike[3] + addedValue;
                                    dr["FINAL"] = final;

                                    nifty = nifty + addedValue;
                                    dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);

                                    FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);
                                    if (FinalValue < NegSpot)
                                    {
                                        dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                    }
                                    else
                                    {
                                        dr["PR"] = Math.Round(FinalValue, 2);
                                    }

                                    if (final < PutShortStrike1)
                                    {
                                        dtGraph.Rows.Add(dr);
                                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                                    }
                                    flag = true;
                                }

                            }
                            strike4 = true;
                        }
                        #endregion
                    }
                    else if (arrStrike.Contains(Final))
                    {
                        #region Array Contains

                        #region Final = 0
                        if (Final == 0)
                        {

                            dr["INITIAL"] = Initial;
                            dr["FINAL"] = Final;
                            dr["NIFTY_PERFORMANCE"] = Nifty;

                            FinalValue = CalculateGoldenCushion(Nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);
                            if (FinalValue < NegSpot)
                            {
                                dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                            }
                            else
                            {
                                dr["PR"] = Math.Round(FinalValue, 2);
                            }

                            dtGraph.Rows.Add(dr);
                            transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });

                            strike = true;

                        }
                        #endregion
                        else
                        {
                            if (Final == arrStrike[0])
                                strike1 = true;

                            if (Final == arrStrike[1])
                                strike2 = true;

                            if (Final == arrStrike[2])
                                strike3 = true;

                            if (Final == arrStrike[3])
                                strike4 = true;


                            if (strike1)
                            {
                                #region Main Row
                                dr = dtGraph.NewRow();
                                dr["INITIAL"] = Initial;
                                dr["FINAL"] = (double)(Initial + ((100 * Nifty) / 100));
                                dr["NIFTY_PERFORMANCE"] = Nifty;
                                FinalValue = CalculateGoldenCushion(Nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);
                                //dr["PR"] = FinalValue;
                                if (FinalValue < NegSpot)
                                {
                                    dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                }
                                else
                                {
                                    dr["PR"] = Math.Round(FinalValue, 2);
                                }

                                dtGraph.Rows.Add(dr);
                                transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                                #endregion

                                #region Next Row
                                dr = dtGraph.NewRow();
                                dr["INITIAL"] = Initial;
                                final = (double)(Initial + ((100 * Nifty) / 100)) + 0.01;
                                dr["FINAL"] = final;
                                nifty = Nifty + 0.01; ;
                                dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);

                                FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);

                                if (FinalValue < NegSpot)
                                {
                                    dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                }
                                else
                                {
                                    dr["PR"] = Math.Round(FinalValue, 2);
                                }
                                dtGraph.Rows.Add(dr);
                                transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });

                                #endregion
                            }
                            #region Put Spread Leg 1
                            else if (strike2)
                            {
                                #region Less Than 15
                                if (StrikeDiff < 15)
                                {
                                    for (int a = 0; a < 10; a++)
                                    {
                                        dr = dtGraph.NewRow();
                                        dr["INITIAL"] = Initial;
                                        final = (Initial + Nifty) + a;
                                        dr["FINAL"] = final;

                                        nifty = Nifty + a;
                                        dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);

                                        FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);
                                        if (FinalValue < NegSpot)
                                        {
                                            dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                        }
                                        else
                                        {
                                            dr["PR"] = Math.Round(FinalValue, 2);
                                        }
                                        if (final < PutSpreadStrike1)
                                        {
                                            dtGraph.Rows.Add(dr);
                                            transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                                        }
                                        if (final == arrStrike[3] && arrStrike[3] < arrStrike[0])
                                        {
                                            if (arrStrike[2] >= (final + a))
                                                a = a + Convert.ToInt32(arrStrike[2] - final);
                                            #region Strike 4
                                            #region Less than 15
                                            if (StrikeDiffrencePut < 15)
                                            {
                                                for (int b = 1; b < 10; b++)
                                                {
                                                    dr = dtGraph.NewRow();
                                                    dr["INITIAL"] = Initial;
                                                    nifty = (arrStrike[3] + b) - Initial;
                                                    final = arrStrike[3] + b;
                                                    dr["FINAL"] = final;

                                                    dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);

                                                    FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);
                                                    if (FinalValue < NegSpot)
                                                    {
                                                        dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                                    }
                                                    else
                                                    {
                                                        dr["PR"] = Math.Round(FinalValue, 2);
                                                    }

                                                    if (final < PutShortStrike1)
                                                    {
                                                        dtGraph.Rows.Add(dr);
                                                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                                                    }
                                                    if (final == arrStrike[2] && arrStrike[2] < arrStrike[0])
                                                    {
                                                        #region Main Row
                                                        dr = dtGraph.NewRow();
                                                        dr["INITIAL"] = Initial;
                                                        dr["FINAL"] = final;
                                                        nifty = final - Initial;
                                                        dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);
                                                        FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);
                                                        //dr["PR"] = FinalValue;
                                                        if (FinalValue < NegSpot)
                                                        {
                                                            dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                                        }
                                                        else
                                                        {
                                                            dr["PR"] = Math.Round(FinalValue, 2);
                                                        }

                                                        dtGraph.Rows.Add(dr);
                                                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                                                        #endregion

                                                        #region Next Row
                                                        dr = dtGraph.NewRow();
                                                        dr["INITIAL"] = Initial;
                                                        final = final + 0.01;
                                                        dr["FINAL"] = final;
                                                        nifty = final - Initial;
                                                        dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);

                                                        FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);

                                                        if (FinalValue < NegSpot)
                                                        {
                                                            dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                                        }
                                                        else
                                                        {
                                                            dr["PR"] = Math.Round(FinalValue, 2);
                                                        }
                                                        dtGraph.Rows.Add(dr);
                                                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });

                                                        #endregion

                                                        Final = Final + 10;
                                                    }
                                                }
                                                //Nifty = Nifty - 10;
                                            }
                                            #endregion

                                            #region Greater than 15
                                            else
                                            {
                                                int e = 0;
                                                e = Convert.ToInt32(StrikeDiffrencePut / 5);
                                                e = e + 1;
                                                bool flag1 = false;
                                                int addedValue1 = 0;
                                                for (int c = 0; c < e; c++)
                                                {
                                                    if (flag1)
                                                        addedValue1 = addedValue1 + 5;

                                                    dr = dtGraph.NewRow();
                                                    dr["INITIAL"] = Initial;
                                                    nifty = arrStrike[3] - Initial;
                                                    final = arrStrike[3] + addedValue1;
                                                    dr["FINAL"] = final;

                                                    nifty = nifty + addedValue1;
                                                    dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);

                                                    FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);
                                                    if (FinalValue < NegSpot)
                                                    {
                                                        dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                                    }
                                                    else
                                                    {
                                                        dr["PR"] = Math.Round(FinalValue, 2);
                                                    }

                                                    if (final < PutShortStrike1)
                                                    {
                                                        dtGraph.Rows.Add(dr);
                                                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                                                    }
                                                    if (final == arrStrike[2] && arrStrike[2] < arrStrike[0])
                                                    {
                                                        #region Main Row
                                                        dr = dtGraph.NewRow();
                                                        dr["INITIAL"] = Initial;
                                                        dr["FINAL"] = final;
                                                        nifty = final - Initial;
                                                        dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);
                                                        FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);
                                                        //dr["PR"] = FinalValue;
                                                        if (FinalValue < NegSpot)
                                                        {
                                                            dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                                        }
                                                        else
                                                        {
                                                            dr["PR"] = Math.Round(FinalValue, 2);
                                                        }

                                                        dtGraph.Rows.Add(dr);
                                                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                                                        #endregion

                                                        #region Next Row
                                                        dr = dtGraph.NewRow();
                                                        dr["INITIAL"] = Initial;
                                                        final = final + 0.01;
                                                        dr["FINAL"] = final;
                                                        nifty = final - Initial;
                                                        dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);

                                                        FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);

                                                        if (FinalValue < NegSpot)
                                                        {
                                                            dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                                        }
                                                        else
                                                        {
                                                            dr["PR"] = Math.Round(FinalValue, 2);
                                                        }
                                                        dtGraph.Rows.Add(dr);
                                                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });

                                                        #endregion

                                                        Final = Final + 10;
                                                    }
                                                }

                                            }
                                            #endregion
                                            strike4 = true;

                                            #endregion
                                        }
                                    }
                                    Nifty = Nifty + 10;

                                }
                                #endregion

                                #region Greater than 15
                                else
                                {
                                    int d = 0;
                                    d = Convert.ToInt32(StrikeDiff / 5);
                                    d = d + 1;
                                    bool flag = false;
                                    int addedValue = 0;
                                    for (int a = 0; a < d; a++)
                                    {
                                        if (flag)
                                            addedValue = addedValue + 5;

                                        dr = dtGraph.NewRow();
                                        dr["INITIAL"] = Initial;
                                        final = (Initial + Nifty) + addedValue;
                                        dr["FINAL"] = final;

                                        nifty = Nifty + addedValue;
                                        dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);

                                        FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);
                                        if (FinalValue < NegSpot)
                                        {
                                            dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                        }
                                        else
                                        {
                                            dr["PR"] = Math.Round(FinalValue, 2);
                                        }
                                        if (final < PutSpreadStrike1)
                                        {
                                            dtGraph.Rows.Add(dr);
                                            transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                                        }
                                        if (final == arrStrike[3] && arrStrike[3] < arrStrike[0])
                                        {
                                            if (arrStrike[2] >= (final + addedValue))
                                                addedValue = addedValue + Convert.ToInt32(arrStrike[2] - final);
                                            #region Strike 4
                                            #region Less than 15
                                            if (StrikeDiffrencePut < 15)
                                            {
                                                for (int b = 1; b < 11; b++)
                                                {
                                                    dr = dtGraph.NewRow();
                                                    dr["INITIAL"] = Initial;
                                                    nifty = (arrStrike[3] + b) - Initial;
                                                    final = arrStrike[3] + b;
                                                    dr["FINAL"] = final;

                                                    dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);

                                                    FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);
                                                    if (FinalValue < NegSpot)
                                                    {
                                                        dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                                    }
                                                    else
                                                    {
                                                        dr["PR"] = Math.Round(FinalValue, 2);
                                                    }

                                                    if (final < PutShortStrike1)
                                                    {
                                                        dtGraph.Rows.Add(dr);
                                                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                                                    }
                                                    if (final == arrStrike[2] && arrStrike[2] < arrStrike[0])
                                                    {
                                                        #region Main Row
                                                        dr = dtGraph.NewRow();
                                                        dr["INITIAL"] = Initial;
                                                        dr["FINAL"] = final;
                                                        nifty = final - Initial;
                                                        dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);
                                                        FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);
                                                        //dr["PR"] = FinalValue;
                                                        if (FinalValue < NegSpot)
                                                        {
                                                            dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                                        }
                                                        else
                                                        {
                                                            dr["PR"] = Math.Round(FinalValue, 2);
                                                        }

                                                        dtGraph.Rows.Add(dr);
                                                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                                                        #endregion

                                                        #region Next Row
                                                        dr = dtGraph.NewRow();
                                                        dr["INITIAL"] = Initial;
                                                        final = final + 0.01;
                                                        dr["FINAL"] = final;
                                                        nifty = final - Initial;
                                                        dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);

                                                        FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);

                                                        if (FinalValue < NegSpot)
                                                        {
                                                            dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                                        }
                                                        else
                                                        {
                                                            dr["PR"] = Math.Round(FinalValue, 2);
                                                        }
                                                        dtGraph.Rows.Add(dr);
                                                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });

                                                        #endregion

                                                        Final = Final + 10;
                                                    }
                                                }
                                                //Nifty = Nifty - 10;
                                            }
                                            #endregion

                                            #region Greater than 15
                                            else
                                            {
                                                int e = 0;
                                                e = Convert.ToInt32(StrikeDiffrencePut / 5);
                                                e = e + 1;
                                                bool flag1 = false;
                                                int addedValue1 = 0;
                                                for (int c = 0; c < e; c++)
                                                {
                                                    if (flag1)
                                                        addedValue1 = addedValue1 + 5;

                                                    dr = dtGraph.NewRow();
                                                    dr["INITIAL"] = Initial;
                                                    nifty = arrStrike[3] - Initial;
                                                    final = arrStrike[3] + addedValue1;
                                                    dr["FINAL"] = final;

                                                    nifty = nifty + addedValue1;
                                                    dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);

                                                    FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);
                                                    if (FinalValue < NegSpot)
                                                    {
                                                        dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                                    }
                                                    else
                                                    {
                                                        dr["PR"] = Math.Round(FinalValue, 2);
                                                    }

                                                    if (final < PutShortStrike1)
                                                    {
                                                        dtGraph.Rows.Add(dr);
                                                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                                                    }
                                                    if (final == arrStrike[2] && arrStrike[2] < arrStrike[0])
                                                    {
                                                        #region Main Row
                                                        dr = dtGraph.NewRow();
                                                        dr["INITIAL"] = Initial;
                                                        dr["FINAL"] = final;
                                                        nifty = final - Initial;
                                                        dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);
                                                        FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);
                                                        //dr["PR"] = FinalValue;
                                                        if (FinalValue < NegSpot)
                                                        {
                                                            dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                                        }
                                                        else
                                                        {
                                                            dr["PR"] = Math.Round(FinalValue, 2);
                                                        }

                                                        dtGraph.Rows.Add(dr);
                                                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                                                        #endregion

                                                        #region Next Row
                                                        dr = dtGraph.NewRow();
                                                        dr["INITIAL"] = Initial;
                                                        final = final + 0.01;
                                                        dr["FINAL"] = final;
                                                        nifty = final - Initial;
                                                        dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);

                                                        FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);

                                                        if (FinalValue < NegSpot)
                                                        {
                                                            dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                                        }
                                                        else
                                                        {
                                                            dr["PR"] = Math.Round(FinalValue, 2);
                                                        }
                                                        dtGraph.Rows.Add(dr);
                                                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });

                                                        #endregion

                                                        Final = Final + 10;
                                                    }
                                                    flag = true;
                                                }

                                            }
                                            #endregion
                                            strike4 = true;

                                            #endregion
                                        }
                                        flag = true;
                                    }
                                    Nifty = Nifty + 10;
                                }
                                #endregion
                            }
                            #endregion
                            else if (strike3)
                            {
                                #region Main Row
                                dr = dtGraph.NewRow();
                                dr["INITIAL"] = Initial;
                                dr["FINAL"] = (double)(Initial + ((100 * Nifty) / 100));
                                dr["NIFTY_PERFORMANCE"] = Nifty;
                                FinalValue = CalculateGoldenCushion(Nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);
                                //dr["PR"] = FinalValue;
                                if (FinalValue < NegSpot)
                                {
                                    dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                }
                                else
                                {
                                    dr["PR"] = Math.Round(FinalValue, 2);
                                }

                                dtGraph.Rows.Add(dr);
                                transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                                #endregion

                                #region Next Row
                                dr = dtGraph.NewRow();
                                dr["INITIAL"] = Initial;
                                final = (double)(Initial + ((100 * Nifty) / 100)) + 0.01;
                                dr["FINAL"] = final;
                                nifty = Nifty + 0.01; ;
                                dr["NIFTY_PERFORMANCE"] = nifty;

                                FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);

                                if (FinalValue < NegSpot)
                                {
                                    dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                }
                                else
                                {
                                    dr["PR"] = Math.Round(FinalValue, 2);
                                }
                                dtGraph.Rows.Add(dr);
                                transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });

                                #endregion
                            }
                            #region PutSpread Leg 2
                            else
                            {
                                #region Less Than 15
                                if (StrikeDiffrencePut < 15)
                                {
                                    for (int a = 0; a < 10; a++)
                                    {
                                        dr = dtGraph.NewRow();
                                        dr["INITIAL"] = Initial;
                                        final = (Initial + Nifty) + a;
                                        dr["FINAL"] = final;

                                        nifty = Nifty + a;
                                        dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);

                                        FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);
                                        if (FinalValue < NegSpot)
                                        {
                                            dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                        }
                                        else
                                        {
                                            dr["PR"] = Math.Round(FinalValue, 2);
                                        }
                                        if (final < PutSpreadStrike1)
                                        {
                                            dtGraph.Rows.Add(dr);
                                            transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                                        }
                                        if (final == arrStrike[1] && arrStrike[1] < arrStrike[0])
                                        {
                                            if (arrStrike[0] >= (final + a))
                                                a = a + Convert.ToInt32(arrStrike[0] - final);

                                            #region Strike 4
                                            #region Less than 15
                                            if (StrikeDiff < 15)
                                            {
                                                for (int b = 1; b < 10; b++)
                                                {
                                                    dr = dtGraph.NewRow();
                                                    dr["INITIAL"] = Initial;
                                                    nifty = (arrStrike[1] + b) - Initial;
                                                    final = arrStrike[1] + b;
                                                    dr["FINAL"] = final;

                                                    dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);

                                                    FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);
                                                    if (FinalValue < NegSpot)
                                                    {
                                                        dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                                    }
                                                    else
                                                    {
                                                        dr["PR"] = Math.Round(FinalValue, 2);
                                                    }

                                                    if (final < PutShortStrike1)
                                                    {
                                                        dtGraph.Rows.Add(dr);
                                                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                                                    }
                                                    if (final == arrStrike[0])
                                                    {
                                                        #region Main Row
                                                        dr = dtGraph.NewRow();
                                                        dr["INITIAL"] = Initial;
                                                        dr["FINAL"] = final;
                                                        nifty = final - Initial;
                                                        dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);
                                                        FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);
                                                        //dr["PR"] = FinalValue;
                                                        if (FinalValue < NegSpot)
                                                        {
                                                            dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                                        }
                                                        else
                                                        {
                                                            dr["PR"] = Math.Round(FinalValue, 2);
                                                        }

                                                        dtGraph.Rows.Add(dr);
                                                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                                                        #endregion

                                                        #region Next Row
                                                        dr = dtGraph.NewRow();
                                                        dr["INITIAL"] = Initial;
                                                        final = final + 0.01;
                                                        dr["FINAL"] = final;
                                                        nifty = final - Initial;
                                                        dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);

                                                        FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);

                                                        if (FinalValue < NegSpot)
                                                        {
                                                            dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                                        }
                                                        else
                                                        {
                                                            dr["PR"] = Math.Round(FinalValue, 2);
                                                        }
                                                        dtGraph.Rows.Add(dr);
                                                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });

                                                        #endregion

                                                        Final = Final + 10;
                                                    }
                                                }
                                                //Nifty = Nifty - 10;
                                            }
                                            #endregion

                                            #region Greater than 15
                                            else
                                            {
                                                int e = 0;
                                                e = Convert.ToInt32(StrikeDiff / 5);
                                                e = e + 1;
                                                bool flag1 = false;
                                                int addedValue1 = 0;
                                                for (int c = 0; c < e; c++)
                                                {
                                                    if (flag1)
                                                        addedValue1 = addedValue1 + 5;

                                                    dr = dtGraph.NewRow();
                                                    dr["INITIAL"] = Initial;
                                                    nifty = arrStrike[1] - Initial;
                                                    final = arrStrike[1] + addedValue1;
                                                    dr["FINAL"] = final;

                                                    nifty = nifty + addedValue1;
                                                    dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);

                                                    FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);
                                                    if (FinalValue < NegSpot)
                                                    {
                                                        dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                                    }
                                                    else
                                                    {
                                                        dr["PR"] = Math.Round(FinalValue, 2);
                                                    }

                                                    if (final < PutShortStrike1)
                                                    {
                                                        dtGraph.Rows.Add(dr);
                                                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                                                    }
                                                    if (final == arrStrike[0])
                                                    {
                                                        #region Main Row
                                                        dr = dtGraph.NewRow();
                                                        dr["INITIAL"] = Initial;
                                                        dr["FINAL"] = final;
                                                        nifty = final - Initial;
                                                        dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);
                                                        FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);
                                                        //dr["PR"] = FinalValue;
                                                        if (FinalValue < NegSpot)
                                                        {
                                                            dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                                        }
                                                        else
                                                        {
                                                            dr["PR"] = Math.Round(FinalValue, 2);
                                                        }

                                                        dtGraph.Rows.Add(dr);
                                                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                                                        #endregion

                                                        #region Next Row
                                                        dr = dtGraph.NewRow();
                                                        dr["INITIAL"] = Initial;
                                                        final = final + 0.01;
                                                        dr["FINAL"] = final;
                                                        nifty = final - Initial;
                                                        dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);

                                                        FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);

                                                        if (FinalValue < NegSpot)
                                                        {
                                                            dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                                        }
                                                        else
                                                        {
                                                            dr["PR"] = Math.Round(FinalValue, 2);
                                                        }
                                                        dtGraph.Rows.Add(dr);
                                                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });

                                                        #endregion

                                                        Final = Final + 10;
                                                    }
                                                }

                                            }
                                            #endregion
                                            strike4 = true;

                                            #endregion
                                        }
                                    }


                                }
                                #endregion

                                #region Greater than 15
                                else
                                {
                                    int d = 0;
                                    d = Convert.ToInt32(StrikeDiffrencePut / 5);
                                    d = d + 1;
                                    bool flag = false;
                                    int addedValue = 0;
                                    for (int a = 0; a < d; a++)
                                    {
                                        if (flag)
                                            addedValue = addedValue + 5;

                                        dr = dtGraph.NewRow();
                                        dr["INITIAL"] = Initial;
                                        final = (Initial + Nifty) + addedValue;
                                        dr["FINAL"] = final;

                                        nifty = Nifty + addedValue;
                                        dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);

                                        FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);
                                        if (FinalValue < NegSpot)
                                        {
                                            dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                        }
                                        else
                                        {
                                            dr["PR"] = Math.Round(FinalValue, 2);
                                        }
                                        if (final < PutShortStrike1)
                                        {
                                            dtGraph.Rows.Add(dr);
                                            transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                                        }
                                        if (final == arrStrike[1] && arrStrike[1] < arrStrike[0])
                                        {
                                            if (arrStrike[0] >= (final + addedValue))
                                                addedValue = addedValue + Convert.ToInt32(arrStrike[0] - final);

                                            #region Strike 4
                                            #region Less than 15
                                            if (StrikeDiff < 15)
                                            {
                                                for (int b = 1; b < 10; b++)
                                                {
                                                    dr = dtGraph.NewRow();
                                                    dr["INITIAL"] = Initial;
                                                    nifty = (arrStrike[1] + b) - Initial;
                                                    final = arrStrike[1] + b;
                                                    dr["FINAL"] = final;

                                                    dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);

                                                    FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);
                                                    if (FinalValue < NegSpot)
                                                    {
                                                        dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                                    }
                                                    else
                                                    {
                                                        dr["PR"] = Math.Round(FinalValue, 2);
                                                    }

                                                    if (final < PutSpreadStrike1)
                                                    {
                                                        dtGraph.Rows.Add(dr);
                                                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                                                    }
                                                    if (final == arrStrike[2])
                                                    {
                                                        #region Main Row
                                                        dr = dtGraph.NewRow();
                                                        dr["INITIAL"] = Initial;
                                                        dr["FINAL"] = final;
                                                        nifty = final - Initial;
                                                        dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);
                                                        FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);
                                                        //dr["PR"] = FinalValue;
                                                        if (FinalValue < NegSpot)
                                                        {
                                                            dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                                        }
                                                        else
                                                        {
                                                            dr["PR"] = Math.Round(FinalValue, 2);
                                                        }

                                                        dtGraph.Rows.Add(dr);
                                                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                                                        #endregion

                                                        #region Next Row
                                                        dr = dtGraph.NewRow();
                                                        dr["INITIAL"] = Initial;
                                                        final = final + 0.01;
                                                        dr["FINAL"] = final;
                                                        nifty = final - Initial;
                                                        dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);

                                                        FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);

                                                        if (FinalValue < NegSpot)
                                                        {
                                                            dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                                        }
                                                        else
                                                        {
                                                            dr["PR"] = Math.Round(FinalValue, 2);
                                                        }
                                                        dtGraph.Rows.Add(dr);
                                                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });

                                                        #endregion

                                                        Final = Final + 10;
                                                    }
                                                    if (final == arrStrike[0])
                                                    {
                                                        #region Main Row
                                                        dr = dtGraph.NewRow();
                                                        dr["INITIAL"] = Initial;
                                                        dr["FINAL"] = final;
                                                        nifty = final - Initial;
                                                        dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);
                                                        FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);
                                                        //dr["PR"] = FinalValue;
                                                        if (FinalValue < NegSpot)
                                                        {
                                                            dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                                        }
                                                        else
                                                        {
                                                            dr["PR"] = Math.Round(FinalValue, 2);
                                                        }

                                                        dtGraph.Rows.Add(dr);
                                                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                                                        #endregion

                                                        #region Next Row
                                                        dr = dtGraph.NewRow();
                                                        dr["INITIAL"] = Initial;
                                                        final = final + 0.01;
                                                        dr["FINAL"] = final;
                                                        nifty = final - Initial;
                                                        dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);

                                                        FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);

                                                        if (FinalValue < NegSpot)
                                                        {
                                                            dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                                        }
                                                        else
                                                        {
                                                            dr["PR"] = Math.Round(FinalValue, 2);
                                                        }
                                                        dtGraph.Rows.Add(dr);
                                                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });

                                                        #endregion

                                                        Final = Final + 10;
                                                    }
                                                }
                                                //Nifty = Nifty - 10;
                                            }
                                            #endregion

                                            #region Greater than 15
                                            else
                                            {
                                                int e = 0;
                                                e = Convert.ToInt32(StrikeDiff / 5);
                                                e = e + 1;
                                                bool flag1 = false;
                                                int addedValue1 = 0;
                                                for (int c = 0; c < e; c++)
                                                {

                                                    addedValue1 = addedValue1 + 5;

                                                    dr = dtGraph.NewRow();
                                                    dr["INITIAL"] = Initial;
                                                    nifty = arrStrike[1] - Initial;
                                                    final = arrStrike[1] + addedValue1;
                                                    dr["FINAL"] = final;

                                                    nifty = nifty + addedValue1;
                                                    dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);

                                                    FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);
                                                    if (FinalValue < NegSpot)
                                                    {
                                                        dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                                    }
                                                    else
                                                    {
                                                        dr["PR"] = Math.Round(FinalValue, 2);
                                                    }

                                                    if (final < PutSpreadStrike1)
                                                    {
                                                        dtGraph.Rows.Add(dr);
                                                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                                                    }
                                                    if (final == arrStrike[2])
                                                    {

                                                        #region Next Row
                                                        dr = dtGraph.NewRow();
                                                        dr["INITIAL"] = Initial;
                                                        final = final + 0.01;
                                                        dr["FINAL"] = final;
                                                        nifty = final - Initial;
                                                        dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);

                                                        FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);

                                                        if (FinalValue < NegSpot)
                                                        {
                                                            dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                                        }
                                                        else
                                                        {
                                                            dr["PR"] = Math.Round(FinalValue, 2);
                                                        }
                                                        dtGraph.Rows.Add(dr);
                                                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });

                                                        #endregion

                                                        Final = Final + 10;
                                                    }
                                                    if (final == arrStrike[0])
                                                    {
                                                        #region Main Row
                                                        dr = dtGraph.NewRow();
                                                        dr["INITIAL"] = Initial;
                                                        dr["FINAL"] = final;
                                                        nifty = final - Initial;
                                                        dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);
                                                        FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);
                                                        //dr["PR"] = FinalValue;
                                                        if (FinalValue < NegSpot)
                                                        {
                                                            dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                                        }
                                                        else
                                                        {
                                                            dr["PR"] = Math.Round(FinalValue, 2);
                                                        }

                                                        dtGraph.Rows.Add(dr);
                                                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                                                        #endregion

                                                        #region Next Row
                                                        dr = dtGraph.NewRow();
                                                        dr["INITIAL"] = Initial;
                                                        final = final + 0.01;
                                                        dr["FINAL"] = final;
                                                        nifty = final - Initial;
                                                        dr["NIFTY_PERFORMANCE"] = Math.Round(nifty, 2);

                                                        FinalValue = CalculateGoldenCushion(nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);

                                                        if (FinalValue < NegSpot)
                                                        {
                                                            dr["PR"] = Math.Round(Convert.ToDouble(NegSpot), 2);
                                                        }
                                                        else
                                                        {
                                                            dr["PR"] = Math.Round(FinalValue, 2);
                                                        }
                                                        dtGraph.Rows.Add(dr);
                                                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });

                                                        #endregion

                                                        Final = Final + 10;
                                                    }
                                                    flag1 = true;
                                                }
                                                //var n = Convert.ToInt32(final - Initial);
                                                //Nifty = n - 20;
                                            }
                                            #endregion
                                            strike4 = true;

                                            #endregion

                                        }
                                        flag = true;
                                    }
                                }
                                #endregion
                            }
                            #endregion

                            strike = true;
                        }
                        #endregion
                    }
                    else
                    {
                        dr["INITIAL"] = Initial;
                        dr["FINAL"] = Final;
                        dr["NIFTY_PERFORMANCE"] = Nifty;

                        FinalValue = CalculateGoldenCushion(Nifty, PutSpreadStrike1, PutSpreadStrike2, PutShortStrike1, PutShortStrike2, LowerCoupon, FixedCoupon, PutSpreadPR, PutPR, PutSpreadOptionType, PutOptionType);
                        if (FinalValue < NegSpot)
                        {
                            dr["PR"] = NegSpot;
                        }
                        else
                        {
                            dr["PR"] = Math.Round(FinalValue, 2);
                        }

                        dtGraph.Rows.Add(dr);
                        transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["PR"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });

                        strike = true;

                    }

                    //if (final == 0)
                    //    final = Final;
                    PrevFinal = Convert.ToInt32(Final);
                }

                return transactionCounts;
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];

                ClearGoldenCushionSession();

                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "GenerateGraphGoldenCushion", objUserMaster.UserID);
                return transactionCounts;
            }
        }

        public double CalculateGoldenCushion(double UnderlyingPerformance, double PutSpreadStrike1, double PutSpreadStrike2, double PutShortStrike1, double PutShortStrike2, double LowerCoupon, double FixedCoupon, double PutSpreadPR, double PutPR, string PutSpreadOptionType, string PutOptionType)
        {
            double dblPR = 0;
            double dblPR1 = 0;
            double dblPR2 = 0;
            double dblSpot = 100;

            if (PutSpreadOptionType == "Put Spread Short")
            {
                if (PutSpreadStrike2 < PutSpreadStrike1 && PutSpreadStrike1 >= dblSpot)
                {
                    if (PutSpreadStrike1 > dblSpot)
                        dblPR1 = Math.Max((LowerCoupon / 100), Math.Min((FixedCoupon / 100), (FixedCoupon / 100) + PutSpreadPR * ((UnderlyingPerformance / 100) + (dblSpot - PutSpreadStrike1) / 100)));
                    //dblPR1 = Math.Max(LowerCoupon / 100, Math.Min(FixedCoupon / 100, (FixedCoupon / 100) + PutSpreadPR * UnderlyingPerformance / 100 + (dblSpot - PutSpreadStrike1) / 100));
                    else if (PutSpreadStrike1 == dblSpot)
                        dblPR1 = Math.Max((LowerCoupon / 100), Math.Min((FixedCoupon / 100), (FixedCoupon / 100) + PutSpreadPR * (UnderlyingPerformance / 100)));
                }
                else if (PutSpreadStrike2 < PutSpreadStrike1 && PutSpreadStrike1 < dblSpot)
                    dblPR1 = Math.Max((LowerCoupon / 100), Math.Min((FixedCoupon / 100), (FixedCoupon / 100) + PutSpreadPR * ((UnderlyingPerformance / 100) + (dblSpot - PutSpreadStrike1) / 100)));
            }

            if (PutOptionType == "Put Spread Short")
            {
                if (PutShortStrike2 < dblSpot && PutShortStrike2 < PutShortStrike1 && PutShortStrike1 < dblSpot)
                {
                    dblPR2 = Math.Min(0, PutPR * Math.Max(((PutShortStrike2 - PutShortStrike1) / 100), ((UnderlyingPerformance / 100) + (dblSpot - PutShortStrike1) / 100)));
                }
                else if (PutShortStrike2 < dblSpot && PutShortStrike2 < PutShortStrike1 && PutShortStrike1 == dblSpot)
                {
                    dblPR2 = Math.Min(0, PutPR * Math.Max(((PutShortStrike2 - dblSpot) / 100), (UnderlyingPerformance / 100)));
                }
            }
            else
            {
                if (PutShortStrike2 == 0 && PutShortStrike1 < dblSpot)
                {
                    dblPR2 = Math.Min(0, PutPR * ((UnderlyingPerformance / 100) + (dblSpot - PutShortStrike1) / 100));
                    //+Min (0, ' + PutParticipatoryRatio + ' * (Underlying performance + (' + (Spot - Strike1) + ')% ))';
                }
                else if (PutShortStrike2 == 0)
                {
                    dblPR2 = Math.Min(0, PutPR * UnderlyingPerformance / 100);
                    //+Min (0,' + PutParticipatoryRatio + ' * Underlying Performance)';
                }
            }

            dblPR = dblPR1 + dblPR2;

            return dblPR * 100;
        }
        #endregion

        #region Call Binary
        [HttpGet]
        public ActionResult CallBinary(string ProductID, string GenerateGraph, bool IsQuotron = false)
        {
            LoginController objLoginController = new LoginController();
            try
            {
                if (ValidateSession())
                {
                    CallBinary objCallBinary = new CallBinary();

                    #region Menu Access By on Role

                    Int32 intResult = 0;
                    // bool PPorNonPP = false;

                    UserMaster objUserMaster = new UserMaster();
                    objUserMaster = (UserMaster)Session["LoggedInUser"];

                    var Result = objSP_PRICINGEntities.VALIDATE_MENU_BY_ROLE(objUserMaster.RoleID, "BCB");
                    intResult = Convert.ToInt32(Result.SingleOrDefault());

                    if (intResult == 0)
                        return RedirectToAction("UserNotAuthorize", "Login");

                    #endregion

                    #region Bind Underlying List
                    List<Underlying> UnderlyingList = new List<Underlying>();

                    ObjectResult<UnderlyingListResult> objUnderlyingListResult;
                    List<UnderlyingListResult> UnderlyingListResultList = new List<UnderlyingListResult>();

                    objUnderlyingListResult = objSP_PRICINGEntities.SP_FETCH_UNDERLYING_DETAILS();
                    UnderlyingListResultList = objUnderlyingListResult.ToList();

                    foreach (UnderlyingListResult oUnderlyingListResult in UnderlyingListResultList)
                    {
                        Underlying objUnderlying = new Underlying();
                        General.ReflectSingleData(objUnderlying, oUnderlyingListResult);

                        UnderlyingList.Add(objUnderlying);
                    }
                    objCallBinary.UnderlyingList = UnderlyingList;

                    //--Set default underlying--Added by Shweta on 3rd May 2016------------START--------------------
                    string strDefaultUnderlying = System.Configuration.ConfigurationManager.AppSettings["DefaultUnderlying"].ToUpper();
                    Underlying objDefaulyUnderlying = objCallBinary.UnderlyingList.Find(delegate(Underlying oUnderlying) { return oUnderlying.UnderlyingShortName == strDefaultUnderlying; });
                    objCallBinary.UnderlyingID = objDefaulyUnderlying.UnderlyingID;

                    objCallBinary.EntityID = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["DefaultEntityID"]);
                    objCallBinary.IsSecuredID = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["DefaultIsSecuredID"]);
                    //--Set default underlying--Added by Shweta on 3rd May 2016------------END----------------------
                    #endregion

                    if (ProductID != "" && ProductID != null)
                    {
                        ObjectResult<CallBinaryEditResult> objCallBinaryEditResult = objSP_PRICINGEntities.FETCH_CALL_BINARY_EDIT_DETAILS(ProductID);
                        List<CallBinaryEditResult> CallBinaryEditResultList = objCallBinaryEditResult.ToList();

                        General.ReflectSingleData(objCallBinary, CallBinaryEditResultList[0]);

                        DataSet dsResult = new DataSet();
                        dsResult = General.ExecuteDataSet("FETCH_UNDERLYING_CREATION_BYID", objCallBinary.UnderlyingID);

                        if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                        {
                            ViewBag.UnderlyingShortName = Convert.ToString(dsResult.Tables[0].Rows[0]["UNDERLYING_SHORTNAME"]);
                        }
                    }

                    if (GenerateGraph == "GenerateGraph")
                    {
                        objCallBinary = (CallBinary)TempData["CallBinaryGraph"];
                        ObjectResult<CallBinaryEditResult> objCallBinaryEditResult = objSP_PRICINGEntities.FETCH_CALL_BINARY_EDIT_DETAILS(objCallBinary.ProductID);
                        List<CallBinaryEditResult> CallBinaryEditResultList = objCallBinaryEditResult.ToList();
                        CallBinary oCallBinary = new CallBinary();
                        General.ReflectSingleData(oCallBinary, CallBinaryEditResultList[0]);

                        objCallBinary.Status = oCallBinary.Status;
                        //objCallBinary.SaveStatus = oCallBinary.SaveStatus;
                        return GenerateCallBinaryGraph(objCallBinary);
                    }

                    else if (Session["CallBinaryCopyQuote"] != null)
                    {
                        ViewBag.Message = "Quote is copied";
                        objCallBinary = (CallBinary)Session["CallBinaryCopyQuote"];
                        objCallBinary.UnderlyingID = Convert.ToInt32(Session["UnderlyingID"]);

                        ObjectResult<CallBinaryEditResult> objCallBinaryEditResult = objSP_PRICINGEntities.FETCH_CALL_BINARY_EDIT_DETAILS("");
                        List<CallBinaryEditResult> CallBinaryEditResultList = objCallBinaryEditResult.ToList();
                        CallBinary oCallBinary = new CallBinary();
                        if (CallBinaryEditResultList != null && CallBinaryEditResultList.Count > 0)
                            General.ReflectSingleData(oCallBinary, CallBinaryEditResultList[0]);

                        objCallBinary.ParentProductID = objCallBinary.ProductID;
                        objCallBinary.ProductID = "";
                        objCallBinary.Status = oCallBinary.Status;
                        objCallBinary.SaveStatus = oCallBinary.SaveStatus;
                        objCallBinary.IsCopyQuote = true;

                        //Added by Shweta on 10th May---------------START-------
                        objCallBinary.CallCustomIV1 = 0;
                        objCallBinary.CallCustomIV2 = 0;
                        objCallBinary.CallCustomRF1 = 0;
                        objCallBinary.CallCustomRF2 = 0;
                        //Added by Shweta on 10th May---------------END---------

                        //-------------Added by Shweta on 22nd July 2016 to Fetch Latest Deployment Rate------------START--------
                        string strDeploymentRate = "";
                        var DeploymentRate = objSP_PRICINGEntities.SP_FETCH_PRICING_DEPLOYMENT_RATE(Convert.ToInt32(objCallBinary.RedemptionPeriodDays), objCallBinary.EntityID, objCallBinary.IsSecuredID);
                        strDeploymentRate = Convert.ToString(DeploymentRate.SingleOrDefault());
                        objCallBinary.DeploymentRate = Convert.ToDouble(strDeploymentRate);
                        //-------------Added by Shweta on 22nd July 2016 to Fetch Latest Deployment Rate------------END----------
                    }

                    else if (Session["CallBinaryChildQuote"] != null)
                    {
                        ViewBag.Message = true;
                        objCallBinary = (CallBinary)Session["CallBinaryChildQuote"];
                        objCallBinary.UnderlyingID = Convert.ToInt32(Session["UnderlyingID"]);

                        ObjectResult<CallBinaryEditResult> objCallBinaryEditResult = objSP_PRICINGEntities.FETCH_CALL_BINARY_EDIT_DETAILS("");
                        List<CallBinaryEditResult> CallBinaryEditResultList = objCallBinaryEditResult.ToList();
                        CallBinary oCallBinary = new CallBinary();
                        if (CallBinaryEditResultList != null && CallBinaryEditResultList.Count > 0)
                            General.ReflectSingleData(oCallBinary, CallBinaryEditResultList[0]);

                        objCallBinary.ParentProductID = objCallBinary.ProductID;
                        objCallBinary.ProductID = "";
                        objCallBinary.Status = oCallBinary.Status;
                        objCallBinary.SaveStatus = oCallBinary.SaveStatus;
                        objCallBinary.IsChildQuote = true;
                    }
                    else if (Session["CancelQuote"] != null)
                    {
                        objCallBinary = (CallBinary)Session["CancelQuote"];

                        ObjectResult<CallBinaryEditResult> objCallBinaryEditResult = objSP_PRICINGEntities.FETCH_CALL_BINARY_EDIT_DETAILS(objCallBinary.ProductID);
                        List<CallBinaryEditResult> CallBinaryEditResultList = objCallBinaryEditResult.ToList();
                        CallBinary oCallBinary = new CallBinary();
                        if (CallBinaryEditResultList != null && CallBinaryEditResultList.Count > 0)
                            General.ReflectSingleData(oCallBinary, CallBinaryEditResultList[0]);

                        objCallBinary.Status = oCallBinary.Status;
                        objCallBinary.SaveStatus = oCallBinary.SaveStatus;

                        Session.Remove("CancelQuote");
                    }
                    else
                    {
                        Session.Remove("IsChildQuoteCallBinary");
                        Session.Remove("ParentProductID");
                        Session.Remove("UnderlyingID");
                    }

                    if (IsQuotron == true)
                    {
                        objCallBinary.IsQuotron = true;
                    }

                    if (Session["CallBinaryChildQuote"] == null && Session["CallBinaryCopyQuote"] == null)
                        objCallBinary.SaveStatus = "";

                    if (Session["CallBinaryCopyQuote"] != null)
                        Session.Remove("CallBinaryCopyQuote");

                    if (Session["CallBinaryChildQuote"] != null)
                        Session.Remove("CallBinaryChildQuote");

                    if (ProductID == null)
                    {
                        objCallBinary.isGraphActive = false;
                        return View(objCallBinary);
                    }
                    else
                    {
                        return GenerateCallBinaryGraph(objCallBinary);
                    }
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

                ClearCallBinarySession();

                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "CallBinary Get", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        private ActionResult GenerateCallBinaryGraph(CallBinary objCallBinary)
        {
            LoginController objLoginController = new LoginController();
            try
            {
                if (ValidateSession())
                {
                    int[] _array = { 3, 5, 7, 9, 11, 13 };
                    int index = Array.IndexOf<int>(_array, 7);

                    double[] dblArray = { 100, 110, 120, 130 };
                    Int32 dblStrikeIndex = Array.IndexOf<int>(_array, 7);

                    var transactionCounts = new List<Graph>();
                    transactionCounts = GenerateFixedMLDGraphCalculation(objCallBinary.CallStrike1, objCallBinary.FixedCoupon, objCallBinary.MaxCoupon, objCallBinary.RedemptionPeriodDays);

                    //CallBinary obj = new CallBinary();

                    #region Pie Chart For Call Binary
                    var xDataMonths = transactionCounts.Select(i => i.Column1).ToArray();
                    var yDataCounts = transactionCounts.Select(i => new object[] { i.Column2 }).ToArray();
                    var yDataCounts1 = transactionCounts.Select(i => new object[] { i.Column3 }).ToArray();



                    var chart = new Highcharts("pie")
                        //define the type of chart 
                                .InitChart(new Chart { DefaultSeriesType = ChartTypes.Line })
                        //overall Title of the chart 
                                .SetTitle(new Title { Text = "Call Binary" })
                        ////small label below the main Title
                        //        .SetSubtitle(new Subtitle { Text = "Accounting" })
                        //load the X values
                                .SetXAxis(new XAxis { Title = new XAxisTitle { Text = "Underlying Returns" }, Categories = xDataMonths, Labels = new XAxisLabels { Step = 2 } })
                        //set the Y title
                                .SetYAxis(new YAxis { Title = new YAxisTitle { Text = "Product Returns" } })
                                .SetTooltip(new Tooltip
                                {
                                    Enabled = true,
                                    Formatter = @"function() { return '<b>'+ this.series.name +'</b><br/>'+ this.x +': '+ this.y; }"
                                })
                                .SetPlotOptions(new PlotOptions
                                {
                                    Line = new PlotOptionsLine
                                    {
                                        DataLabels = new PlotOptionsLineDataLabels
                                        {
                                            Enabled = false
                                        },
                                        EnableMouseTracking = true
                                    }
                                })
                        //load the Y values 
                                .SetSeries(new[]
                    {
                        new Series {Name = "Coupon", Data = new Data(yDataCounts)},
                            //you can add more y data to create a second line
                            // new Series { Name = "Strike", Data = new Data(yDataCounts1) }
                    });
                    #endregion

                    if (Session["CallBinaryCopyQuote"] != null)
                    {
                        ViewBag.Message = "Quote is copied";
                        objCallBinary = (CallBinary)Session["CallBinaryCopyQuote"];
                        Session.Remove("CallBinaryCopyQuote");
                    }

                    objCallBinary.CallBinaryChart = chart;

                    return View(objCallBinary);
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

                ClearCallBinarySession();

                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "GenerateCallBinaryGraph", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        [HttpPost, ValidateInput(false)]
        public ActionResult CallBinary(string Command, CallBinary objCallBinary, FormCollection objFormCollection)
        {
            LoginController objLoginController = new LoginController();
            try
            {
                if (ValidateSession())
                {
                    #region Bind Underlying List
                    List<Underlying> UnderlyingList = new List<Underlying>();

                    ObjectResult<UnderlyingListResult> objUnderlyingListResult;
                    List<UnderlyingListResult> UnderlyingListResultList = new List<UnderlyingListResult>();

                    objUnderlyingListResult = objSP_PRICINGEntities.SP_FETCH_UNDERLYING_DETAILS();
                    UnderlyingListResultList = objUnderlyingListResult.ToList();

                    foreach (UnderlyingListResult oUnderlyingListResult in UnderlyingListResultList)
                    {
                        Underlying objUnderlying = new Underlying();
                        General.ReflectSingleData(objUnderlying, oUnderlyingListResult);

                        UnderlyingList.Add(objUnderlying);
                    }
                    objCallBinary.UnderlyingList = UnderlyingList;
                    #endregion

                    if (Command == "ExportToExcel")
                    {
                        ExportCallBinary(objCallBinary);

                        return RedirectToAction("CallBinary");
                    }
                    else if (Command == "ExportCallStrike1Grid")
                    {
                        string CallStrike1Summary = objFormCollection["ExportCallStrike1Summary"].Replace("<html><head></head><body class=\"table-responsive col-md-12\">", "");
                        CallStrike1Summary = CallStrike1Summary.Replace("</body></html>", "");

                        string CallStrike2Summary = objFormCollection["ExportCallStrike2Summary"].Replace("<html><head></head><body class=\"table-responsive col-md-12\">", "");
                        CallStrike2Summary = CallStrike2Summary.Replace("</body></html>", "");

                        string StrikeHTML = CallStrike1Summary + "<br />" + CallStrike2Summary;

                        ExportStrikeGrid(StrikeHTML);

                        return RedirectToAction("CallBinary");
                    }
                    else if (Command == "ExportCallStrike2Grid")
                    {
                        string StrikeHTML = objFormCollection["ExportCallStrike2Summary"];

                        ExportStrikeGrid(StrikeHTML);

                        return RedirectToAction("FixedOrPR");
                    }
                    else if (Command == "CopyQuote")
                    {
                        Session["CallBinaryCopyQuote"] = objCallBinary;
                        Session["UnderlyingID"] = objCallBinary.UnderlyingID;

                        return RedirectToAction("CallBinary");
                    }
                    else if (Command == "CreateChildQuote")
                    {
                        Session.Remove("ParentProductID");
                        Session.Remove("IsChildQuoteCallBinary");
                        Session.Remove("UnderlyingID");

                        Session["ParentProductID"] = objCallBinary.ProductID;
                        Session["UnderlyingID"] = objCallBinary.UnderlyingID;

                        objCallBinary.IsChildQuote = true;

                        Session["CallBinaryChildQuote"] = objCallBinary;
                        Session["IsChildQuoteCallBinary"] = objCallBinary.IsChildQuote;

                        return RedirectToAction("CallBinary");
                    }
                    else if (Command == "GenerateGraph")
                    {
                        objCallBinary.isGraphActive = true;
                        TempData["CallBinaryGraph"] = objCallBinary;
                        return RedirectToAction("CallBinary", new { GenerateGraph = "GenerateGraph" });
                    }
                    else if (Command == "AddNewProduct")
                    {
                        var productID = objCallBinary.ProductID;
                        UserMaster objUserMaster = new UserMaster();
                        objUserMaster = (UserMaster)Session["LoggedInUser"];

                        EncryptDecrypt obj = new EncryptDecrypt();
                        var encryptedpaswd = obj.Encrypt(objUserMaster.Password, "SPPricing", CryptographyEngine.AlgorithmType.DES);
                        var ProductType = "PP";

                        var Url = "http://edemumnewuatvm4:63400/Login.aspx?UserId=" + objUserMaster.LoginName + "&Key=" + encryptedpaswd + "&ProductId=" + productID + "&ProductType=" + ProductType;
                        return Redirect(Url);
                    }
                    else if (Command == "Cancel")
                    {
                        Session["CancelQuote"] = objCallBinary;

                        return RedirectToAction("CallBinary");
                    }
                    else if (Command == "PricingInExcel")
                    {
                        objCallBinary.IsWorkingFileExport = OpenWorkingExcelFile("CB", objCallBinary.ProductID);

                        if (!objCallBinary.IsWorkingFileExport)
                            objCallBinary.WorkingFileStatus = "File Not Found";

                        return View(objCallBinary);
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

                ClearCallBinarySession();

                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "CallBinary Post", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        public JsonResult ExportCallBinaryWorkingFile(string ProductID, string Distributor, string EdelweissBuiltIn, string DistributorBuiltIn, string BuiltInAdjustment, string TotalBuiltIn, string FixedCoupon, string IRR, string IsIRR, string MaxCoupon,
            string MaxCouponIRR, string IsMaxCouponIRR, string DeploymentRate, string CustomerDeploymentRate, string Remaining, string Underlying, string TotalOptionPrice, string NetRemaining,
            string OptionTenureMonth, string RedemptionPeriodMonth, string IsRedemptionPeriodMonth, string RedemptionPeriodDays, string InitialAveragingMonth, string InitialAveragingDaysDiff, string FinalAveragingMonth,
            string FinalAveragingDaysDiff, string CallBinaryLong, string CallUnderlying, string CallStrike1, string CallStrike2, string CallCouponRise, string CallPrice, string CallDiscountedPrice,
            string CallPRAdjustedPrice, string CallIV1, string CallCustomIV1, string CallRF1, string CallCustomRF1, string CallIV2, string CallCustomIV2, string CallRF2,
            string CallCustomRF2, string SalesComments, string TradingComments, string CouponScenario,
            string CallStrike1Summary, string CallStrike2Summary, string ExportCallStrike1Summary, string ExportCallStrike2Summary, string Entity, string IsSecured)
        {
            try
            {
                string strTemplateFilePath = Server.MapPath("~/Templates");
                string strTemplateFileName = strTemplateFilePath + "\\CallBinaryTemplateWorkingFile.xlsx";

                string strTargetFilePath = System.Configuration.ConfigurationManager.AppSettings["WorkingFilePath"];
                string strTargetFileName = strTargetFilePath + "\\" + ProductID + "_CallBinary.xlsx";

                string Role = Convert.ToString(Session["Role"]);

                if (System.IO.File.Exists(strTargetFileName))
                    System.IO.File.Delete(strTargetFileName);

                FileInfo objTemplateFileInfo = new FileInfo(strTemplateFileName);
                objTemplateFileInfo.CopyTo(strTargetFileName);

                FileInfo objTargetFileInfo = new FileInfo(strTargetFileName);

                using (var xlPackage = new ExcelPackage(objTargetFileInfo))
                {
                    var worksheet = xlPackage.Workbook.Worksheets["CallBinary"];

                    worksheet.Cell(1, 2).Value = ProductID.ToString();
                    worksheet.Cell(1, 4).Value = Distributor.ToString().ToUpper();
                    worksheet.Cell(1, 6).Value = Underlying;

                    worksheet.Cell(2, 2).Formula = "=" + EdelweissBuiltIn.ToString() + "%";
                    worksheet.Cell(2, 4).Formula = "=" + DistributorBuiltIn.ToString() + "%";
                    worksheet.Cell(2, 6).Formula = "=B2+D2+H2";
                    worksheet.Cell(2, 8).Formula = "=" + BuiltInAdjustment.ToString() + "%";

                    if (IsIRR.ToUpper() == "TRUE")
                    {
                        worksheet.Cell(3, 2).Formula = "=" + IRR.ToString() + "%";
                        worksheet.Cell(3, 4).Formula = "=(POWER((1 + B3), (F4 / 365)) - 1)*100%";
                    }
                    else
                    {
                        worksheet.Cell(3, 2).Formula = "=((POWER((1+D3),(12/D4))-1) * 100) %";
                        worksheet.Cell(3, 4).Formula = "=" + FixedCoupon.ToString() + "%";
                    }
                    if (IsMaxCouponIRR.ToUpper() == "TRUE")
                    {
                        worksheet.Cell(3, 6).Formula = "=" + MaxCouponIRR.ToString() + "%";
                        worksheet.Cell(3, 8).Formula = "=(POWER((1 + F3), (F4 / 365)) - 1)*100%";
                    }
                    else
                    {
                        worksheet.Cell(3, 6).Formula = "=((POWER((1+H3),(12/D4))-1) * 100) %";
                        worksheet.Cell(3, 8).Formula = "=" + MaxCoupon.ToString() + "%";
                    }

                    worksheet.Cell(4, 2).Value = OptionTenureMonth.ToString();
                    if (IsRedemptionPeriodMonth.ToUpper() == "TRUE")
                    {
                        worksheet.Cell(4, 4).Formula = RedemptionPeriodMonth;
                        worksheet.Cell(4, 6).Formula = "=ROUND(D4*30.417, 0)";
                    }
                    else
                    {
                        worksheet.Cell(4, 4).Formula = "=ROUND(F4/30.417,2)";
                        worksheet.Cell(4, 6).Formula = RedemptionPeriodDays.ToString();
                    }

                    #region Get Entity Name
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
                    LookupMaster objLookupMasterEntity = EntityList.Find(delegate(LookupMaster oLookupMaster) { return oLookupMaster.LookupID == Convert.ToInt32(Entity); });
                    worksheet.Cell(5, 2).Value = objLookupMasterEntity.LookupDescription;
                    #endregion

                    #region Get Is Secured
                    objLookupResult = null;
                    LookupResultList = null;
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
                    LookupMaster objLookupMasterIsSecured = IsSecuredList.Find(delegate(LookupMaster oLookupMaster) { return oLookupMaster.LookupID == Convert.ToInt32(IsSecured); });
                    worksheet.Cell(5, 4).Value = objLookupMasterIsSecured.LookupDescription;
                    #endregion

                    worksheet.Cell(5, 6).Formula = "=" + DeploymentRate.ToString() + "%";

                    if (CustomerDeploymentRate == "")
                        CustomerDeploymentRate = "0";
                    worksheet.Cell(5, 8).Formula = "=" + CustomerDeploymentRate.ToString() + "%";

                    worksheet.Cell(6, 2).Formula = "=(100-(B2+D2)*100)-(100*(1+ROUND(D3,4)))/(POWER((1+(0.11)),(36.16/12)))";
                    worksheet.Cell(6, 4).Formula = "=G11";
                    worksheet.Cell(6, 6).Formula = "=B6 + D6";

                    worksheet.Cell(8, 2).Value = InitialAveragingMonth.ToString();
                    worksheet.Cell(8, 4).Value = InitialAveragingDaysDiff.ToString();
                    worksheet.Cell(8, 6).Value = FinalAveragingMonth.ToString();
                    worksheet.Cell(8, 8).Value = FinalAveragingDaysDiff.ToString();

                    worksheet.Cell(11, 1).Value = "Call Binary Long";
                    worksheet.Cell(11, 2).Value = Underlying;
                    worksheet.Cell(11, 3).Value = CallStrike1.ToString();
                    worksheet.Cell(11, 4).Formula = "=H3-D3";
                    worksheet.Cell(11, 5).Formula = "=(AVERAGE(INDIRECT(\"$Q$2:\"&ADDRESS(1+$B$8,16+$F$8))))-(AVERAGE(INDIRECT(\"$Y$2:\"&ADDRESS(1+$B$8,24+$F$8))))";
                    worksheet.Cell(11, 6).Formula = "=E11*-1";
                    worksheet.Cell(11, 7).Formula = "=((D11*F11)/5)*100";

                    worksheet.Cell(11, 8).Formula = "=" + CallIV1.ToString() + "%";

                    if (CallCustomIV1 == "")
                        CallCustomIV1 = "0";
                    worksheet.Cell(11, 9).Formula = "=" + CallCustomIV1.ToString() + "%";

                    worksheet.Cell(11, 10).Formula = "=" + CallRF1.ToString() + "%";

                    if (CallCustomRF1 == "")
                        CallCustomRF1 = "0";
                    worksheet.Cell(11, 11).Formula = "=" + CallCustomRF1.ToString() + "%";

                    worksheet.Cell(11, 12).Formula = "=" + CallIV2.ToString() + "%";

                    if (CallCustomIV2 == "")
                        CallCustomIV2 = "0";
                    worksheet.Cell(11, 13).Formula = "=" + CallCustomIV2.ToString() + "%";

                    worksheet.Cell(11, 14).Formula = "=" + CallRF2.ToString() + "%";

                    if (CallCustomRF2 == "")
                        CallCustomRF2 = "0";
                    worksheet.Cell(11, 15).Formula = "=" + CallCustomRF2.ToString() + "%";

                    if (SalesComments != null)
                        worksheet.Cell(13, 2).Value = SalesComments.ToString();
                    else
                        worksheet.Cell(13, 2).Value = "";

                    if (TradingComments != null)
                        worksheet.Cell(14, 2).Value = TradingComments.ToString();
                    else
                        worksheet.Cell(14, 2).Value = "";

                    if (CouponScenario != null)
                        worksheet.Cell(15, 2).Value = CouponScenario.ToString();
                    else
                        worksheet.Cell(15, 2).Value = "";

                    //---------------Write Strike 1 IV Grid-----------------START------------
                    worksheet.Cell(1, 16).Formula = "=C11-5";
                    worksheet.Cell(1, 17).Formula = "=ROUND($B$4*30.417,0)";
                    worksheet.Cell(1, 18).Formula = "=ROUND($B$4*30.417,0) - (1*$H$8)";
                    worksheet.Cell(1, 19).Formula = "=ROUND($B$4*30.417,0) - (2*$H$8)";
                    worksheet.Cell(1, 20).Formula = "=ROUND($B$4*30.417,0) - (3*$H$8)";
                    worksheet.Cell(1, 21).Formula = "=ROUND($B$4*30.417,0) - (4*$H$8)";
                    worksheet.Cell(1, 22).Formula = "=ROUND($B$4*30.417,0) - (5*$H$8)";

                    worksheet.Cell(2, 16).Formula = "0";
                    worksheet.Cell(3, 16).Formula = "=$P$2+1*$D$8";
                    worksheet.Cell(4, 16).Formula = "=$P$2+2*$D$8";
                    worksheet.Cell(5, 16).Formula = "=$P$2+3*$D$8";
                    worksheet.Cell(6, 16).Formula = "=$P$2+4*$D$8";
                    worksheet.Cell(7, 16).Formula = "=$P$2+5*$D$8";

                    worksheet.Cell(2, 17).Formula = "=HoadleyOptions2(\"p\",1,\"c\",($C$11-5),100,Q1,IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(3, 17).Formula = "=HoadleyOptions2(\"p\",1,\"c\",($C$11-5),100,($Q$1-P3),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(4, 17).Formula = "=HoadleyOptions2(\"p\",1,\"c\",($C$11-5),100,($Q$1-P4),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(5, 17).Formula = "=HoadleyOptions2(\"p\",1,\"c\",($C$11-5),100,($Q$1-P5),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(6, 17).Formula = "=HoadleyOptions2(\"p\",1,\"c\",($C$11-5),100,($Q$1-P6),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(7, 17).Formula = "=HoadleyOptions2(\"p\",1,\"c\",($C$11-5),100,($Q$1-P7),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";

                    worksheet.Cell(2, 18).Formula = "=HoadleyOptions2(\"p\",1,\"c\",($C$11-5),100,R1,IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(3, 18).Formula = "=HoadleyOptions2(\"p\",1,\"c\",($C$11-5),100,($R$1-P3),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(4, 18).Formula = "=HoadleyOptions2(\"p\",1,\"c\",($C$11-5),100,($R$1-P4),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(5, 18).Formula = "=HoadleyOptions2(\"p\",1,\"c\",($C$11-5),100,($R$1-P5),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(6, 18).Formula = "=HoadleyOptions2(\"p\",1,\"c\",($C$11-5),100,($R$1-P6),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(7, 18).Formula = "=HoadleyOptions2(\"p\",1,\"c\",($C$11-5),100,($R$1-P7),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";

                    worksheet.Cell(2, 19).Formula = "=HoadleyOptions2(\"p\",1,\"c\",($C$11-5),100,S1,IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(3, 19).Formula = "=HoadleyOptions2(\"p\",1,\"c\",($C$11-5),100,($S$1-P3),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(4, 19).Formula = "=HoadleyOptions2(\"p\",1,\"c\",($C$11-5),100,($S$1-P4),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(5, 19).Formula = "=HoadleyOptions2(\"p\",1,\"c\",($C$11-5),100,($S$1-P5),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(6, 19).Formula = "=HoadleyOptions2(\"p\",1,\"c\",($C$11-5),100,($S$1-P6),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(7, 19).Formula = "=HoadleyOptions2(\"p\",1,\"c\",($C$11-5),100,($S$1-P7),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";

                    worksheet.Cell(2, 20).Formula = "=HoadleyOptions2(\"p\",1,\"c\",($C$11-5),100,T1,IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(3, 20).Formula = "=HoadleyOptions2(\"p\",1,\"c\",($C$11-5),100,($T$1-P3),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(4, 20).Formula = "=HoadleyOptions2(\"p\",1,\"c\",($C$11-5),100,($T$1-P4),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(5, 20).Formula = "=HoadleyOptions2(\"p\",1,\"c\",($C$11-5),100,($T$1-P5),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(6, 20).Formula = "=HoadleyOptions2(\"p\",1,\"c\",($C$11-5),100,($T$1-P6),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(7, 20).Formula = "=HoadleyOptions2(\"p\",1,\"c\",($C$11-5),100,($T$1-P7),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";

                    worksheet.Cell(2, 21).Formula = "=HoadleyOptions2(\"p\",1,\"c\",($C$11-5),100,U1,IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(3, 21).Formula = "=HoadleyOptions2(\"p\",1,\"c\",($C$11-5),100,($U$1-P3),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(4, 21).Formula = "=HoadleyOptions2(\"p\",1,\"c\",($C$11-5),100,($U$1-P4),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(5, 21).Formula = "=HoadleyOptions2(\"p\",1,\"c\",($C$11-5),100,($U$1-P5),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(6, 21).Formula = "=HoadleyOptions2(\"p\",1,\"c\",($C$11-5),100,($U$1-P6),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(7, 21).Formula = "=HoadleyOptions2(\"p\",1,\"c\",($C$11-5),100,($U$1-P7),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";

                    worksheet.Cell(2, 22).Formula = "=HoadleyOptions2(\"p\",1,\"c\",($C$11-5),100,V1,IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(3, 22).Formula = "=HoadleyOptions2(\"p\",1,\"c\",($C$11-5),100,($V$1-P3),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(4, 22).Formula = "=HoadleyOptions2(\"p\",1,\"c\",($C$11-5),100,($V$1-P4),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(5, 22).Formula = "=HoadleyOptions2(\"p\",1,\"c\",($C$11-5),100,($V$1-P5),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(6, 22).Formula = "=HoadleyOptions2(\"p\",1,\"c\",($C$11-5),100,($V$1-P6),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(7, 22).Formula = "=HoadleyOptions2(\"p\",1,\"c\",($C$11-5),100,($V$1-P7),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";

                    //Int32 intColStart = 16;
                    //worksheet.Cell(1, intColStart).Formula = Convert.ToString(objCallBinary.CallStrike1);

                    //for (int i = 0; i < objCallBinary.FinalAveragingMonth; i++)
                    //{
                    //    worksheet.Cell(1, intColStart + (i + 1)).Formula = Convert.ToString(Convert.ToInt32(objCallBinary.OptionTenure * 30.417) - objCallBinary.FinalAveragingDaysDiff * i);
                    //}

                    //for (int i = 0; i < objCallBinary.InitialAveragingMonth; i++)
                    //{
                    //    worksheet.Cell(i + 2, 16).Value = Convert.ToString(objCallBinary.InitialAveragingDaysDiff * i);
                    //    intColStart = 16;

                    //    if (i == 0)
                    //    {
                    //        for (int j = 0; j < objCallBinary.FinalAveragingMonth; j++)
                    //        {
                    //            worksheet.Cell(i + 2, intColStart + (j + 1)).Formula = Convert.ToString(Convert.ToInt32(objCallBinary.OptionTenure * 30.417) - objCallBinary.FinalAveragingDaysDiff * j);
                    //        }

                    //        if (objCallBinary.InitialAveragingMonth >= 1)
                    //            worksheet.Cell(i + 2, 17).Formula = "=HoadleyOptions2(\"p\",1,\"c\",($C$11-5),100,Q1,IF(I11>0,I11,H11),IF(K11>0,K11,J11))";

                    //        if (objCallBinary.InitialAveragingMonth >= 2)
                    //            worksheet.Cell(i + 2, 18).Formula = "=HoadleyOptions2(\"p\",1,\"c\",($C$11-5),100,R1,IF(I11>0,I11,H11),IF(K11>0,K11,J11))";

                    //        if (objCallBinary.InitialAveragingMonth >= 3)
                    //            worksheet.Cell(i + 2, 19).Formula = "=HoadleyOptions2(\"p\",1,\"c\",($C$11-5),100,S1,IF(I11>0,I11,H11),IF(K11>0,K11,J11))";

                    //        if (objCallBinary.InitialAveragingMonth >= 4)
                    //            worksheet.Cell(i + 2, 20).Formula = "=HoadleyOptions2(\"p\",1,\"c\",($C$11-5),100,T1,IF(I11>0,I11,H11),IF(K11>0,K11,J11))";

                    //        if (objCallBinary.InitialAveragingMonth >= 5)
                    //            worksheet.Cell(i + 2, 21).Formula = "=HoadleyOptions2(\"p\",1,\"c\",($C$11-5),100,U1,IF(I11>0,I11,H11),IF(K11>0,K11,J11))";

                    //        if (objCallBinary.InitialAveragingMonth >= 6)
                    //            worksheet.Cell(i + 2, 22).Formula = "=HoadleyOptions2(\"p\",1,\"c\",($C$11-5),100,V1,IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    //    }
                    //    else
                    //    {
                    //        if (objCallBinary.InitialAveragingMonth >= 1)
                    //            worksheet.Cell(i + 2, 17).Formula = "=HoadleyOptions2(\"p\",1,\"c\",($C$11-5),100,($Q$1-P" + Convert.ToInt32(i + 2) + "),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";

                    //        if (objCallBinary.InitialAveragingMonth >= 2)
                    //            worksheet.Cell(i + 2, 18).Formula = "=HoadleyOptions2(\"p\",1,\"c\",($C$11-5),100,($R$1-P" + Convert.ToInt32(i + 2) + "),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";

                    //        if (objCallBinary.InitialAveragingMonth >= 3)
                    //            worksheet.Cell(i + 2, 19).Formula = "=HoadleyOptions2(\"p\",1,\"c\",($C$11-5),100,($S$1-P" + Convert.ToInt32(i + 2) + "),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";

                    //        if (objCallBinary.InitialAveragingMonth >= 4)
                    //            worksheet.Cell(i + 2, 20).Formula = "=HoadleyOptions2(\"p\",1,\"c\",($C$11-5),100,($T$1-P" + Convert.ToInt32(i + 2) + "),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";

                    //        if (objCallBinary.InitialAveragingMonth >= 5)
                    //            worksheet.Cell(i + 2, 21).Formula = "=HoadleyOptions2(\"p\",1,\"c\",($C$11-5),100,($U$1-P" + Convert.ToInt32(i + 2) + "),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";

                    //        if (objCallBinary.InitialAveragingMonth >= 6)
                    //            worksheet.Cell(i + 2, 22).Formula = "=HoadleyOptions2(\"p\",1,\"c\",($C$11-5),100,($V$1-P" + Convert.ToInt32(i + 2) + "),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    //    }
                    //}
                    //---------------Write Strike 1 IV Grid-----------------END--------------

                    //---------------Write Strike 2 IV Grid-----------------START------------
                    worksheet.Cell(1, 24).Formula = "=C11";
                    worksheet.Cell(1, 25).Formula = "=ROUND($B$4*30.417,0)";
                    worksheet.Cell(1, 26).Formula = "=ROUND($B$4*30.417,0) - (1*$H$8)";
                    worksheet.Cell(1, 27).Formula = "=ROUND($B$4*30.417,0) - (2*$H$8)";
                    worksheet.Cell(1, 28).Formula = "=ROUND($B$4*30.417,0) - (3*$H$8)";
                    worksheet.Cell(1, 29).Formula = "=ROUND($B$4*30.417,0) - (4*$H$8)";
                    worksheet.Cell(1, 30).Formula = "=ROUND($B$4*30.417,0) - (5*$H$8)";

                    worksheet.Cell(2, 24).Formula = "0";
                    worksheet.Cell(3, 24).Formula = "=$X$2+1*$D$8";
                    worksheet.Cell(4, 24).Formula = "=$X$2+2*$D$8";
                    worksheet.Cell(5, 24).Formula = "=$X$2+3*$D$8";
                    worksheet.Cell(6, 24).Formula = "=$X$2+4*$D$8";
                    worksheet.Cell(7, 24).Formula = "=$X$2+5*$D$8";

                    worksheet.Cell(2, 25).Formula = "=HoadleyOptions2(\"p\",1,\"c\",$C$11,100,Y1,IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(3, 25).Formula = "=HoadleyOptions2(\"p\",1,\"c\",$C$11,100,($Y$1-X3),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(4, 25).Formula = "=HoadleyOptions2(\"p\",1,\"c\",$C$11,100,($Y$1-X4),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(5, 25).Formula = "=HoadleyOptions2(\"p\",1,\"c\",$C$11,100,($Y$1-X5),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(6, 25).Formula = "=HoadleyOptions2(\"p\",1,\"c\",$C$11,100,($Y$1-X6),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(7, 25).Formula = "=HoadleyOptions2(\"p\",1,\"c\",$C$11,100,($Y$1-X7),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";

                    worksheet.Cell(2, 26).Formula = "=HoadleyOptions2(\"p\",1,\"c\",$C$11,100,Z1,IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(3, 26).Formula = "=HoadleyOptions2(\"p\",1,\"c\",$C$11,100,($Z$1-X3),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(4, 26).Formula = "=HoadleyOptions2(\"p\",1,\"c\",$C$11,100,($Z$1-X4),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(5, 26).Formula = "=HoadleyOptions2(\"p\",1,\"c\",$C$11,100,($Z$1-X5),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(6, 26).Formula = "=HoadleyOptions2(\"p\",1,\"c\",$C$11,100,($Z$1-X6),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(7, 26).Formula = "=HoadleyOptions2(\"p\",1,\"c\",$C$11,100,($Z$1-X7),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";

                    worksheet.Cell(2, 27).Formula = "=HoadleyOptions2(\"p\",1,\"c\",$C$11,100,AA1,IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(3, 27).Formula = "=HoadleyOptions2(\"p\",1,\"c\",$C$11,100,($AA$1-X3),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(4, 27).Formula = "=HoadleyOptions2(\"p\",1,\"c\",$C$11,100,($AA$1-X4),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(5, 27).Formula = "=HoadleyOptions2(\"p\",1,\"c\",$C$11,100,($AA$1-X5),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(6, 27).Formula = "=HoadleyOptions2(\"p\",1,\"c\",$C$11,100,($AA$1-X6),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(7, 27).Formula = "=HoadleyOptions2(\"p\",1,\"c\",$C$11,100,($AA$1-X7),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";

                    worksheet.Cell(2, 28).Formula = "=HoadleyOptions2(\"p\",1,\"c\",$C$11,100,AB1,IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(3, 28).Formula = "=HoadleyOptions2(\"p\",1,\"c\",$C$11,100,($AB$1-X3),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(4, 28).Formula = "=HoadleyOptions2(\"p\",1,\"c\",$C$11,100,($AB$1-X4),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(5, 28).Formula = "=HoadleyOptions2(\"p\",1,\"c\",$C$11,100,($AB$1-X5),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(6, 28).Formula = "=HoadleyOptions2(\"p\",1,\"c\",$C$11,100,($AB$1-X6),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(7, 28).Formula = "=HoadleyOptions2(\"p\",1,\"c\",$C$11,100,($AB$1-X7),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";

                    worksheet.Cell(2, 29).Formula = "=HoadleyOptions2(\"p\",1,\"c\",$C$11,100,AC1,IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(3, 29).Formula = "=HoadleyOptions2(\"p\",1,\"c\",$C$11,100,($AC$1-X3),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(4, 29).Formula = "=HoadleyOptions2(\"p\",1,\"c\",$C$11,100,($AC$1-X4),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(5, 29).Formula = "=HoadleyOptions2(\"p\",1,\"c\",$C$11,100,($AC$1-X5),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(6, 29).Formula = "=HoadleyOptions2(\"p\",1,\"c\",$C$11,100,($AC$1-X6),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(7, 29).Formula = "=HoadleyOptions2(\"p\",1,\"c\",$C$11,100,($AC$1-X7),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";

                    worksheet.Cell(2, 30).Formula = "=HoadleyOptions2(\"p\",1,\"c\",$C$11,100,AD1,IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(3, 30).Formula = "=HoadleyOptions2(\"p\",1,\"c\",$C$11,100,($AD$1-X3),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(4, 30).Formula = "=HoadleyOptions2(\"p\",1,\"c\",$C$11,100,($AD$1-X4),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(5, 30).Formula = "=HoadleyOptions2(\"p\",1,\"c\",$C$11,100,($AD$1-X5),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(6, 30).Formula = "=HoadleyOptions2(\"p\",1,\"c\",$C$11,100,($AD$1-X6),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(7, 30).Formula = "=HoadleyOptions2(\"p\",1,\"c\",$C$11,100,($AD$1-X7),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    //---------------Write Strike 2 IV Grid-----------------END--------------

                    xlPackage.Save();
                }

                //if (System.IO.File.Exists(strTargetFileName))
                //{
                //    FileInfo TemplateFile = new FileInfo(strTargetFileName);

                //    Response.Clear();
                //    Response.ClearHeaders();
                //    Response.ClearContent();
                //    Response.AddHeader("content-disposition", "attachment; filename=" + Path.GetFileName(strTargetFileName));
                //    Response.AddHeader("Content-Type", "application/Excel");
                //    Response.ContentType = "application/vnd.xls";
                //    Response.AddHeader("Content-Length", TemplateFile.Length.ToString());
                //    Response.WriteFile(TemplateFile.FullName);
                //    Response.End();
                //}

                return Json("");
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];

                ClearCallBinarySession();

                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "ExportCallBinary", objUserMaster.UserID);
                //return RedirectToAction("ErrorPage", "Login");

                return Json("");
            }
        }

        public virtual void ExportCallBinary(CallBinary objCallBinary)
        {
            try
            {
                string strTemplateFilePath = Server.MapPath("~/Templates");
                string strTemplateFileName = strTemplateFilePath + "\\CallBinaryTemplate.xlsx";

                string strTargetFilePath = Server.MapPath("~/OutputFiles");
                string strTargetFileName = strTargetFilePath + "\\" + objCallBinary.ProductID + "_CallBinary.xlsx";

                string Role = Convert.ToString(Session["Role"]);

                if (System.IO.File.Exists(strTargetFileName))
                    System.IO.File.Delete(strTargetFileName);

                FileInfo objTemplateFileInfo = new FileInfo(strTemplateFileName);
                objTemplateFileInfo.CopyTo(strTargetFileName);

                FileInfo objTargetFileInfo = new FileInfo(strTargetFileName);

                Underlying objUnderlying = objCallBinary.UnderlyingList.Find(delegate(Underlying oUnderlying) { return oUnderlying.UnderlyingID == objCallBinary.UnderlyingID; });

                using (var xlPackage = new ExcelPackage(objTargetFileInfo))
                {
                    var worksheet = xlPackage.Workbook.Worksheets["CallBinary"];

                    worksheet.Cell(1, 2).Value = objCallBinary.ProductID.ToString();
                    worksheet.Cell(1, 4).Value = objCallBinary.Distributor.ToString().ToUpper();
                    worksheet.Cell(1, 6).Value = objUnderlying.UnderlyingShortName;

                    worksheet.Cell(2, 2).Formula = "=" + objCallBinary.EdelweissBuiltIn.ToString() + "%";
                    worksheet.Cell(2, 4).Formula = "=" + objCallBinary.DistributorBuiltIn.ToString() + "%";
                    worksheet.Cell(2, 6).Formula = "=B2+D2+H2";
                    worksheet.Cell(2, 8).Formula = "=" + objCallBinary.BuiltInAdjustment.ToString() + "%";

                    worksheet.Cell(3, 2).Formula = "=((POWER((1+D3),(12/D4))-1)*100) %";
                    worksheet.Cell(3, 4).Formula = "=" + objCallBinary.FixedCoupon.ToString() + "%";
                    worksheet.Cell(3, 6).Formula = "=((POWER((1+H3),(12/D4))-1)*100) %";
                    worksheet.Cell(3, 8).Formula = "=" + objCallBinary.MaxCoupon.ToString() + "%";

                    worksheet.Cell(4, 2).Value = objCallBinary.OptionTenure.ToString();
                    worksheet.Cell(4, 4).Formula = "=ROUND(F4/30.417,0)";
                    worksheet.Cell(4, 6).Value = objCallBinary.RedemptionPeriodDays.ToString();

                    #region Get Entity Name
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
                    LookupMaster objLookupMasterEntity = EntityList.Find(delegate(LookupMaster oLookupMaster) { return oLookupMaster.LookupID == objCallBinary.EntityID; });
                    worksheet.Cell(5, 2).Value = objLookupMasterEntity.LookupDescription;
                    #endregion

                    #region Get Is Secured
                    objLookupResult = null;
                    LookupResultList = null;
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
                    LookupMaster objLookupMasterIsSecured = IsSecuredList.Find(delegate(LookupMaster oLookupMaster) { return oLookupMaster.LookupID == objCallBinary.IsSecuredID; });
                    worksheet.Cell(5, 4).Value = objLookupMasterIsSecured.LookupDescription;
                    #endregion
                    worksheet.Cell(5, 6).Formula = "=" + objCallBinary.DeploymentRate.ToString() + "%";
                    worksheet.Cell(5, 8).Formula = "=" + objCallBinary.CustomDeploymentRate.ToString() + "%";

                    worksheet.Cell(6, 2).Formula = "=(100-(B2+D2)*100)-(100*(1+ROUND(D3,4)))/(POWER((1+(IF(H5>0,H5,F5))),(D4/12)))";
                    worksheet.Cell(6, 4).Formula = "=G11";
                    worksheet.Cell(6, 6).Formula = "=B6 + D6";

                    worksheet.Cell(8, 2).Value = objCallBinary.InitialAveragingMonth.ToString();
                    worksheet.Cell(8, 4).Value = objCallBinary.InitialAveragingDaysDiff.ToString();
                    worksheet.Cell(8, 6).Value = objCallBinary.FinalAveragingMonth.ToString();
                    worksheet.Cell(8, 8).Value = objCallBinary.FinalAveragingDaysDiff.ToString();

                    worksheet.Cell(11, 1).Value = "Call Binary Long";
                    worksheet.Cell(11, 2).Value = objUnderlying.UnderlyingShortName;
                    worksheet.Cell(11, 3).Value = objCallBinary.CallStrike1.ToString();
                    worksheet.Cell(11, 4).Formula = "=" + objCallBinary.CallCouponRise.ToString() + "%";
                    worksheet.Cell(11, 5).Value = objCallBinary.CallPrice.ToString();
                    worksheet.Cell(11, 6).Value = objCallBinary.CallDiscountedPrice.ToString();
                    worksheet.Cell(11, 7).Value = objCallBinary.CallPrAdjustmentPrice.ToString();

                    if (Role == "Sales")
                    {
                        worksheet.Cell(10, 8).Value = "";
                        worksheet.Cell(10, 9).Value = "";
                        worksheet.Cell(10, 10).Value = "";
                        worksheet.Cell(10, 11).Value = "";
                        worksheet.Cell(10, 12).Value = "";
                        worksheet.Cell(10, 13).Value = "";
                        worksheet.Cell(10, 14).Value = "";
                        worksheet.Cell(10, 15).Value = "";
                    }
                    else
                    {
                        worksheet.Cell(11, 8).Formula = "=" + objCallBinary.CallIV1.ToString() + "%";
                        worksheet.Cell(11, 9).Formula = "=" + objCallBinary.CallCustomIV1.ToString() + "%";
                        worksheet.Cell(11, 10).Formula = "=" + objCallBinary.CallRF1.ToString() + "%";
                        worksheet.Cell(11, 11).Formula = "=" + objCallBinary.CallCustomRF1.ToString() + "%";
                        worksheet.Cell(11, 12).Formula = "=" + objCallBinary.CallIV2.ToString() + "%";
                        worksheet.Cell(11, 13).Formula = "=" + objCallBinary.CallCustomIV2.ToString() + "%";
                        worksheet.Cell(11, 14).Formula = "=" + objCallBinary.CallRF2.ToString() + "%";
                        worksheet.Cell(11, 15).Formula = "=" + objCallBinary.CallCustomRF2.ToString() + "%";
                    }

                    if (objCallBinary.SalesComments != null)
                        worksheet.Cell(13, 2).Value = objCallBinary.SalesComments.ToString();
                    else
                        worksheet.Cell(13, 2).Value = "";

                    if (objCallBinary.TradingComments != null)
                        worksheet.Cell(14, 2).Value = objCallBinary.TradingComments.ToString();
                    else
                        worksheet.Cell(14, 2).Value = "";

                    if (objCallBinary.CouponScenario != null)
                        worksheet.Cell(15, 2).Value = objCallBinary.CouponScenario.ToString();
                    else
                        worksheet.Cell(15, 2).Value = "";

                    xlPackage.Save();
                }

                if (System.IO.File.Exists(strTargetFileName))
                {
                    FileInfo TemplateFile = new FileInfo(strTargetFileName);

                    Response.Clear();
                    Response.ClearHeaders();
                    Response.ClearContent();
                    Response.AddHeader("content-disposition", "attachment; filename=" + Path.GetFileName(strTargetFileName));
                    Response.AddHeader("Content-Type", "application/Excel");
                    Response.ContentType = "application/vnd.xls";
                    Response.AddHeader("Content-Length", TemplateFile.Length.ToString());
                    Response.WriteFile(TemplateFile.FullName);
                    Response.End();
                }
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];

                ClearCallBinarySession();

                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "ExportCallBinary", objUserMaster.UserID);
                //return RedirectToAction("ErrorPage", "Login");
            }
        }

        public JsonResult ManageCallBinary(string ProductID, string Distributor, string EdelweissBuiltIn, string DistributorBuiltIn, string BuiltInAdjustment, string TotalBuiltIn, string FixedCoupon, string IRR, string IsIRR, string MaxCoupon,
            string MaxCouponIRR, string IsMaxCouponIRR, string DeploymentRate, string CustomerDeploymentRate, string Remaining, string Underlying, string TotalOptionPrice, string NetRemaining,
            string OptionTenureMonth, string RedemptionPeriodMonth, string IsRedemptionPeriodMonth, string RedemptionPeriodDays, string InitialAveragingMonth, string InitialAveragingDaysDiff, string FinalAveragingMonth,
            string FinalAveragingDaysDiff, string CallBinaryLong, string CallUnderlying, string CallStrike1, string CallStrike2, string CallCouponRise, string CallPrice, string CallDiscountedPrice,
            string CallPRAdjustedPrice, string CallIV1, string CallCustomIV1, string CallRF1, string CallCustomRF1, string CallIV2, string CallCustomIV2, string CallRF2,
            string CallCustomRF2, string SalesComments, string TradingComments, string CouponScenario, string CopyProductID,
            string CallStrike1Summary, string CallStrike2Summary, string ExportCallStrike1Summary, string ExportCallStrike2Summary, string Entity, string IsSecured)
        {
            try
            {
                CallStrike1Summary = System.Uri.UnescapeDataString(CallStrike1Summary);
                CallStrike2Summary = System.Uri.UnescapeDataString(CallStrike2Summary);
                ExportCallStrike1Summary = System.Uri.UnescapeDataString(ExportCallStrike1Summary);
                ExportCallStrike2Summary = System.Uri.UnescapeDataString(ExportCallStrike2Summary);

                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];

                if (CustomerDeploymentRate == "")
                    CustomerDeploymentRate = "0";

                if (TotalOptionPrice == "")
                    TotalOptionPrice = "0";

                if (NetRemaining == "")
                    NetRemaining = "0";

                if (CallStrike1 == "")
                    CallStrike1 = "0";

                if (CallCouponRise == "")
                    CallCouponRise = "0";

                if (CallPrice == "")
                    CallPrice = "0";

                if (CallDiscountedPrice == "")
                    CallDiscountedPrice = "0";

                if (CallPRAdjustedPrice == "")
                    CallPRAdjustedPrice = "0";

                if (CallIV1 == "")
                    CallIV1 = "0";

                if (CallCustomIV1 == "")
                    CallCustomIV1 = "0";

                if (CallRF1 == "")
                    CallRF1 = "0";

                if (CallCustomRF1 == "")
                    CallCustomRF1 = "0";

                if (CallIV2 == "")
                    CallIV2 = "0";

                if (CallCustomIV2 == "")
                    CallCustomIV2 = "0";

                if (CallRF2 == "")
                    CallRF2 = "0";

                if (CallCustomRF2 == "")
                    CallCustomRF2 = "0";

                string ParentProductID = "";
                if (Session["ParentProductID"] != null)
                    ParentProductID = (string)Session["ParentProductID"];

                ObjectResult<ManageCallBinaryResult> objManageCallBinaryResult = objSP_PRICINGEntities.SP_MANAGE_CALL_BINARY_DETAILS(ProductID, ParentProductID, Distributor, Convert.ToDouble(EdelweissBuiltIn),
                        Convert.ToDouble(DistributorBuiltIn), Convert.ToDouble(BuiltInAdjustment), Convert.ToDouble(TotalBuiltIn), Convert.ToDouble(FixedCoupon), Convert.ToDouble(IRR), Convert.ToBoolean(IsIRR), Convert.ToDouble(MaxCoupon), Convert.ToDouble(MaxCouponIRR), Convert.ToBoolean(IsMaxCouponIRR),
                        Convert.ToDouble(DeploymentRate), Convert.ToDouble(CustomerDeploymentRate), Convert.ToDouble(Remaining), Convert.ToInt32(Underlying), Convert.ToDouble(TotalOptionPrice),
                        Convert.ToDouble(NetRemaining), Convert.ToInt32(OptionTenureMonth), Convert.ToDouble(RedemptionPeriodMonth), Convert.ToBoolean(IsRedemptionPeriodMonth), Convert.ToInt32(RedemptionPeriodDays), Convert.ToInt32(InitialAveragingMonth),
                        Convert.ToInt32(InitialAveragingDaysDiff), Convert.ToInt32(FinalAveragingMonth), Convert.ToInt32(FinalAveragingDaysDiff), SalesComments, TradingComments, CouponScenario, Convert.ToInt32(Entity), Convert.ToInt32(IsSecured), objUserMaster.UserID,
                        CallBinaryLong, Convert.ToDouble(CallStrike1), Convert.ToDouble(CallStrike2), Convert.ToDouble(CallCouponRise), Convert.ToDouble(CallPrice), Convert.ToDouble(CallDiscountedPrice),
                        Convert.ToDouble(CallPRAdjustedPrice), Convert.ToDouble(CallIV1), Convert.ToDouble(CallCustomIV1), Convert.ToDouble(CallRF1), Convert.ToDouble(CallCustomRF1),
                        Convert.ToDouble(CallIV2), Convert.ToDouble(CallCustomIV2), Convert.ToDouble(CallRF2), Convert.ToDouble(CallCustomRF2), CopyProductID,
                        ExportCallStrike1Summary, CallStrike1Summary, ExportCallStrike2Summary, CallStrike2Summary);
                List<ManageCallBinaryResult> ManageCallBinaryResultList = objManageCallBinaryResult.ToList();

                Session.Remove("ParentProductID");

                return Json(ManageCallBinaryResultList[0].ProductID);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];

                ClearCallBinarySession();

                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "ExportCallBinary", objUserMaster.UserID);
                return Json("");// RedirectToAction("ErrorPage", "Login");
            }
        }

        #endregion

        #region Put Binary
        [HttpGet]
        public ActionResult PutBinary(string ProductID, string GenerateGraph, bool IsQuotron = false)
        {
            LoginController objLoginController = new LoginController();
            try
            {
                if (ValidateSession())
                {
                    PutBinary objPutBinary = new PutBinary();

                    #region Menu Access By on Role

                    Int32 intResult = 0;
                    // bool PPorNonPP = false;

                    UserMaster objUserMaster = new UserMaster();
                    objUserMaster = (UserMaster)Session["LoggedInUser"];

                    var Result = objSP_PRICINGEntities.VALIDATE_MENU_BY_ROLE(objUserMaster.RoleID, "BPB");
                    intResult = Convert.ToInt32(Result.SingleOrDefault());

                    if (intResult == 0)
                        return RedirectToAction("UserNotAuthorize", "Login");

                    #endregion

                    #region Bind Underlying List
                    List<Underlying> UnderlyingList = new List<Underlying>();

                    ObjectResult<UnderlyingListResult> objUnderlyingListResult;
                    List<UnderlyingListResult> UnderlyingListResultList = new List<UnderlyingListResult>();

                    objUnderlyingListResult = objSP_PRICINGEntities.SP_FETCH_UNDERLYING_DETAILS();
                    UnderlyingListResultList = objUnderlyingListResult.ToList();

                    foreach (UnderlyingListResult oUnderlyingListResult in UnderlyingListResultList)
                    {
                        Underlying objUnderlying = new Underlying();
                        General.ReflectSingleData(objUnderlying, oUnderlyingListResult);

                        UnderlyingList.Add(objUnderlying);
                    }
                    objPutBinary.UnderlyingList = UnderlyingList;

                    //--Set default underlying--Added by Shweta on 3rd May 2016------------START--------------------
                    string strDefaultUnderlying = System.Configuration.ConfigurationManager.AppSettings["DefaultUnderlying"].ToUpper();
                    Underlying objDefaulyUnderlying = objPutBinary.UnderlyingList.Find(delegate(Underlying oUnderlying) { return oUnderlying.UnderlyingShortName == strDefaultUnderlying; });
                    objPutBinary.UnderlyingID = objDefaulyUnderlying.UnderlyingID;

                    objPutBinary.EntityID = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["DefaultEntityID"]);
                    objPutBinary.IsSecuredID = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["DefaultIsSecuredID"]);
                    //--Set default underlying--Added by Shweta on 3rd May 2016------------END----------------------
                    #endregion

                    if (ProductID != "" && ProductID != null)
                    {
                        ObjectResult<PutBinaryEditResult> objPutBinaryEditResult = objSP_PRICINGEntities.FETCH_PUT_BINARY_EDIT_DETAILS(ProductID);
                        List<PutBinaryEditResult> PutBinaryEditResultList = objPutBinaryEditResult.ToList();

                        General.ReflectSingleData(objPutBinary, PutBinaryEditResultList[0]);

                        DataSet dsResult = new DataSet();
                        dsResult = General.ExecuteDataSet("FETCH_UNDERLYING_CREATION_BYID", objPutBinary.UnderlyingID);

                        if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                        {
                            ViewBag.UnderlyingShortName = Convert.ToString(dsResult.Tables[0].Rows[0]["UNDERLYING_SHORTNAME"]);
                        }
                    }
                    if (GenerateGraph == "GenerateGraph")
                    {
                        objPutBinary = (PutBinary)TempData["PutBinaryGraph"];
                        ObjectResult<PutBinaryEditResult> objPutBinaryEditResult = objSP_PRICINGEntities.FETCH_PUT_BINARY_EDIT_DETAILS(objPutBinary.ProductID);
                        List<PutBinaryEditResult> PutBinaryEditResultList = objPutBinaryEditResult.ToList();
                        PutBinary oPutBinary = new PutBinary();
                        General.ReflectSingleData(oPutBinary, PutBinaryEditResultList[0]);

                        objPutBinary.Status = oPutBinary.Status;
                        // objPutBinary.SaveStatus = oPutBinary.SaveStatus;
                        return GeneratePutBinaryGraph(objPutBinary);
                    }

                    else if (Session["PutBinaryCopyQuote"] != null)
                    {
                        ViewBag.Message = "Quote is copied";
                        objPutBinary = (PutBinary)Session["PutBinaryCopyQuote"];
                        objPutBinary.UnderlyingID = Convert.ToInt32(Session["UnderlyingID"]);

                        ObjectResult<PutBinaryEditResult> objPutBinaryEditResult = objSP_PRICINGEntities.FETCH_PUT_BINARY_EDIT_DETAILS("");
                        List<PutBinaryEditResult> PutBinaryEditResultList = objPutBinaryEditResult.ToList();
                        PutBinary oPutBinary = new PutBinary();
                        if (PutBinaryEditResultList != null && PutBinaryEditResultList.Count > 0)
                            General.ReflectSingleData(oPutBinary, PutBinaryEditResultList[0]);

                        objPutBinary.ParentProductID = objPutBinary.ProductID;
                        objPutBinary.ProductID = "";
                        objPutBinary.Status = oPutBinary.Status;
                        objPutBinary.SaveStatus = oPutBinary.SaveStatus;
                        objPutBinary.IsCopyQuote = true;

                        //Added by Shweta on 10th May---------------START-------
                        objPutBinary.PutCustomIV1 = 0;
                        objPutBinary.PutCustomIV2 = 0;
                        objPutBinary.PutCustomRF1 = 0;
                        objPutBinary.PutCustomRF2 = 0;
                        //Added by Shweta on 10th May---------------END---------

                        //-------------Added by Shweta on 22nd July 2016 to Fetch Latest Deployment Rate------------START--------
                        string strDeploymentRate = "";
                        var DeploymentRate = objSP_PRICINGEntities.SP_FETCH_PRICING_DEPLOYMENT_RATE(Convert.ToInt32(objPutBinary.RedemptionPeriodDays), objPutBinary.EntityID, objPutBinary.IsSecuredID);
                        strDeploymentRate = Convert.ToString(DeploymentRate.SingleOrDefault());
                        objPutBinary.DeploymentRate = Convert.ToDouble(strDeploymentRate);
                        //-------------Added by Shweta on 22nd July 2016 to Fetch Latest Deployment Rate------------END----------
                    }

                    else if (Session["PutBinaryChildQuote"] != null)
                    {
                        ViewBag.Message = true;
                        objPutBinary = (PutBinary)Session["PutBinaryChildQuote"];
                        objPutBinary.UnderlyingID = Convert.ToInt32(Session["UnderlyingID"]);

                        ObjectResult<PutBinaryEditResult> objPutBinaryEditResult = objSP_PRICINGEntities.FETCH_PUT_BINARY_EDIT_DETAILS("");
                        List<PutBinaryEditResult> PutBinaryEditResultList = objPutBinaryEditResult.ToList();
                        PutBinary oPutBinary = new PutBinary();
                        if (PutBinaryEditResultList != null && PutBinaryEditResultList.Count > 0)
                            General.ReflectSingleData(oPutBinary, PutBinaryEditResultList[0]);

                        objPutBinary.ParentProductID = objPutBinary.ProductID;
                        objPutBinary.ProductID = "";
                        objPutBinary.Status = oPutBinary.Status;
                        objPutBinary.SaveStatus = oPutBinary.SaveStatus;
                        objPutBinary.IsChildQuote = true;
                    }
                    else if (Session["CancelQuote"] != null)
                    {
                        objPutBinary = (PutBinary)Session["CancelQuote"];

                        ObjectResult<PutBinaryEditResult> objPutBinaryEditResult = objSP_PRICINGEntities.FETCH_PUT_BINARY_EDIT_DETAILS(objPutBinary.ProductID);
                        List<PutBinaryEditResult> PutBinaryEditResultList = objPutBinaryEditResult.ToList();
                        PutBinary oPutBinary = new PutBinary();
                        if (PutBinaryEditResultList != null && PutBinaryEditResultList.Count > 0)
                            General.ReflectSingleData(oPutBinary, PutBinaryEditResultList[0]);

                        objPutBinary.Status = oPutBinary.Status;
                        objPutBinary.SaveStatus = oPutBinary.SaveStatus;

                        Session.Remove("CancelQuote");
                    }
                    else
                    {
                        Session.Remove("IsChildQuotePutBinary");
                        Session.Remove("ParentProductID");
                        Session.Remove("UnderlyingID");
                    }

                    if (IsQuotron == true)
                    {
                        objPutBinary.IsQuotron = true;
                    }

                    if (Session["PutBinaryChildQuote"] == null && Session["PutBinaryCopyQuote"] == null)
                        objPutBinary.SaveStatus = "";

                    if (Session["PutBinaryCopyQuote"] != null)
                        Session.Remove("PutBinaryCopyQuote");

                    if (Session["PutBinaryChildQuote"] != null)
                        Session.Remove("PutBinaryChildQuote");

                    if (ProductID == null)
                    {
                        objPutBinary.isGraphActive = false;
                        return View(objPutBinary);
                    }
                    else
                    {
                        return GeneratePutBinaryGraph(objPutBinary);
                    }
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

                ClearPutBinarySession();

                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "PutBinary Get", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        private ActionResult GeneratePutBinaryGraph(PutBinary objPutBinary)
        {
            LoginController objLoginController = new LoginController();
            try
            {
                if (ValidateSession())
                {
                    var transactionCounts = new List<Graph>();
                    transactionCounts = GenerateFixedMLDGraphCalculation(objPutBinary.PutStrike1, objPutBinary.MaxCoupon, objPutBinary.FixedCoupon, objPutBinary.RedemptionPeriodDays);
                    //PutBinary obj = new PutBinary();

                    #region Pie Chart For Put Binary
                    var xDataMonths = transactionCounts.Select(i => i.Column1).ToArray();
                    var yDataCounts = transactionCounts.Select(i => new object[] { i.Column2 }).ToArray();
                    var yDataCounts1 = transactionCounts.Select(i => new object[] { i.Column3 }).ToArray();



                    var PutBinaryChart = new Highcharts("pie")
                        //define the type of chart 
                                .InitChart(new Chart { DefaultSeriesType = ChartTypes.Line })
                        //overall Title of the chart 
                                .SetTitle(new Title { Text = "Put Binary" })
                        ////small label below the main Title
                        //        .SetSubtitle(new Subtitle { Text = "Accounting" })
                        //load the X values
                                .SetXAxis(new XAxis { Title = new XAxisTitle { Text = "Underlying Returns" }, Categories = xDataMonths, Labels = new XAxisLabels { Step = 2 } })
                        //set the Y title
                                .SetYAxis(new YAxis { Title = new YAxisTitle { Text = "Product Returns" } })
                                .SetTooltip(new Tooltip
                                {
                                    Enabled = true,
                                    Formatter = @"function() { return '<b>'+ this.series.name +'</b><br/>'+ this.x +': '+ this.y; }"
                                })
                                .SetPlotOptions(new PlotOptions
                                {
                                    Line = new PlotOptionsLine
                                    {
                                        DataLabels = new PlotOptionsLineDataLabels
                                        {
                                            Enabled = false
                                        },
                                        EnableMouseTracking = true
                                    }
                                })
                        //load the Y values 
                                .SetSeries(new[]
                    {
                        new Series {Name = "Coupon", Data = new Data(yDataCounts)},
                            //you can add more y data to create a second line
                           //  new Series { Name = "Strike", Data = new Data(yDataCounts1) }
                    });
                    #endregion

                    if (Session["PutBinaryCopyQuote"] != null)
                    {
                        ViewBag.Message = "Quote is copied";
                        objPutBinary = (PutBinary)Session["PutBinaryCopyQuote"];
                        Session.Remove("PutBinaryCopyQuote");
                    }
                    objPutBinary.PutBinaryChart = PutBinaryChart;
                    return View(objPutBinary);
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

                ClearPutBinarySession();

                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "GeneratePutBinaryGraph", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        [HttpPost, ValidateInput(false)]
        public ActionResult PutBinary(string Command, PutBinary objPutBinary, FormCollection objFormCollection)
        {
            LoginController objLoginController = new LoginController();
            try
            {
                if (ValidateSession())
                {
                    #region Bind Underlying List
                    List<Underlying> UnderlyingList = new List<Underlying>();

                    ObjectResult<UnderlyingListResult> objUnderlyingListResult;
                    List<UnderlyingListResult> UnderlyingListResultList = new List<UnderlyingListResult>();

                    objUnderlyingListResult = objSP_PRICINGEntities.SP_FETCH_UNDERLYING_DETAILS();
                    UnderlyingListResultList = objUnderlyingListResult.ToList();

                    foreach (UnderlyingListResult oUnderlyingListResult in UnderlyingListResultList)
                    {
                        Underlying objUnderlying = new Underlying();
                        General.ReflectSingleData(objUnderlying, oUnderlyingListResult);

                        UnderlyingList.Add(objUnderlying);
                    }
                    objPutBinary.UnderlyingList = UnderlyingList;
                    #endregion
                    if (Command == "ExportToExcel")
                    {
                        ExportPutBinary(objPutBinary);

                        return RedirectToAction("PutBinary");
                    }
                    else if (Command == "ExportPutStrike1Grid")
                    {
                        string PutStrike1Summary = objFormCollection["ExportPutStrike1Summary"].Replace("<html><head></head><body class=\"table-responsive col-md-12\">", "");
                        PutStrike1Summary = PutStrike1Summary.Replace("</body></html>", "");

                        string PutStrike2Summary = objFormCollection["ExportPutStrike2Summary"].Replace("<html><head></head><body class=\"table-responsive col-md-12\">", "");
                        PutStrike2Summary = PutStrike2Summary.Replace("</body></html>", "");

                        string StrikeHTML = PutStrike1Summary + "<br />" + PutStrike2Summary;

                        ExportStrikeGrid(StrikeHTML);

                        return RedirectToAction("PutBinary");
                    }
                    else if (Command == "ExportPutStrike2Grid")
                    {
                        string StrikeHTML = objFormCollection["ExportPutStrike2Summary"];

                        ExportStrikeGrid(StrikeHTML);

                        return RedirectToAction("FixedOrPR");
                    }
                    else if (Command == "CopyQuote")
                    {
                        Session["PutBinaryCopyQuote"] = objPutBinary;
                        Session["UnderlyingID"] = objPutBinary.UnderlyingID;

                        return RedirectToAction("PutBinary");
                    }
                    else if (Command == "CreateChildQuote")
                    {
                        Session.Remove("ParentProductID");
                        Session.Remove("IsChildQuotePutBinary");
                        Session.Remove("UnderlyingID");

                        Session["ParentProductID"] = objPutBinary.ProductID;
                        Session["UnderlyingID"] = objPutBinary.UnderlyingID;

                        objPutBinary.IsChildQuote = true;

                        Session["PutBinaryChildQuote"] = objPutBinary;
                        Session["IsChildQuotePutBinary"] = objPutBinary.IsChildQuote;

                        return RedirectToAction("PutBinary");
                    }
                    else if (Command == "GenerateGraph")
                    {
                        objPutBinary.isGraphActive = true;
                        TempData["PutBinaryGraph"] = objPutBinary;
                        return RedirectToAction("PutBinary", new { GenerateGraph = "GenerateGraph" });
                    }
                    else if (Command == "AddNewProduct")
                    {
                        var productID = objPutBinary.ProductID;
                        UserMaster objUserMaster = new UserMaster();
                        objUserMaster = (UserMaster)Session["LoggedInUser"];

                        EncryptDecrypt obj = new EncryptDecrypt();
                        var encryptedpaswd = obj.Encrypt(objUserMaster.Password, "SPPricing", CryptographyEngine.AlgorithmType.DES);
                        var ProductType = "PP";

                        var Url = "http://edemumnewuatvm4:63400/Login.aspx?UserId=" + objUserMaster.LoginName + "&Key=" + encryptedpaswd + "&ProductId=" + productID + "&ProductType=" + ProductType;
                        return Redirect(Url);
                    }
                    else if (Command == "Cancel")
                    {
                        Session["CancelQuote"] = objPutBinary;

                        return RedirectToAction("PutBinary");
                    }
                    else if (Command == "PricingInExcel")
                    {
                        objPutBinary.IsWorkingFileExport = OpenWorkingExcelFile("PB", objPutBinary.ProductID);

                        if (!objPutBinary.IsWorkingFileExport)
                            objPutBinary.WorkingFileStatus = "File Not Found";

                        return View(objPutBinary);
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

                ClearPutBinarySession();

                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "PutBinary Post", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        public virtual void ExportPutBinary(PutBinary objPutBinary)
        {
            try
            {
                string strTemplateFilePath = Server.MapPath("~/Templates");
                string strTemplateFileName = strTemplateFilePath + "//PutBinaryTemplate.xlsx";

                string strTargetFilePath = Server.MapPath("~/OutputFiles");
                string strTargetFileName = strTargetFilePath + "//" + objPutBinary.ProductID + "_PutBinary.xlsx";

                string Role = Convert.ToString(Session["Role"]);

                if (System.IO.File.Exists(strTargetFileName))
                    System.IO.File.Delete(strTargetFileName);

                FileInfo objTemplateFileInfo = new FileInfo(strTemplateFileName);
                objTemplateFileInfo.CopyTo(strTargetFileName);

                FileInfo objTargetFileInfo = new FileInfo(strTargetFileName);

                Underlying objUnderlying = objPutBinary.UnderlyingList.Find(delegate(Underlying oUnderlying) { return oUnderlying.UnderlyingID == objPutBinary.UnderlyingID; });

                using (var xlPackage = new ExcelPackage(objTargetFileInfo))
                {
                    var worksheet = xlPackage.Workbook.Worksheets["PutBinary"];

                    worksheet.Cell(1, 2).Value = objPutBinary.ProductID.ToString();
                    worksheet.Cell(1, 4).Value = objPutBinary.Distributor.ToString().ToUpper();
                    worksheet.Cell(1, 6).Value = objUnderlying.UnderlyingShortName;

                    worksheet.Cell(2, 2).Formula = "=" + objPutBinary.EdelweissBuiltIn.ToString() + "%";
                    worksheet.Cell(2, 4).Formula = "=" + objPutBinary.DistributorBuiltIn.ToString() + "%";
                    worksheet.Cell(2, 6).Formula = "=B2+D2+H2";
                    worksheet.Cell(2, 8).Formula = "=" + objPutBinary.BuiltInAdjustment.ToString() + "%";

                    worksheet.Cell(3, 2).Formula = "=((POWER((1+D3),(12/D4))-1)*100) %";
                    worksheet.Cell(3, 4).Formula = "=" + objPutBinary.FixedCoupon.ToString() + "%";
                    worksheet.Cell(3, 6).Formula = "=((POWER((1+H3),(12/D4))-1)*100) %";
                    worksheet.Cell(3, 8).Formula = "=" + objPutBinary.MaxCoupon.ToString() + "%";

                    worksheet.Cell(4, 2).Value = objPutBinary.OptionTenure.ToString();
                    worksheet.Cell(4, 4).Formula = "=ROUND(F4/30.417,0)";
                    worksheet.Cell(4, 6).Value = objPutBinary.RedemptionPeriodDays.ToString();

                    #region Get Entity Name
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
                    LookupMaster objLookupMasterEntity = EntityList.Find(delegate(LookupMaster oLookupMaster) { return oLookupMaster.LookupID == objPutBinary.EntityID; });
                    worksheet.Cell(5, 2).Value = objLookupMasterEntity.LookupDescription;
                    #endregion

                    #region Get Is Secured
                    objLookupResult = null;
                    LookupResultList = null;
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
                    LookupMaster objLookupMasterIsSecured = IsSecuredList.Find(delegate(LookupMaster oLookupMaster) { return oLookupMaster.LookupID == objPutBinary.IsSecuredID; });
                    worksheet.Cell(5, 4).Value = objLookupMasterIsSecured.LookupDescription;
                    #endregion
                    worksheet.Cell(5, 6).Formula = "=" + objPutBinary.DeploymentRate.ToString() + "%";
                    worksheet.Cell(5, 8).Formula = "=" + objPutBinary.CustomDeploymentRate.ToString() + "%";

                    worksheet.Cell(6, 2).Formula = "=(100-(B2+D2)*100)-(100*(1+ROUND(D3,4)))/(POWER((1+(IF(H5>0,H5,F5))),(D4/12)))";
                    worksheet.Cell(6, 4).Formula = "=G11";
                    worksheet.Cell(6, 6).Formula = "=B6 + D6";

                    worksheet.Cell(8, 2).Value = objPutBinary.InitialAveragingMonth.ToString();
                    worksheet.Cell(8, 4).Value = objPutBinary.InitialAveragingDaysDiff.ToString();
                    worksheet.Cell(8, 6).Value = objPutBinary.FinalAveragingMonth.ToString();
                    worksheet.Cell(8, 8).Value = objPutBinary.FinalAveragingDaysDiff.ToString();

                    worksheet.Cell(11, 1).Value = "Put Binary Short";
                    worksheet.Cell(11, 2).Value = objUnderlying.UnderlyingShortName;
                    worksheet.Cell(11, 3).Value = objPutBinary.PutStrike1.ToString();
                    worksheet.Cell(11, 4).Formula = "=" + objPutBinary.PutCouponFall.ToString() + "%";
                    worksheet.Cell(11, 5).Value = objPutBinary.PutPrice.ToString();
                    worksheet.Cell(11, 6).Value = objPutBinary.PutDiscountedPrice.ToString();
                    worksheet.Cell(11, 7).Value = objPutBinary.PutPrAdjustmentPrice.ToString();

                    if (Role == "Sales")
                    {
                        worksheet.Cell(10, 8).Value = "";
                        worksheet.Cell(10, 9).Value = "";
                        worksheet.Cell(10, 10).Value = "";
                        worksheet.Cell(10, 11).Value = "";
                        worksheet.Cell(10, 12).Value = "";
                        worksheet.Cell(10, 13).Value = "";
                        worksheet.Cell(10, 14).Value = "";
                        worksheet.Cell(10, 15).Value = "";
                    }
                    else
                    {
                        worksheet.Cell(11, 8).Formula = "=" + objPutBinary.PutIV1.ToString() + "%";
                        worksheet.Cell(11, 9).Formula = "=" + objPutBinary.PutCustomIV1.ToString() + "%";
                        worksheet.Cell(11, 10).Formula = "=" + objPutBinary.PutRF1.ToString() + "%";
                        worksheet.Cell(11, 11).Formula = "=" + objPutBinary.PutCustomRF1.ToString() + "%";
                        worksheet.Cell(11, 12).Formula = "=" + objPutBinary.PutIV2.ToString() + "%";
                        worksheet.Cell(11, 13).Formula = "=" + objPutBinary.PutCustomIV2.ToString() + "%";
                        worksheet.Cell(11, 14).Formula = "=" + objPutBinary.PutRF2.ToString() + "%";
                        worksheet.Cell(11, 15).Formula = "=" + objPutBinary.PutCustomRF2.ToString() + "%";
                    }

                    if (objPutBinary.SalesComments != null)
                        worksheet.Cell(13, 2).Value = objPutBinary.SalesComments.ToString();
                    else
                        worksheet.Cell(13, 2).Value = "";

                    if (objPutBinary.TradingComments != null)
                        worksheet.Cell(14, 2).Value = objPutBinary.TradingComments.ToString();
                    else
                        worksheet.Cell(14, 2).Value = "";

                    if (objPutBinary.CouponScenario != null)
                        worksheet.Cell(15, 2).Value = objPutBinary.CouponScenario.ToString();
                    else
                        worksheet.Cell(15, 2).Value = "";

                    xlPackage.Save();
                }

                if (System.IO.File.Exists(strTargetFileName))
                {
                    FileInfo TemplateFile = new FileInfo(strTargetFileName);

                    Response.Clear();
                    Response.ClearHeaders();
                    Response.ClearContent();
                    Response.AddHeader("content-disposition", "attachment; filename=" + Path.GetFileName(strTargetFileName));
                    Response.AddHeader("Content-Type", "application/Excel");
                    Response.ContentType = "application/vnd.xls";
                    Response.AddHeader("Content-Length", TemplateFile.Length.ToString());
                    Response.WriteFile(TemplateFile.FullName);
                    Response.End();
                }
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];

                ClearPutBinarySession();

                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "ExportPutBinary", objUserMaster.UserID);
                // return Json("");// return RedirectToAction("ErrorPage", "Login");
            }
        }

        public JsonResult ExportPutBinaryWorkingFile(string ProductID, string Distributor, string EdelweissBuiltIn, string DistributorBuiltIn, string BuiltInAdjustment, string TotalBuiltIn, string FixedCoupon, string IRR, string IsIRR, string MaxCoupon,
            string MaxCouponIRR, string IsMaxCouponIRR, string DeploymentRate, string CustomerDeploymentRate, string Remaining, string Underlying, string TotalOptionPrice, string NetRemaining,
            string OptionTenureMonth, string RedemptionPeriodMonth, string IsRedemptionPeriodMonth, string RedemptionPeriodDays, string InitialAveragingMonth, string InitialAveragingDaysDiff, string FinalAveragingMonth,
            string FinalAveragingDaysDiff, string PutBinaryLong, string PutUnderlying, string PutStrike1, string PutStrike2, string PutCouponFall, string PutPrice, string PutDiscountedPrice,
            string PutPRAdjustedPrice, string PutIV1, string PutCustomIV1, string PutRF1, string PutCustomRF1, string PutIV2, string PutCustomIV2, string PutRF2,
            string PutCustomRF2, string SalesComments, string TradingComments, string CouponScenario,
            string PutStrike1Summary, string PutStrike2Summary, string ExportPutStrike1Summary, string ExportPutStrike2Summary, string Entity, string IsSecured)
        {
            try
            {
                string strTemplateFilePath = Server.MapPath("~/Templates");
                string strTemplateFileName = strTemplateFilePath + "//PutBinaryTemplateWorkingFile.xlsx";

                string strTargetFilePath = System.Configuration.ConfigurationManager.AppSettings["WorkingFilePath"];
                string strTargetFileName = strTargetFilePath + "//" + ProductID + "_PutBinary.xlsx";

                string Role = Convert.ToString(Session["Role"]);

                if (System.IO.File.Exists(strTargetFileName))
                    System.IO.File.Delete(strTargetFileName);

                FileInfo objTemplateFileInfo = new FileInfo(strTemplateFileName);
                objTemplateFileInfo.CopyTo(strTargetFileName);

                FileInfo objTargetFileInfo = new FileInfo(strTargetFileName);

                using (var xlPackage = new ExcelPackage(objTargetFileInfo))
                {
                    var worksheet = xlPackage.Workbook.Worksheets["PutBinary"];

                    worksheet.Cell(1, 2).Value = ProductID.ToString();
                    worksheet.Cell(1, 4).Value = Distributor.ToString().ToUpper();
                    worksheet.Cell(1, 6).Value = Underlying;

                    worksheet.Cell(2, 2).Formula = "=" + EdelweissBuiltIn.ToString() + "%";
                    worksheet.Cell(2, 4).Formula = "=" + DistributorBuiltIn.ToString() + "%";
                    worksheet.Cell(2, 6).Formula = "=B2+D2+H2";
                    worksheet.Cell(2, 8).Formula = "=" + BuiltInAdjustment.ToString() + "%";

                    if (IsIRR.ToUpper() == "TRUE")
                    {
                        worksheet.Cell(3, 2).Formula = "=" + IRR.ToString() + "%";
                        worksheet.Cell(3, 4).Formula = "=(POWER((1 + B3), (F4 / 365)) - 1)*100%";
                    }
                    else
                    {
                        worksheet.Cell(3, 2).Formula = "=((POWER((1+D3),(12/D4))-1) * 100) %";
                        worksheet.Cell(3, 4).Formula = "=" + FixedCoupon.ToString() + "%";
                    }

                    if (IsMaxCouponIRR.ToUpper() == "TRUE")
                    {
                        worksheet.Cell(3, 6).Formula = "=" + MaxCouponIRR.ToString() + "%";
                        worksheet.Cell(3, 8).Formula = "=(POWER((1 + F3), (F4 / 365)) - 1)*100%";
                    }
                    else
                    {
                        worksheet.Cell(3, 6).Formula = "=((POWER((1+H3),(12/D4))-1) * 100) %";
                        worksheet.Cell(3, 8).Formula = "=" + MaxCoupon.ToString() + "%";
                    }

                    worksheet.Cell(4, 2).Value = OptionTenureMonth.ToString();
                    if (IsRedemptionPeriodMonth.ToUpper() == "TRUE")
                    {
                        worksheet.Cell(4, 4).Formula = RedemptionPeriodMonth;
                        worksheet.Cell(4, 6).Formula = "=ROUND(D4*30.417, 0)";
                    }
                    else
                    {
                        worksheet.Cell(4, 4).Formula = "=ROUND(F4/30.417,2)";
                        worksheet.Cell(4, 6).Formula = RedemptionPeriodDays.ToString();
                    }

                    #region Get Entity Name
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
                    LookupMaster objLookupMasterEntity = EntityList.Find(delegate(LookupMaster oLookupMaster) { return oLookupMaster.LookupID == Convert.ToInt32(Entity); });
                    worksheet.Cell(5, 2).Value = objLookupMasterEntity.LookupDescription;
                    #endregion

                    #region Get Is Secured
                    objLookupResult = null;
                    LookupResultList = null;
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
                    LookupMaster objLookupMasterIsSecured = IsSecuredList.Find(delegate(LookupMaster oLookupMaster) { return oLookupMaster.LookupID == Convert.ToInt32(IsSecured); });
                    worksheet.Cell(5, 4).Value = objLookupMasterIsSecured.LookupDescription;
                    #endregion

                    worksheet.Cell(5, 6).Formula = "=" + DeploymentRate.ToString() + "%";

                    if (CustomerDeploymentRate == "")
                        CustomerDeploymentRate = "0";
                    worksheet.Cell(5, 8).Formula = "=" + CustomerDeploymentRate.ToString() + "%";

                    worksheet.Cell(6, 2).Formula = "=(100-(B2+D2)*100)-(100*(1+ROUND(D3,4)))/(POWER((1+(IF(H5>0,H5,F5))),(D4/12)))";
                    worksheet.Cell(6, 4).Formula = "=G11";
                    worksheet.Cell(6, 6).Formula = "=B6 + D6";

                    worksheet.Cell(8, 2).Value = InitialAveragingMonth.ToString();
                    worksheet.Cell(8, 4).Value = InitialAveragingDaysDiff.ToString();
                    worksheet.Cell(8, 6).Value = FinalAveragingMonth.ToString();
                    worksheet.Cell(8, 8).Value = FinalAveragingDaysDiff.ToString();

                    worksheet.Cell(11, 1).Value = "Put Binary Short";
                    worksheet.Cell(11, 2).Value = Underlying;
                    worksheet.Cell(11, 3).Value = PutStrike1.ToString();
                    worksheet.Cell(11, 4).Formula = "=ROUND(D3,4)-ROUND(H3,4)";
                    worksheet.Cell(11, 5).Formula = "=(AVERAGE(INDIRECT(\"$Q$2:\"&ADDRESS(1+$B$8,16+$F$8))))-(AVERAGE(INDIRECT(\"$Y$2:\"&ADDRESS(1+$B$8,24+$F$8))))";
                    worksheet.Cell(11, 6).Formula = "=E11 * (1 / POWER((1 + ((1 + IF(H5>0,H5,F5)) / (1 + IF(K11>0,K11,J11)) - 1)), (D4 / 12)))";
                    worksheet.Cell(11, 7).Formula = "=F11 * (D11*100) / (C11 - (C11-5))";

                    worksheet.Cell(11, 8).Formula = "=" + PutIV1.ToString() + "%";

                    if (PutCustomIV1 == "")
                        PutCustomIV1 = "0";
                    worksheet.Cell(11, 9).Formula = "=" + PutCustomIV1.ToString() + "%";

                    worksheet.Cell(11, 10).Formula = "=" + PutRF1.ToString() + "%";

                    if (PutCustomRF1 == "")
                        PutCustomRF1 = "0";
                    worksheet.Cell(11, 11).Formula = "=" + PutCustomRF1.ToString() + "%";

                    worksheet.Cell(11, 12).Formula = "=" + PutIV2.ToString() + "%";

                    if (PutCustomIV2 == "")
                        PutCustomIV2 = "0";
                    worksheet.Cell(11, 13).Formula = "=" + PutCustomIV2.ToString() + "%";

                    worksheet.Cell(11, 14).Formula = "=" + PutRF2.ToString() + "%";

                    if (PutCustomRF2 == "")
                        PutCustomRF2 = "0";
                    worksheet.Cell(11, 15).Formula = "=" + PutCustomRF2.ToString() + "%";

                    if (SalesComments != null)
                        worksheet.Cell(13, 2).Value = SalesComments.ToString();
                    else
                        worksheet.Cell(13, 2).Value = "";

                    if (TradingComments != null)
                        worksheet.Cell(14, 2).Value = TradingComments.ToString();
                    else
                        worksheet.Cell(14, 2).Value = "";

                    if (CouponScenario != null)
                        worksheet.Cell(15, 2).Value = CouponScenario.ToString();
                    else
                        worksheet.Cell(15, 2).Value = "";

                    //---------------Write Strike 1 IV Grid-----------------START------------
                    worksheet.Cell(1, 16).Formula = "=C11";
                    worksheet.Cell(1, 17).Formula = "=ROUND($B$4*30.417,0)";
                    worksheet.Cell(1, 18).Formula = "=ROUND($B$4*30.417,0) - (1*$H$8)";
                    worksheet.Cell(1, 19).Formula = "=ROUND($B$4*30.417,0) - (2*$H$8)";
                    worksheet.Cell(1, 20).Formula = "=ROUND($B$4*30.417,0) - (3*$H$8)";
                    worksheet.Cell(1, 21).Formula = "=ROUND($B$4*30.417,0) - (4*$H$8)";
                    worksheet.Cell(1, 22).Formula = "=ROUND($B$4*30.417,0) - (5*$H$8)";

                    worksheet.Cell(2, 16).Formula = "0";
                    worksheet.Cell(3, 16).Formula = "=$P$2+1*$D$8";
                    worksheet.Cell(4, 16).Formula = "=$P$2+2*$D$8";
                    worksheet.Cell(5, 16).Formula = "=$P$2+3*$D$8";
                    worksheet.Cell(6, 16).Formula = "=$P$2+4*$D$8";
                    worksheet.Cell(7, 16).Formula = "=$P$2+5*$D$8";

                    worksheet.Cell(2, 17).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,Q1,IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(3, 17).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($Q$1-P3),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(4, 17).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($Q$1-P4),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(5, 17).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($Q$1-P5),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(6, 17).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($Q$1-P6),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(7, 17).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($Q$1-P7),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";

                    worksheet.Cell(2, 18).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,R1,IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(3, 18).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($R$1-P3),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(4, 18).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($R$1-P4),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(5, 18).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($R$1-P5),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(6, 18).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($R$1-P6),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(7, 18).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($R$1-P7),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";

                    worksheet.Cell(2, 19).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,S1,IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(3, 19).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($S$1-P3),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(4, 19).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($S$1-P4),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(5, 19).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($S$1-P5),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(6, 19).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($S$1-P6),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(7, 19).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($S$1-P7),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";

                    worksheet.Cell(2, 20).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,T1,IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(3, 20).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($T$1-P3),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(4, 20).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($T$1-P4),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(5, 20).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($T$1-P5),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(6, 20).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($T$1-P6),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(7, 20).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($T$1-P7),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";

                    worksheet.Cell(2, 21).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,U1,IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(3, 21).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($U$1-P3),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(4, 21).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($U$1-P4),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(5, 21).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($U$1-P5),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(6, 21).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($U$1-P6),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(7, 21).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($U$1-P7),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";

                    worksheet.Cell(2, 22).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,V1,IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(3, 22).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($V$1-P3),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(4, 22).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($V$1-P4),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(5, 22).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($V$1-P5),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(6, 22).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($V$1-P6),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    worksheet.Cell(7, 22).Formula = "=HoadleyOptions2(\"p\",1,\"P\",$C$11,100,($V$1-P7),IF(I11>0,I11,H11),IF(K11>0,K11,J11))";
                    //---------------Write Strike 1 IV Grid-----------------END--------------

                    //---------------Write Strike 2 IV Grid-----------------START------------
                    worksheet.Cell(1, 24).Formula = "=C11-5";
                    worksheet.Cell(1, 25).Formula = "=ROUND($B$4*30.417,0)";
                    worksheet.Cell(1, 26).Formula = "=ROUND($B$4*30.417,0) - (1*$H$8)";
                    worksheet.Cell(1, 27).Formula = "=ROUND($B$4*30.417,0) - (2*$H$8)";
                    worksheet.Cell(1, 28).Formula = "=ROUND($B$4*30.417,0) - (3*$H$8)";
                    worksheet.Cell(1, 29).Formula = "=ROUND($B$4*30.417,0) - (4*$H$8)";
                    worksheet.Cell(1, 30).Formula = "=ROUND($B$4*30.417,0) - (5*$H$8)";

                    worksheet.Cell(2, 24).Formula = "0";
                    worksheet.Cell(3, 24).Formula = "=$X$2+1*$D$8";
                    worksheet.Cell(4, 24).Formula = "=$X$2+2*$D$8";
                    worksheet.Cell(5, 24).Formula = "=$X$2+3*$D$8";
                    worksheet.Cell(6, 24).Formula = "=$X$2+4*$D$8";
                    worksheet.Cell(7, 24).Formula = "=$X$2+5*$D$8";

                    worksheet.Cell(2, 25).Formula = "=HoadleyOptions2(\"p\",1,\"P\",($C$11-5),100,Y1,IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(3, 25).Formula = "=HoadleyOptions2(\"p\",1,\"P\",($C$11-5),100,($Y$1-X3),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(4, 25).Formula = "=HoadleyOptions2(\"p\",1,\"P\",($C$11-5),100,($Y$1-X4),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(5, 25).Formula = "=HoadleyOptions2(\"p\",1,\"P\",($C$11-5),100,($Y$1-X5),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(6, 25).Formula = "=HoadleyOptions2(\"p\",1,\"P\",($C$11-5),100,($Y$1-X6),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(7, 25).Formula = "=HoadleyOptions2(\"p\",1,\"P\",($C$11-5),100,($Y$1-X7),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";

                    worksheet.Cell(2, 26).Formula = "=HoadleyOptions2(\"p\",1,\"P\",($C$11-5),100,Z1,IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(3, 26).Formula = "=HoadleyOptions2(\"p\",1,\"P\",($C$11-5),100,($Z$1-X3),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(4, 26).Formula = "=HoadleyOptions2(\"p\",1,\"P\",($C$11-5),100,($Z$1-X4),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(5, 26).Formula = "=HoadleyOptions2(\"p\",1,\"P\",($C$11-5),100,($Z$1-X5),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(6, 26).Formula = "=HoadleyOptions2(\"p\",1,\"P\",($C$11-5),100,($Z$1-X6),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(7, 26).Formula = "=HoadleyOptions2(\"p\",1,\"P\",($C$11-5),100,($Z$1-X7),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";

                    worksheet.Cell(2, 27).Formula = "=HoadleyOptions2(\"p\",1,\"P\",($C$11-5),100,AA1,IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(3, 27).Formula = "=HoadleyOptions2(\"p\",1,\"P\",($C$11-5),100,($AA$1-X3),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(4, 27).Formula = "=HoadleyOptions2(\"p\",1,\"P\",($C$11-5),100,($AA$1-X4),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(5, 27).Formula = "=HoadleyOptions2(\"p\",1,\"P\",($C$11-5),100,($AA$1-X5),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(6, 27).Formula = "=HoadleyOptions2(\"p\",1,\"P\",($C$11-5),100,($AA$1-X6),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(7, 27).Formula = "=HoadleyOptions2(\"p\",1,\"P\",($C$11-5),100,($AA$1-X7),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";

                    worksheet.Cell(2, 28).Formula = "=HoadleyOptions2(\"p\",1,\"P\",($C$11-5),100,AB1,IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(3, 28).Formula = "=HoadleyOptions2(\"p\",1,\"P\",($C$11-5),100,($AB$1-X3),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(4, 28).Formula = "=HoadleyOptions2(\"p\",1,\"P\",($C$11-5),100,($AB$1-X4),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(5, 28).Formula = "=HoadleyOptions2(\"p\",1,\"P\",($C$11-5),100,($AB$1-X5),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(6, 28).Formula = "=HoadleyOptions2(\"p\",1,\"P\",($C$11-5),100,($AB$1-X6),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(7, 28).Formula = "=HoadleyOptions2(\"p\",1,\"P\",($C$11-5),100,($AB$1-X7),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";

                    worksheet.Cell(2, 29).Formula = "=HoadleyOptions2(\"p\",1,\"P\",($C$11-5),100,AC1,IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(3, 29).Formula = "=HoadleyOptions2(\"p\",1,\"P\",($C$11-5),100,($AC$1-X3),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(4, 29).Formula = "=HoadleyOptions2(\"p\",1,\"P\",($C$11-5),100,($AC$1-X4),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(5, 29).Formula = "=HoadleyOptions2(\"p\",1,\"P\",($C$11-5),100,($AC$1-X5),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(6, 29).Formula = "=HoadleyOptions2(\"p\",1,\"P\",($C$11-5),100,($AC$1-X6),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(7, 29).Formula = "=HoadleyOptions2(\"p\",1,\"P\",($C$11-5),100,($AC$1-X7),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";

                    worksheet.Cell(2, 30).Formula = "=HoadleyOptions2(\"p\",1,\"P\",($C$11-5),100,AD1,IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(3, 30).Formula = "=HoadleyOptions2(\"p\",1,\"P\",($C$11-5),100,($AD$1-X3),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(4, 30).Formula = "=HoadleyOptions2(\"p\",1,\"P\",($C$11-5),100,($AD$1-X4),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(5, 30).Formula = "=HoadleyOptions2(\"p\",1,\"P\",($C$11-5),100,($AD$1-X5),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(6, 30).Formula = "=HoadleyOptions2(\"p\",1,\"P\",($C$11-5),100,($AD$1-X6),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    worksheet.Cell(7, 30).Formula = "=HoadleyOptions2(\"p\",1,\"P\",($C$11-5),100,($AD$1-X7),IF(M11>0,M11,L11),IF(O11>0,O11,N11))";
                    //---------------Write Strike 2 IV Grid-----------------END--------------

                    xlPackage.Save();
                }

                return Json("");
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];

                ClearPutBinarySession();

                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "ExportPutBinary", objUserMaster.UserID);

                return Json("");
            }
        }

        public JsonResult ManagePutBinary(string ProductID, string Distributor, string EdelweissBuiltIn, string DistributorBuiltIn, string BuiltInAdjustment, string TotalBuiltIn, string FixedCoupon, string IRR, string IsIRR, string MaxCoupon,
            string MaxCouponIRR, string IsMaxCouponIRR, string DeploymentRate, string CustomerDeploymentRate, string Remaining, string Underlying, string TotalOptionPrice, string NetRemaining,
            string OptionTenureMonth, string RedemptionPeriodMonth, string IsRedemptionPeriodMonth, string RedemptionPeriodDays, string InitialAveragingMonth, string InitialAveragingDaysDiff, string FinalAveragingMonth,
            string FinalAveragingDaysDiff, string PutBinaryLong, string PutUnderlying, string PutStrike1, string PutStrike2, string PutCouponFall, string PutPrice, string PutDiscountedPrice,
            string PutPRAdjustedPrice, string PutIV1, string PutCustomIV1, string PutRF1, string PutCustomRF1, string PutIV2, string PutCustomIV2, string PutRF2,
            string PutCustomRF2, string SalesComments, string TradingComments, string CouponScenario, string CopyProductID,
            string PutStrike1Summary, string PutStrike2Summary, string ExportPutStrike1Summary, string ExportPutStrike2Summary, string Entity, string IsSecured)
        {
            try
            {
                PutStrike1Summary = System.Uri.UnescapeDataString(PutStrike1Summary);
                PutStrike2Summary = System.Uri.UnescapeDataString(PutStrike2Summary);
                ExportPutStrike1Summary = System.Uri.UnescapeDataString(ExportPutStrike1Summary);
                ExportPutStrike2Summary = System.Uri.UnescapeDataString(ExportPutStrike2Summary);

                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];

                if (CustomerDeploymentRate == "")
                    CustomerDeploymentRate = "0";

                if (TotalOptionPrice == "")
                    TotalOptionPrice = "0";

                if (NetRemaining == "")
                    NetRemaining = "0";

                if (PutStrike1 == "")
                    PutStrike1 = "0";

                if (PutCouponFall == "")
                    PutCouponFall = "0";

                if (PutPrice == "")
                    PutPrice = "0";

                if (PutDiscountedPrice == "")
                    PutDiscountedPrice = "0";

                if (PutPRAdjustedPrice == "")
                    PutPRAdjustedPrice = "0";

                if (PutIV1 == "")
                    PutIV1 = "0";

                if (PutCustomIV1 == "")
                    PutCustomIV1 = "0";

                if (PutRF1 == "")
                    PutRF1 = "0";

                if (PutCustomRF1 == "")
                    PutCustomRF1 = "0";

                if (PutIV2 == "")
                    PutIV2 = "0";

                if (PutCustomIV2 == "")
                    PutCustomIV2 = "0";

                if (PutRF2 == "")
                    PutRF2 = "0";

                if (PutCustomRF2 == "")
                    PutCustomRF2 = "0";

                string ParentProductID = "";
                if (Session["ParentProductID"] != null)
                    ParentProductID = (string)Session["ParentProductID"];

                ObjectResult<ManagePutBinaryResult> objManagePutBinaryResult = objSP_PRICINGEntities.SP_MANAGE_PUT_BINARY_DETAILS(ProductID, ParentProductID, Distributor, Convert.ToDouble(EdelweissBuiltIn),
                        Convert.ToDouble(DistributorBuiltIn), Convert.ToDouble(BuiltInAdjustment), Convert.ToDouble(TotalBuiltIn), Convert.ToDouble(FixedCoupon), Convert.ToDouble(IRR), Convert.ToBoolean(IsIRR), Convert.ToDouble(MaxCoupon), Convert.ToDouble(MaxCouponIRR), Convert.ToBoolean(IsMaxCouponIRR),
                        Convert.ToDouble(DeploymentRate), Convert.ToDouble(CustomerDeploymentRate), Convert.ToDouble(Remaining), Convert.ToInt32(Underlying), Convert.ToDouble(TotalOptionPrice),
                        Convert.ToDouble(NetRemaining), Convert.ToInt32(OptionTenureMonth), Convert.ToDouble(RedemptionPeriodMonth), Convert.ToBoolean(IsRedemptionPeriodMonth), Convert.ToInt32(RedemptionPeriodDays), Convert.ToInt32(InitialAveragingMonth),
                        Convert.ToInt32(InitialAveragingDaysDiff), Convert.ToInt32(FinalAveragingMonth), Convert.ToInt32(FinalAveragingDaysDiff), SalesComments, TradingComments, CouponScenario, Convert.ToInt32(Entity), Convert.ToInt32(IsSecured), objUserMaster.UserID,
                        PutBinaryLong, Convert.ToDouble(PutStrike1), Convert.ToDouble(PutStrike2), Convert.ToDouble(PutCouponFall), Convert.ToDouble(PutPrice), Convert.ToDouble(PutDiscountedPrice),
                        Convert.ToDouble(PutPRAdjustedPrice), Convert.ToDouble(PutIV1), Convert.ToDouble(PutCustomIV1), Convert.ToDouble(PutRF1), Convert.ToDouble(PutCustomRF1),
                        Convert.ToDouble(PutIV2), Convert.ToDouble(PutCustomIV2), Convert.ToDouble(PutRF2), Convert.ToDouble(PutCustomRF2), CopyProductID, ExportPutStrike1Summary, PutStrike1Summary,
                        ExportPutStrike2Summary, PutStrike2Summary);
                List<ManagePutBinaryResult> ManagePutBinaryResultList = objManagePutBinaryResult.ToList();

                Session.Remove("ParentProductID");

                return Json(ManagePutBinaryResultList[0].ProductID);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];

                ClearPutBinarySession();

                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "ManagePutBinary", objUserMaster.UserID);
                return Json("");// return RedirectToAction("ErrorPage", "Login");
            }
        }
        #endregion

        public ActionResult Generic()
        {
            return View();
        }

        public List<Graph> GenerateGraphCalculation(Int32 Strike, Int32 BelowStrikeCoupon, Int32 AfterStrikeCoupon, Int32 RedemptionDays)
        {
            var transactionCounts = new List<Graph>();
            try
            {
                DataTable dtGraph = new DataTable();

                dtGraph.Columns.Add("INITIAL");
                dtGraph.Columns.Add("FINAL");
                dtGraph.Columns.Add("NIFTY_PERFORMANCE");
                dtGraph.Columns.Add("COUPON");
                dtGraph.Columns.Add("ANNUALIZED_RETURN");

                int Count = 23;
                if (Strike % 10 != 0)
                {
                    Count = Count + 1;
                }

                DataRow dr;
                var Initial = 100;
                var Nifty = -110;
                var Final = 0;
                var PlotFinal = 0.0;
                var PlotNifty = 0.0;
                bool flag = false;
                bool strike = false;
                //var transactionCounts = new List<Graph>();
                for (int i = 1; i <= Count; i++)
                {
                    dr = dtGraph.NewRow();

                    Nifty = Nifty + 10;
                    Final = (int)(Initial + ((100 * Nifty) / 100));

                    dr["INITIAL"] = Initial;

                    if (Final == Strike)
                    {
                        if (flag == false)
                        {
                            PlotFinal = Convert.ToDouble(Final) - 0.1;
                            PlotNifty = Convert.ToDouble(Nifty) - 0.1;
                            dr["FINAL"] = PlotFinal;
                            dr["NIFTY_PERFORMANCE"] = PlotNifty;
                            Nifty = Nifty - 10;
                        }
                        else
                        {
                            if (strike == true)
                            {
                                PlotFinal = Convert.ToDouble(Final) + 0.1;
                                PlotNifty = Convert.ToDouble(Nifty) + 0.1;
                                dr["FINAL"] = PlotFinal;
                                dr["NIFTY_PERFORMANCE"] = PlotNifty;
                            }
                            else
                            {
                                dr["FINAL"] = Final;
                                dr["NIFTY_PERFORMANCE"] = Nifty;
                                Nifty = Nifty - 10;
                                strike = true;
                            }
                        }
                        flag = true;
                    }
                    else
                    {
                        dr["FINAL"] = Final;
                        dr["NIFTY_PERFORMANCE"] = Nifty;
                    }
                    if (strike == true)
                    {
                        if (PlotFinal <= Strike)
                        {
                            dr["COUPON"] = BelowStrikeCoupon;
                            dr["ANNUALIZED_RETURN"] = Math.Pow((1 + BelowStrikeCoupon), (365 * 1.000 / RedemptionDays * 1.000)) + 1;
                        }
                        else
                        {
                            dr["COUPON"] = AfterStrikeCoupon;
                            dr["ANNUALIZED_RETURN"] = Math.Pow((1 + AfterStrikeCoupon), (365 * 1.000 / RedemptionDays * 1.000)) + 1;
                        }
                    }
                    else
                    {
                        if (Final <= Strike)
                        {
                            dr["COUPON"] = BelowStrikeCoupon;
                            dr["ANNUALIZED_RETURN"] = Math.Pow((1 + BelowStrikeCoupon), (365 * 1.000 / RedemptionDays * 1.000)) + 1;
                        }
                        else
                        {
                            dr["COUPON"] = AfterStrikeCoupon;
                            dr["ANNUALIZED_RETURN"] = Math.Pow((1 + AfterStrikeCoupon), (365 * 1.000 / RedemptionDays * 1.000)) + 1;
                        }
                    }
                    dtGraph.Rows.Add(dr);

                    transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["COUPON"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });
                }

                return transactionCounts;
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "GenerateGraphCalculation", objUserMaster.UserID);
                return transactionCounts;//Json("");// return RedirectToAction("ErrorPage", "Login");
            }


        }

        public DataTable GenerateRow(DataTable dtGraph, double Strike, double Nifty, double BelowCoupon, double AfterCoupon, Int32 RedemptionDays, List<Graph> transactionCounts)
        {
            DataRow dr;
            var Initial = 100;

            //Previous row
            dr = dtGraph.NewRow();
            try
            {
                dr["INITIAL"] = Initial;
                var final = Strike - 0.01;
                dr["FINAL"] = final;//(double)(Initial + ((100 * Nifty) / 100)) - 0.01;
                dr["NIFTY_PERFORMANCE"] = Math.Round(final - Initial, 2);//Nifty - 0.01;

                dr["COUPON"] = BelowCoupon;
                dr["ANNUALIZED_RETURN"] = Math.Pow((1 + BelowCoupon), (365 * 1.000 / RedemptionDays * 1.000)) + 1;
                dtGraph.Rows.Add(dr);
                transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["COUPON"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });

                //Main Row
                dr = dtGraph.NewRow();
                dr["INITIAL"] = Initial;
                final = Strike;
                dr["FINAL"] = Strike;//(double)(Initial + ((100 * Nifty) / 100));
                dr["NIFTY_PERFORMANCE"] = Math.Round(final - Initial, 2);

                dr["COUPON"] = BelowCoupon;
                dr["ANNUALIZED_RETURN"] = Math.Pow((1 + BelowCoupon), (365 * 1.000 / RedemptionDays * 1.000)) + 1;
                dtGraph.Rows.Add(dr);
                transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["COUPON"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });

                //Next Row
                dr = dtGraph.NewRow();
                dr["INITIAL"] = Initial;
                final = Strike + 0.01;
                dr["FINAL"] = final;//(double)(Initial + ((100 * Nifty) / 100)) + 0.01;
                dr["NIFTY_PERFORMANCE"] = Math.Round(final - Initial, 2);

                dr["COUPON"] = AfterCoupon;
                dr["ANNUALIZED_RETURN"] = Math.Pow((1 + BelowCoupon), (365 * 1.000 / RedemptionDays * 1.000)) + 1;
                dtGraph.Rows.Add(dr);
                transactionCounts.Add(new Graph() { Column1 = Convert.ToString(dr["NIFTY_PERFORMANCE"]), Column2 = Convert.ToDouble(dr["COUPON"]), Column3 = Convert.ToDouble(dr["NIFTY_PERFORMANCE"]) });

                return dtGraph;
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "GenerateRow", objUserMaster.UserID);
                return dtGraph;//Json("");// return RedirectToAction("ErrorPage", "Login");
            }
        }

        #region Fixed Coupon

        [HttpGet]
        public ActionResult FixedCouponList(string Status)
        {
            LoginController objLoginController = new LoginController();
            try
            {
                if (ValidateSession())
                {
                    FixedCoupon objFixedCoupon = new FixedCoupon();
                    ObjectResult<LookupResult> objLookupResult;
                    List<LookupResult> LookupResultList;
                    List<LookupMaster> StatusList = new List<LookupMaster>();

                    #region Menu Access By on Role

                    Int32 intResult = 0;
                    // bool PPorNonPP = false;

                    UserMaster objUserMaster = new UserMaster();
                    objUserMaster = (UserMaster)Session["LoggedInUser"];

                    var Result = objSP_PRICINGEntities.VALIDATE_MENU_BY_ROLE(objUserMaster.RoleID, "BFCL");
                    intResult = Convert.ToInt32(Result.SingleOrDefault());

                    if (intResult == 0)
                        return RedirectToAction("UserNotAuthorize", "Login");

                    #endregion

                    objLookupResult = objSP_PRICINGEntities.SP_FETCH_LOOKUP_VALUES("SFSM", false);
                    LookupResultList = objLookupResult.ToList();

                    if (LookupResultList != null && LookupResultList.Count > 0)
                    {
                        foreach (var LookupResult in LookupResultList)
                        {
                            LookupMaster objLookupMaster = new LookupMaster();
                            General.ReflectSingleData(objLookupMaster, LookupResult);

                            StatusList.Add(objLookupMaster);
                        }
                    }

                    objFixedCoupon.StatusList = StatusList;

                    //--Set Status--Added by Shweta on 27th May 2016------------START--------------------
                    if (Status != null && Status != "")
                    {
                        LookupMaster objStatus = objFixedCoupon.StatusList.Find(delegate(LookupMaster oLookupMaster) { return oLookupMaster.LookupID == Convert.ToInt32(Status); });
                        objFixedCoupon.FilterStatus = Convert.ToString(objStatus.LookupID);
                    }
                    //--Set Status--Added by Shweta on 27th May 2016------------END----------------------

                    return View(objFixedCoupon);
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
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "FixedCouponList", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        public JsonResult FetchFixedCouponList(string ProductID, string Status, string ProductTenure, string FixedCoupon, string FixedCouponIRR, string Distributor, string FromDate, string ToDate, string SalesComments, string TradingComments)
        {
            try
            {
                List<FixedCoupon> FixedCouponList = new List<FixedCoupon>();

                if (ProductID == "" || ProductID == "--Select--")
                    ProductID = "ALL";

                if (ProductTenure == "" || ProductTenure == "0" || ProductTenure == "--Select--")
                    ProductTenure = "ALL";

                if (FixedCoupon == "" || FixedCoupon == "0" || FixedCoupon == "--Select--")
                    FixedCoupon = "ALL";

                if (FixedCouponIRR == "" || FixedCouponIRR == "--Select--")
                    FixedCouponIRR = "ALL";

                if (Status == "" || Status == "0" || Status == "--Select--")
                    Status = "ALL";

                if (FromDate == "")
                    FromDate = "1900-01-01";
                else
                    FromDate = FromDate.Substring(6, 4) + '-' + FromDate.Substring(0, 2) + '-' + FromDate.Substring(3, 2);

                if (ToDate == "")
                    ToDate = "2900-01-01";
                else
                    ToDate = ToDate.Substring(6, 4) + '-' + ToDate.Substring(0, 2) + '-' + ToDate.Substring(3, 2);

                DateTime dtFromDate = Convert.ToDateTime(FromDate);
                DateTime dtToDate = Convert.ToDateTime(ToDate);

                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("FETCH_FIXED_COUPON", ProductID, Status, ProductTenure, FixedCoupon, FixedCouponIRR, Distributor, dtFromDate, dtToDate, SalesComments, TradingComments);

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in dsResult.Tables[0].Rows)
                    {
                        FixedCoupon obj = new FixedCoupon();

                        obj.ProductID = Convert.ToString(dr["ProductID"]);
                        obj.ConfirmedOn = Convert.ToString(dr["ConfirmedOn"]);
                        obj.ProductTenure = Convert.ToString(dr["ProductTenure"]);
                        obj.FixedCouponIRR = Convert.ToString(dr["FixedCouponIRR"]);
                        obj.Distributor = Convert.ToString(dr["Distributor"]);
                        obj.DistributorBuiltIn = Convert.ToDouble(dr["DistributorBuiltIn"]);
                        obj.EdelweissBuiltIn = Convert.ToDouble(dr["EdelweissBuiltIn"]);
                        obj.Status = Convert.ToString(dr["Status"]);
                        obj.SalesComments = Convert.ToString(dr["SalesComment"]);
                        obj.TradingComments = Convert.ToString(dr["TradingComment"]);
                        obj.CouponScenario = Convert.ToString(dr["CouponScenario"]);
                        obj.IsFavourite = Convert.ToBoolean(dr["IsFavourite"]);

                        FixedCouponList.Add(obj);
                    }
                }

                var FixedCouponListData = FixedCouponList.ToList();
                return Json(FixedCouponListData, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "FetchFixedCouponList", objUserMaster.UserID);
                return Json("");
            }
        }

        public ActionResult AutoCompleteQuoteID(string term)
        {
            try
            {
                List<FixedCoupon> FixedCouponList = new List<FixedCoupon>();

                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("FETCH_FIXED_COUPON", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", Convert.ToDateTime("1900-01-01"), Convert.ToDateTime("2900-01-01"), "ALL", "ALL");

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in dsResult.Tables[0].Rows)
                    {
                        FixedCoupon obj = new FixedCoupon();
                        obj.ProductID = Convert.ToString(dr["ProductID"]);
                        FixedCouponList.Add(obj);
                    }
                }

                var DistinctItems = FixedCouponList.GroupBy(x => x.ProductID).Select(y => y.First());

                var result = (from objRuleList in DistinctItems
                              where objRuleList.ProductID.ToLower().StartsWith(term.ToLower())
                              select objRuleList);

                return Json(result);

            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "AutoCompleteQuoteID", objUserMaster.UserID);
                Session["ErrorData"] = ex.Message;
                return RedirectToAction("ErrorPage", "Login");
            }

        }

        public ActionResult AutoCompleteDistributor(string term)
        {
            try
            {
                List<FixedCoupon> FixedCouponList = new List<FixedCoupon>();

                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("FETCH_FIXED_COUPON", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", Convert.ToDateTime("1900-01-01"), Convert.ToDateTime("2900-01-01"), "ALL", "ALL");

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in dsResult.Tables[0].Rows)
                    {
                        FixedCoupon obj = new FixedCoupon();

                        obj.ProductID = Convert.ToString(dr["ProductID"]);
                        obj.ProductTenure = Convert.ToString(dr["ProductTenure"]);
                        obj.DeploymentRate = Convert.ToDouble(dr["DeploymentRate"]);
                        obj.EdelweissBuiltIn = Convert.ToDouble(dr["EdelweissBuiltIn"]);
                        obj.DistributorBuiltIn = Convert.ToDouble(dr["DistributorBuiltIn"]);
                        obj.CouponScenario = Convert.ToString(dr["CouponScenario"]);
                        obj.Distributor = Convert.ToString(dr["Distributor"]);
                        obj.ConfirmedOn = Convert.ToString(dr["ConfirmedOn"]);

                        FixedCouponList.Add(obj);
                    }
                }

                var DistinctItems = FixedCouponList.GroupBy(x => x.Distributor).Select(y => y.First());

                var result = (from objRuleList in DistinctItems
                              where objRuleList.Distributor.ToLower().StartsWith(term.ToLower())
                              select objRuleList);

                return Json(result);

            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "AutoCompleteDistributor", objUserMaster.UserID);
                Session["ErrorData"] = ex.Message;
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        #endregion

        #region Fixed Coupon MLD
        [HttpGet]
        public ActionResult FixedCouponMLDList(string Status)
        {
            LoginController objLoginController = new LoginController();
            try
            {
                if (ValidateSession())
                {
                    FixedCouponMLD objFixedCouponMLD = new FixedCouponMLD();

                    #region Menu Access By on Role

                    Int32 intResult = 0;
                    // bool PPorNonPP = false;

                    UserMaster objUserMaster = new UserMaster();
                    objUserMaster = (UserMaster)Session["LoggedInUser"];

                    var Result = objSP_PRICINGEntities.VALIDATE_MENU_BY_ROLE(objUserMaster.RoleID, "BFCML");
                    intResult = Convert.ToInt32(Result.SingleOrDefault());

                    if (intResult == 0)
                        return RedirectToAction("UserNotAuthorize", "Login");

                    #endregion

                    ObjectResult<LookupResult> objLookupResult;
                    List<LookupResult> LookupResultList;
                    List<LookupMaster> StatusList = new List<LookupMaster>();

                    objLookupResult = objSP_PRICINGEntities.SP_FETCH_LOOKUP_VALUES("SFSM", false);
                    LookupResultList = objLookupResult.ToList();

                    if (LookupResultList != null && LookupResultList.Count > 0)
                    {
                        foreach (var LookupResult in LookupResultList)
                        {
                            LookupMaster objLookupMaster = new LookupMaster();
                            General.ReflectSingleData(objLookupMaster, LookupResult);

                            StatusList.Add(objLookupMaster);
                        }
                    }

                    objFixedCouponMLD.StatusList = StatusList;

                    //--Set Status--Added by Shweta on 27th May 2016------------START--------------------
                    if (Status != null && Status != "")
                    {
                        LookupMaster objStatus = objFixedCouponMLD.StatusList.Find(delegate(LookupMaster oLookupMaster) { return oLookupMaster.LookupID == Convert.ToInt32(Status); });
                        objFixedCouponMLD.FilterStatus = Convert.ToString(objStatus.LookupID);
                    }
                    //--Set Status--Added by Shweta on 27th May 2016------------END----------------------

                    return View(objFixedCouponMLD);
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
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "FixedCouponMLDList", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        public JsonResult FetchFixedCouponMLDList(string ProductID, string Status, string OptionTenure, string ProductTenure, string Underlying, string Strike1, string CouponAboveIRR, string CouponAboveStrike, string CouponBelowIRR, string CouponBelowStrike, string Distributor, string FromDate, string ToDate, string Sales, string Trading)
        {
            try
            {
                List<FixedCouponMLD> FixedCouponMLDList = new List<FixedCouponMLD>();

                if (ProductID == "")
                    ProductID = "ALL";

                if (ProductTenure == "" || ProductTenure == "0")
                    ProductTenure = "ALL";

                if (Strike1 == "" || Strike1 == "0")
                    Strike1 = "ALL";

                if (Underlying == "" || Underlying == "0" || Underlying == "--Select--")
                    Underlying = "ALL";

                if (OptionTenure == "" || OptionTenure == "0")
                    OptionTenure = "ALL";

                if (Status == "" || Status == "0" || Status == "--Select--")
                    Status = "ALL";

                if (CouponAboveStrike == "" || CouponAboveStrike == "0")
                    CouponAboveStrike = "ALL";

                if (CouponAboveIRR == "")
                    CouponAboveIRR = "ALL";

                if (CouponBelowStrike == "" || CouponBelowStrike == "0")
                    CouponBelowStrike = "ALL";

                if (CouponBelowIRR == "")
                    CouponBelowIRR = "ALL";

                if (FromDate == "")
                    FromDate = "1900-01-01";
                else
                    FromDate = FromDate.Substring(6, 4) + '-' + FromDate.Substring(0, 2) + '-' + FromDate.Substring(3, 2);

                if (ToDate == "")
                    ToDate = "2900-01-01";
                else
                    ToDate = ToDate.Substring(6, 4) + '-' + ToDate.Substring(0, 2) + '-' + ToDate.Substring(3, 2);

                DateTime dtFromDate = Convert.ToDateTime(FromDate);
                DateTime dtToDate = Convert.ToDateTime(ToDate);

                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("FETCH_FIXED_COUPON_MLD", ProductID, Status, OptionTenure, ProductTenure, Underlying, Strike1, CouponAboveStrike, CouponAboveIRR, CouponBelowStrike, CouponBelowIRR, Distributor, dtFromDate, dtToDate, Sales, Trading);

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in dsResult.Tables[0].Rows)
                    {
                        FixedCouponMLD obj = new FixedCouponMLD();

                        obj.ProductID = Convert.ToString(dr["ProductID"]);
                        obj.UnderlyingName = Convert.ToString(dr["UnderlyingID"]);
                        obj.ConfirmedOn = Convert.ToString(dr["ConfirmedOn"]);
                        obj.OptionTenureMonth = Convert.ToInt32(dr["OptionTenureMonth"]);
                        obj.ProductTenure = Convert.ToString(dr["ProductTenure"]);
                        obj.Strike = Convert.ToDouble(dr["Strike"]);
                        obj.CouponBelowIRRValue = Convert.ToString(dr["MinCouponIRR"]);
                        obj.CouponAboveIRRValue = Convert.ToString(dr["MaxCouponIRR"]);
                        obj.Distributor = Convert.ToString(dr["Distributor"]);
                        obj.DistributorBuiltIn = Convert.ToDouble(dr["DistributorBuiltIn"]);
                        obj.EdelweissBuiltIn = Convert.ToDouble(dr["EdelweissBuiltIn"]);
                        obj.Status = Convert.ToString(dr["Status"]);
                        obj.InitialAveragingMonth = Convert.ToInt32(dr["InitialAveragingMonth"]);
                        obj.FinalAveragingMonth = Convert.ToInt32(dr["FinalAveragingMonth"]);
                        obj.SalesComments = Convert.ToString(dr["SalesComment"]);
                        obj.TradingComments = Convert.ToString(dr["TradingComment"]);
                        obj.CouponScenario = Convert.ToString(dr["CouponScenario"]);
                        obj.IsFavourite = Convert.ToBoolean(dr["IsFavourite"]);

                        FixedCouponMLDList.Add(obj);
                    }
                }

                var FixedCouponMLDListData = FixedCouponMLDList.ToList();
                return Json(FixedCouponMLDListData, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "FetchFixedCouponMLDList", objUserMaster.UserID);
                return Json("");//("Index", "ErrorDetails");
            }
        }

        public ActionResult AutoCompleteProductID(string term)
        {
            try
            {
                List<FixedCouponMLD> FixedCouponMLDList = new List<FixedCouponMLD>();

                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("FETCH_FIXED_COUPON_MLD", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "", Convert.ToDateTime("1900-01-01"), Convert.ToDateTime("2900-01-01"), "", "");

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in dsResult.Tables[0].Rows)
                    {
                        FixedCouponMLD obj = new FixedCouponMLD();
                        obj.ProductID = Convert.ToString(dr["ProductID"]);
                        FixedCouponMLDList.Add(obj);
                    }
                }

                var DistinctItems = FixedCouponMLDList.GroupBy(x => x.ProductID).Select(y => y.First());

                var result = (from objRuleList in DistinctItems
                              where objRuleList.ProductID.ToLower().StartsWith(term.ToLower())
                              select objRuleList);

                return Json(result);

            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "AutoCompleteProductID", objUserMaster.UserID);
                Session["ErrorData"] = ex.Message;
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        public ActionResult AutoCompleteDistributorMLD(string term)
        {
            try
            {
                List<FixedCouponMLD> FixedCouponMLDList = new List<FixedCouponMLD>();

                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("FETCH_FIXED_COUPON_MLD", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "", Convert.ToDateTime("1900-01-01"), Convert.ToDateTime("2900-01-01"), "", "");

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in dsResult.Tables[0].Rows)
                    {
                        FixedCouponMLD obj = new FixedCouponMLD();
                        obj.Distributor = Convert.ToString(dr["Distributor"]);
                        FixedCouponMLDList.Add(obj);
                    }
                }

                var DistinctItems = FixedCouponMLDList.GroupBy(x => x.Distributor).Select(y => y.First());

                var result = (from objRuleList in DistinctItems
                              where objRuleList.Distributor.ToLower().StartsWith(term.ToLower())
                              select objRuleList);

                return Json(result);

            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "AutoCompleteDistributorMLD", objUserMaster.UserID);
                Session["ErrorData"] = ex.Message;
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        public ActionResult AutoCompleteUnderlyingIDMLD(string term)
        {
            List<FixedCouponMLD> FixedCouponMLDList = new List<FixedCouponMLD>();
            try
            {

                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("FETCH_FIXED_COUPON_MLD", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "", Convert.ToDateTime("1900-01-01"), Convert.ToDateTime("2900-01-01"), "", "");

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in dsResult.Tables[0].Rows)
                    {
                        FixedCouponMLD obj = new FixedCouponMLD();
                        obj.UnderlyingName = Convert.ToString(dr["UnderlyingID"]);
                        FixedCouponMLDList.Add(obj);
                    }
                }

                var DistinctItems = FixedCouponMLDList.GroupBy(x => x.UnderlyingName).Select(y => y.First());

                var result = (from objRuleList in DistinctItems
                              where objRuleList.UnderlyingName.ToLower().StartsWith(term.ToLower())
                              select objRuleList);

                return Json(result);

            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "AutoCompleteUnderlyingIDMLD", objUserMaster.UserID);
                Session["ErrorData"] = ex.Message;
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        #endregion

        #region Call Binary

        [HttpGet]
        public ActionResult CallBinaryList(string Status)
        {
            LoginController objLoginController = new LoginController();
            try
            {
                if (ValidateSession())
                {
                    CallBinary objCallBinary = new CallBinary();

                    #region Menu Access By on Role

                    Int32 intResult = 0;
                    // bool PPorNonPP = false;

                    UserMaster objUserMaster = new UserMaster();
                    objUserMaster = (UserMaster)Session["LoggedInUser"];

                    var Result = objSP_PRICINGEntities.VALIDATE_MENU_BY_ROLE(objUserMaster.RoleID, "BCBL");
                    intResult = Convert.ToInt32(Result.SingleOrDefault());

                    if (intResult == 0)
                        return RedirectToAction("UserNotAuthorize", "Login");

                    #endregion

                    ObjectResult<LookupResult> objLookupResult;
                    List<LookupResult> LookupResultList;
                    List<LookupMaster> StatusList = new List<LookupMaster>();

                    objLookupResult = objSP_PRICINGEntities.SP_FETCH_LOOKUP_VALUES("SFSM", false);
                    LookupResultList = objLookupResult.ToList();

                    if (LookupResultList != null && LookupResultList.Count > 0)
                    {
                        foreach (var LookupResult in LookupResultList)
                        {
                            LookupMaster objLookupMaster = new LookupMaster();
                            General.ReflectSingleData(objLookupMaster, LookupResult);

                            StatusList.Add(objLookupMaster);
                        }
                    }

                    objCallBinary.StatusList = StatusList;

                    //--Set Status--Added by Shweta on 27th May 2016------------START--------------------
                    if (Status != null && Status != "")
                    {
                        LookupMaster objStatus = objCallBinary.StatusList.Find(delegate(LookupMaster oLookupMaster) { return oLookupMaster.LookupID == Convert.ToInt32(Status); });
                        objCallBinary.FilterStatus = Convert.ToString(objStatus.LookupID);
                    }
                    //--Set Status--Added by Shweta on 27th May 2016------------END----------------------

                    return View(objCallBinary);
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
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "CallBinaryList", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        public JsonResult FetchCallBinaryList(string ProductID, string Status, string OptionTenure, string ProductTenure, string Underlying, string Strike1, string MinIRR, string MinCoupon, string MaxIRR, string MaxCoupon, string Distributor, string FromDate, string ToDate, string Sales, string Trading)
        {
            try
            {
                List<CallBinary> CallBinaryList = new List<CallBinary>();

                if (ProductID == "" || ProductID == "--Select--")
                    ProductID = "ALL";

                if (ProductTenure == "" || ProductTenure == "0" || ProductTenure == "--Select--")
                    ProductTenure = "ALL";

                if (Strike1 == "" || Strike1 == "0" || Strike1 == "--Select--")
                    Strike1 = "ALL";

                if (Underlying == "" || Underlying == "0" || Underlying == "--Select--")
                    Underlying = "ALL";

                if (OptionTenure == "" || OptionTenure == "0" || OptionTenure == "--Select--")
                    OptionTenure = "ALL";

                if (Status == "" || Status == "0" || Status == "--Select--")
                    Status = "ALL";

                if (MinIRR == "")
                    MinIRR = "ALL";

                if (MinCoupon == "" || MinCoupon == "0")
                    MinCoupon = "ALL";

                if (MaxIRR == "")
                    MaxIRR = "ALL";

                if (MaxCoupon == "" || MaxCoupon == "0")
                    MaxCoupon = "ALL";

                if (FromDate == "")
                    FromDate = "1900-01-01";
                else
                    FromDate = FromDate.Substring(6, 4) + '-' + FromDate.Substring(0, 2) + '-' + FromDate.Substring(3, 2);

                if (ToDate == "")
                    ToDate = "2900-01-01";
                else
                    ToDate = ToDate.Substring(6, 4) + '-' + ToDate.Substring(0, 2) + '-' + ToDate.Substring(3, 2);

                DateTime dtFromDate = Convert.ToDateTime(FromDate);
                DateTime dtToDate = Convert.ToDateTime(ToDate);

                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("FETCH_CALL_BINARY", ProductID, Status, OptionTenure, ProductTenure, Underlying, Strike1, MinIRR, MinCoupon, MaxIRR, MaxCoupon, Distributor, dtFromDate, dtToDate, Sales, Trading);

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in dsResult.Tables[0].Rows)
                    {
                        CallBinary obj = new CallBinary();

                        obj.ProductID = Convert.ToString(dr["ProductID"]);
                        obj.UnderlyingName = Convert.ToString(dr["UnderlyingID"]);
                        obj.ConfirmedOn = Convert.ToString(dr["ConfirmedOn"]);
                        obj.OptionTenure = Convert.ToInt32(dr["OptionTenureMonth"]);
                        obj.ProductTenure = Convert.ToString(dr["ProductTenure"]);
                        obj.Strike1 = Convert.ToDouble(dr["Strike1"]);
                        obj.FixedCouponIRRValue = Convert.ToString(dr["MinCouponIRR"]);
                        obj.MaxCouponIRRValue = Convert.ToString(dr["MaxCouponIRR"]);
                        obj.Distributor = Convert.ToString(dr["Distributor"]);
                        obj.DistributorBuiltIn = Convert.ToDouble(dr["DistributorBuiltIn"]);
                        obj.EdelweissBuiltIn = Convert.ToDouble(dr["EdelweissBuiltIn"]);
                        obj.Status = Convert.ToString(dr["Status"]);
                        obj.InitialAveragingMonth = Convert.ToInt32(dr["InitialAveragingMonth"]);
                        obj.FinalAveragingMonth = Convert.ToInt32(dr["FinalAveragingMonth"]);
                        obj.SalesComments = Convert.ToString(dr["SalesComment"]);
                        obj.TradingComments = Convert.ToString(dr["TradingComment"]);
                        obj.CouponScenario = Convert.ToString(dr["CouponScenario"]);
                        obj.IsFavourite = Convert.ToBoolean(dr["IsFavourite"]);

                        CallBinaryList.Add(obj);
                    }
                }

                var FixedCouponMLDListData = CallBinaryList.ToList();
                return Json(FixedCouponMLDListData, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "FetchCallBinaryList", objUserMaster.UserID);
                return Json("");
            }
        }

        public ActionResult AutoCompleteProductIDCallBinary(string term)
        {
            try
            {
                List<CallBinary> FixedCouponList = new List<CallBinary>();

                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("FETCH_CALL_BINARY", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "", Convert.ToDateTime("1900-01-01"), Convert.ToDateTime("2900-01-01"), "", "");

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in dsResult.Tables[0].Rows)
                    {
                        CallBinary obj = new CallBinary();
                        obj.ProductID = Convert.ToString(dr["ProductID"]);
                        FixedCouponList.Add(obj);
                    }
                }

                var DistinctItems = FixedCouponList.GroupBy(x => x.ProductID).Select(y => y.First());

                var result = (from objRuleList in DistinctItems
                              where objRuleList.ProductID.ToLower().StartsWith(term.ToLower())
                              select objRuleList);

                return Json(result);

            }
            catch (Exception ex)
            {
                Session["ErrorData"] = ex.Message;
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "AutoCompleteProductIDCallBinary", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        public ActionResult AutoCompleteDistributorCallBinary(string term)
        {
            try
            {
                List<CallBinary> FixedCouponList = new List<CallBinary>();

                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("FETCH_CALL_BINARY", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "", Convert.ToDateTime("1900-01-01"), Convert.ToDateTime("2900-01-01"), "", "");

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in dsResult.Tables[0].Rows)
                    {
                        CallBinary obj = new CallBinary();
                        obj.Distributor = Convert.ToString(dr["Distributor"]);
                        FixedCouponList.Add(obj);
                    }
                }

                var DistinctItems = FixedCouponList.GroupBy(x => x.Distributor).Select(y => y.First());

                var result = (from objRuleList in DistinctItems
                              where objRuleList.Distributor.ToLower().StartsWith(term.ToLower())
                              select objRuleList);

                return Json(result);

            }
            catch (Exception ex)
            {
                Session["ErrorData"] = ex.Message;
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "AutoCompleteDistributorCallBinary", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        public ActionResult AutoCompleteUnderlyingIDCallBinary(string term)
        {
            try
            {
                List<CallBinary> FixedCouponList = new List<CallBinary>();

                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("FETCH_CALL_BINARY", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "", Convert.ToDateTime("1900-01-01"), Convert.ToDateTime("2900-01-01"), "", "");

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in dsResult.Tables[0].Rows)
                    {
                        CallBinary obj = new CallBinary();
                        obj.UnderlyingName = Convert.ToString(dr["UnderlyingID"]);
                        FixedCouponList.Add(obj);
                    }
                }

                var DistinctItems = FixedCouponList.GroupBy(x => x.UnderlyingName).Select(y => y.First());

                var result = (from objRuleList in DistinctItems
                              where objRuleList.UnderlyingName.ToLower().StartsWith(term.ToLower())
                              select objRuleList);

                return Json(result);

            }
            catch (Exception ex)
            {
                Session["ErrorData"] = ex.Message;
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "AutoCompleteUnderlyingIDCallBinary", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        #endregion

        #region Put Binary

        [HttpGet]
        public ActionResult PutBinaryList(string Status)
        {
            LoginController objLoginController = new LoginController();
            try
            {
                if (ValidateSession())
                {
                    PutBinary objPutBinary = new PutBinary();

                    #region Menu Access By on Role

                    Int32 intResult = 0;
                    // bool PPorNonPP = false;

                    UserMaster objUserMaster = new UserMaster();
                    objUserMaster = (UserMaster)Session["LoggedInUser"];

                    var Result = objSP_PRICINGEntities.VALIDATE_MENU_BY_ROLE(objUserMaster.RoleID, "BPBL");
                    intResult = Convert.ToInt32(Result.SingleOrDefault());

                    if (intResult == 0)
                        return RedirectToAction("UserNotAuthorize", "Login");

                    #endregion

                    ObjectResult<LookupResult> objLookupResult;
                    List<LookupResult> LookupResultList;
                    List<LookupMaster> StatusList = new List<LookupMaster>();

                    objLookupResult = objSP_PRICINGEntities.SP_FETCH_LOOKUP_VALUES("SFSM", false);
                    LookupResultList = objLookupResult.ToList();

                    if (LookupResultList != null && LookupResultList.Count > 0)
                    {
                        foreach (var LookupResult in LookupResultList)
                        {
                            LookupMaster objLookupMaster = new LookupMaster();
                            General.ReflectSingleData(objLookupMaster, LookupResult);

                            StatusList.Add(objLookupMaster);
                        }
                    }

                    objPutBinary.StatusList = StatusList;

                    //--Set Status--Added by Shweta on 27th May 2016------------START--------------------
                    if (Status != null && Status != "")
                    {
                        LookupMaster objStatus = objPutBinary.StatusList.Find(delegate(LookupMaster oLookupMaster) { return oLookupMaster.LookupID == Convert.ToInt32(Status); });
                        objPutBinary.FilterStatus = Convert.ToString(objStatus.LookupID);
                    }
                    //--Set Status--Added by Shweta on 27th May 2016------------END----------------------

                    return View(objPutBinary);
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
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "PutBinaryList", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        public JsonResult FetchPutBinaryList(string ProductID, string Status, string OptionTenure, string ProductTenure, string Underlying, string Strike1, string MinIRR, string MinCoupon, string MaxIRR, string MaxCoupon, string Distributor, string FromDate, string ToDate, string Sales, string Trading)
        {
            try
            {
                List<PutBinary> PutBinaryList = new List<PutBinary>();

                if (ProductID == "" || ProductID == "--Select--")
                    ProductID = "ALL";

                if (ProductTenure == "" || ProductTenure == "0" || ProductTenure == "--Select--")
                    ProductTenure = "ALL";

                if (Strike1 == "" || Strike1 == "0" || Strike1 == "--Select--")
                    Strike1 = "ALL";

                if (Underlying == "" || Underlying == "0" || Underlying == "--Select--")
                    Underlying = "ALL";

                if (OptionTenure == "" || OptionTenure == "0" || OptionTenure == "--Select--")
                    OptionTenure = "ALL";

                if (Status == "" || Status == "0" || Status == "--Select--")
                    Status = "ALL";

                if (MinIRR == "")
                    MinIRR = "ALL";

                if (MinCoupon == "" || MinCoupon == "0")
                    MinCoupon = "ALL";

                if (MaxIRR == "")
                    MaxIRR = "ALL";

                if (MaxCoupon == "" || MaxCoupon == "0")
                    MaxCoupon = "ALL";

                if (FromDate == "")
                    FromDate = "1900-01-01";
                else
                    FromDate = FromDate.Substring(6, 4) + '-' + FromDate.Substring(0, 2) + '-' + FromDate.Substring(3, 2);

                if (ToDate == "")
                    ToDate = "2900-01-01";
                else
                    ToDate = ToDate.Substring(6, 4) + '-' + ToDate.Substring(0, 2) + '-' + ToDate.Substring(3, 2);

                DateTime dtFromDate = Convert.ToDateTime(FromDate);
                DateTime dtToDate = Convert.ToDateTime(ToDate);

                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("FETCH_PUT_BINARY", ProductID, Status, OptionTenure, ProductTenure, Underlying, Strike1, MinIRR, MinCoupon, MaxIRR, MaxCoupon, Distributor, dtFromDate, dtToDate, Sales, Trading);

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in dsResult.Tables[0].Rows)
                    {
                        PutBinary obj = new PutBinary();

                        obj.ProductID = Convert.ToString(dr["ProductID"]);
                        obj.UnderlyingName = Convert.ToString(dr["UnderlyingID"]);
                        obj.ConfirmedOn = Convert.ToString(dr["ConfirmedOn"]);
                        obj.OptionTenure = Convert.ToInt32(dr["OptionTenureMonth"]);
                        obj.ProductTenure = Convert.ToString(dr["ProductTenure"]);
                        obj.Strike1 = Convert.ToDouble(dr["Strike1"]);
                        obj.FixedCouponIRRValue = Convert.ToString(dr["MaxCouponIRR"]);
                        obj.MaxCouponIRRValue = Convert.ToString(dr["MinCouponIRR"]);
                        obj.Distributor = Convert.ToString(dr["Distributor"]);
                        obj.DistributorBuiltIn = Convert.ToDouble(dr["DistributorBuiltIn"]);
                        obj.EdelweissBuiltIn = Convert.ToDouble(dr["EdelweissBuiltIn"]);
                        obj.Status = Convert.ToString(dr["Status"]);
                        obj.InitialAveragingMonth = Convert.ToInt32(dr["InitialAveragingMonth"]);
                        obj.FinalAveragingMonth = Convert.ToInt32(dr["FinalAveragingMonth"]);
                        obj.SalesComments = Convert.ToString(dr["SalesComment"]);
                        obj.TradingComments = Convert.ToString(dr["TradingComment"]);
                        obj.CouponScenario = Convert.ToString(dr["CouponScenario"]);
                        obj.IsFavourite = Convert.ToBoolean(dr["IsFavourite"]);

                        PutBinaryList.Add(obj);
                    }
                }

                var PutBinaryListData = PutBinaryList.ToList();
                return Json(PutBinaryListData, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "FetchPutBinaryList", objUserMaster.UserID);
                return Json("");
            }
        }

        public ActionResult AutoCompleteProductIDPutBinary(string term)
        {
            try
            {
                List<PutBinary> PutBinaryList = new List<PutBinary>();

                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("FETCH_PUT_BINARY", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "", Convert.ToDateTime("1900-01-01"), Convert.ToDateTime("2900-01-01"), "", "");

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in dsResult.Tables[0].Rows)
                    {
                        PutBinary obj = new PutBinary();
                        obj.ProductID = Convert.ToString(dr["ProductID"]);
                        PutBinaryList.Add(obj);
                    }
                }

                var DistinctItems = PutBinaryList.GroupBy(x => x.ProductID).Select(y => y.First());

                var result = (from objRuleList in DistinctItems
                              where objRuleList.ProductID.ToLower().StartsWith(term.ToLower())
                              select objRuleList);

                return Json(result);

            }
            catch (Exception ex)
            {
                Session["ErrorData"] = ex.Message;
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "AutoCompleteProductIDPutBinary", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        public ActionResult AutoCompleteDistributorPutBinary(string term)
        {
            try
            {
                List<PutBinary> PutBinaryList = new List<PutBinary>();

                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("FETCH_PUT_BINARY", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "", Convert.ToDateTime("1900-01-01"), Convert.ToDateTime("2900-01-01"), "", "");

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in dsResult.Tables[0].Rows)
                    {
                        PutBinary obj = new PutBinary();
                        obj.Distributor = Convert.ToString(dr["Distributor"]);
                        PutBinaryList.Add(obj);
                    }
                }

                var DistinctItems = PutBinaryList.GroupBy(x => x.Distributor).Select(y => y.First());

                var result = (from objRuleList in DistinctItems
                              where objRuleList.Distributor.ToLower().StartsWith(term.ToLower())
                              select objRuleList);

                return Json(result);

            }
            catch (Exception ex)
            {
                Session["ErrorData"] = ex.Message;
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "AutoCompleteDistributorPutBinary", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        public ActionResult AutoCompleteUnderlyingIDPutBinary(string term)
        {
            try
            {
                List<PutBinary> PutBinaryList = new List<PutBinary>();

                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("FETCH_PUT_BINARY", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "", Convert.ToDateTime("1900-01-01"), Convert.ToDateTime("2900-01-01"), "", "");

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in dsResult.Tables[0].Rows)
                    {
                        PutBinary obj = new PutBinary();
                        obj.UnderlyingName = Convert.ToString(dr["UnderlyingID"]);
                        PutBinaryList.Add(obj);
                    }
                }

                var DistinctItems = PutBinaryList.GroupBy(x => x.UnderlyingName).Select(y => y.First());

                var result = (from objRuleList in DistinctItems
                              where objRuleList.UnderlyingName.ToLower().StartsWith(term.ToLower())
                              select objRuleList);

                return Json(result);

            }
            catch (Exception ex)
            {
                Session["ErrorData"] = ex.Message;
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "AutoCompleteUnderlyingIDPutBinary", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        #endregion

        #region Golden Cushion
        [HttpGet]
        public ActionResult GoldenCushionList(string Status)
        {
            LoginController objLoginController = new LoginController();
            try
            {
                if (ValidateSession())
                {
                    GoldenCushion objGoldenCushion = new GoldenCushion();

                    #region Menu Access By on Role

                    Int32 intResult = 0;
                    // bool PPorNonPP = false;

                    UserMaster objUserMaster = new UserMaster();
                    objUserMaster = (UserMaster)Session["LoggedInUser"];

                    var Result = objSP_PRICINGEntities.VALIDATE_MENU_BY_ROLE(objUserMaster.RoleID, "BGCL");
                    intResult = Convert.ToInt32(Result.SingleOrDefault());

                    if (intResult == 0)
                        return RedirectToAction("UserNotAuthorize", "Login");

                    #endregion

                    ObjectResult<LookupResult> objLookupResult;
                    List<LookupResult> LookupResultList;
                    List<LookupMaster> StatusList = new List<LookupMaster>();

                    objLookupResult = objSP_PRICINGEntities.SP_FETCH_LOOKUP_VALUES("SFSM", false);
                    LookupResultList = objLookupResult.ToList();

                    if (LookupResultList != null && LookupResultList.Count > 0)
                    {
                        foreach (var LookupResult in LookupResultList)
                        {
                            LookupMaster objLookupMaster = new LookupMaster();
                            General.ReflectSingleData(objLookupMaster, LookupResult);

                            StatusList.Add(objLookupMaster);
                        }
                    }

                    objGoldenCushion.StatusList = StatusList;

                    //--Set Status--Added by Shweta on 27th May 2016------------START--------------------
                    if (Status != null && Status != "")
                    {
                        LookupMaster objStatus = objGoldenCushion.StatusList.Find(delegate(LookupMaster oLookupMaster) { return oLookupMaster.LookupID == Convert.ToInt32(Status); });
                        objGoldenCushion.FilterStatus = Convert.ToString(objStatus.LookupID);
                    }
                    //--Set Status--Added by Shweta on 27th May 2016------------END----------------------

                    return View(objGoldenCushion);
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
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "GoldenCushionList", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        public JsonResult FetchGoldenCushionList(string ProductID, string PP, string Status, string OptionTenure, string ProductTenure, string UnderlyingID, string Participation, string MaxIRR, string MaxCoupon, string MinIRR, string MinCoupon, string Strike1, string Strike2, string Distributor, string FromDate, string ToDate, string Sales, string Trading)
        {
            try
            {
                List<GoldenCushion> GoldenCushionList = new List<GoldenCushion>();

                if (ProductID == "" || ProductID == "--Select--")
                    ProductID = "ALL";

                if (ProductTenure == "" || ProductTenure == "0" || ProductTenure == "--Select--")
                    ProductTenure = "ALL";

                if (Strike1 == "" || Strike1 == "--Select--")
                    Strike1 = "ALL";

                if (Strike2 == "" || Strike2 == "--Select--")
                    Strike2 = "ALL";

                if (UnderlyingID == "" || UnderlyingID == "0" || UnderlyingID == "--Select--")
                    UnderlyingID = "ALL";

                if (OptionTenure == "" || OptionTenure == "0" || OptionTenure == "--Select--")
                    OptionTenure = "ALL";

                if (Status == "" || Status == "0" || Status == "--Select--")
                    Status = "ALL";

                if (Participation == "" || Participation == "0")
                    Participation = "ALL";

                if (MaxIRR == "")
                    MaxIRR = "ALL";

                if (MaxCoupon == "" || MaxCoupon == "0")
                    MaxCoupon = "ALL";

                if (MinCoupon == "" || MinCoupon == "0")
                    MinCoupon = "ALL";

                if (MinIRR == "")
                    MinIRR = "ALL";

                if (FromDate == "")
                    FromDate = "1900-01-01";
                else
                    FromDate = FromDate.Substring(6, 4) + '-' + FromDate.Substring(0, 2) + '-' + FromDate.Substring(3, 2);

                if (ToDate == "")
                    ToDate = "2900-01-01";
                else
                    ToDate = ToDate.Substring(6, 4) + '-' + ToDate.Substring(0, 2) + '-' + ToDate.Substring(3, 2);

                DateTime dtFromDate = Convert.ToDateTime(FromDate);
                DateTime dtToDate = Convert.ToDateTime(ToDate);

                if (PP == "--Select--")
                    PP = "ALL";

                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("FETCH_GOLDEN_CUSHION", ProductID, PP, Status, OptionTenure, ProductTenure, UnderlyingID, Participation, MaxIRR, MaxCoupon, MinIRR, MinCoupon, Strike1, Strike2, Distributor, dtFromDate, dtToDate, Sales, Trading);

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in dsResult.Tables[0].Rows)
                    {
                        GoldenCushion obj = new GoldenCushion();

                        obj.ProductID = Convert.ToString(dr["ProductID"]);
                        obj.UnderlyingName = Convert.ToString(dr["UnderlyingID"]);
                        obj.ConfirmedOn = Convert.ToString(dr["ConfirmedOn"]);
                        obj.OptionTenure = Convert.ToInt32(dr["OptionTenureMonth"]);
                        obj.ProductTenure = Convert.ToString(dr["ProductTenure"]);
                        obj.Strike1 = Convert.ToDouble(dr["Strike1"]);
                        obj.Strike2 = Convert.ToDouble(dr["Strike2"]);
                        obj.FixedCouponIRRValue = Convert.ToString(dr["MaxCouponIRR"]);
                        obj.LowerCouponIRRValue = Convert.ToString(dr["MinCouponIRR"]);
                        obj.Participation = Convert.ToDouble(dr["Participation"]);
                        obj.Distributor = Convert.ToString(dr["Distributor"]);
                        obj.DistributorBuiltIn = Convert.ToDouble(dr["DistributorBuiltIn"]);
                        obj.EdelweissBuiltIn = Convert.ToDouble(dr["EdelweissBuiltIn"]);
                        obj.Status = Convert.ToString(dr["Status"]);
                        obj.InitialAveragingMonth = Convert.ToInt32(dr["InitialAveragingMonth"]);
                        obj.FinalAveragingMonth = Convert.ToInt32(dr["FinalAveragingMonth"]);
                        obj.SalesComments = Convert.ToString(dr["SalesComment"]);
                        obj.TradingComments = Convert.ToString(dr["TradingComment"]);
                        obj.CouponScenario1 = Convert.ToString(dr["CouponScenario1"]);
                        obj.IsFavourite = Convert.ToBoolean(dr["IsFavourite"]);

                        GoldenCushionList.Add(obj);
                    }
                }

                var GoldenCushionListData = GoldenCushionList.ToList();
                return Json(GoldenCushionListData, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                return Json("");
            }
        }

        public ActionResult AutoCompleteProductIDGolden(string term)
        {
            try
            {
                List<GoldenCushion> GoldenCushionList = new List<GoldenCushion>();

                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("FETCH_GOLDEN_CUSHION", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "", Convert.ToDateTime("1900-01-01"), Convert.ToDateTime("2900-01-01"), "", "");

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in dsResult.Tables[0].Rows)
                    {
                        GoldenCushion obj = new GoldenCushion();
                        obj.ProductID = Convert.ToString(dr["ProductID"]);
                        GoldenCushionList.Add(obj);
                    }
                }

                var DistinctItems = GoldenCushionList.GroupBy(x => x.ProductID).Select(y => y.First());

                var result = (from objRuleList in DistinctItems
                              where objRuleList.ProductID.ToLower().StartsWith(term.ToLower())
                              select objRuleList);

                return Json(result);

            }
            catch (Exception ex)
            {
                Session["ErrorData"] = ex.Message;
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "AutoCompleteProductIDGolden", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        public ActionResult AutoCompleteDistributorGolden(string term)
        {
            try
            {
                List<GoldenCushion> GoldenCushionList = new List<GoldenCushion>();

                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("FETCH_GOLDEN_CUSHION", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "", Convert.ToDateTime("1900-01-01"), Convert.ToDateTime("2900-01-01"), "", "");

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in dsResult.Tables[0].Rows)
                    {
                        GoldenCushion obj = new GoldenCushion();
                        obj.Distributor = Convert.ToString(dr["Distributor"]);
                        GoldenCushionList.Add(obj);
                    }
                }

                var DistinctItems = GoldenCushionList.GroupBy(x => x.Distributor).Select(y => y.First());

                var result = (from objRuleList in DistinctItems
                              where objRuleList.Distributor.ToLower().StartsWith(term.ToLower())
                              select objRuleList);

                return Json(result);

            }
            catch (Exception ex)
            {
                Session["ErrorData"] = ex.Message;
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "AutoCompleteDistributorGolden", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        public ActionResult AutoCompleteUnderlyingIDGolden(string term)
        {
            try
            {
                List<GoldenCushion> GoldenCushionList = new List<GoldenCushion>();

                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("FETCH_GOLDEN_CUSHION", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "", Convert.ToDateTime("1900-01-01"), Convert.ToDateTime("2900-01-01"), "", "");

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in dsResult.Tables[0].Rows)
                    {
                        GoldenCushion obj = new GoldenCushion();
                        obj.UnderlyingName = Convert.ToString(dr["UnderlyingID"]);
                        GoldenCushionList.Add(obj);
                    }
                }

                var DistinctItems = GoldenCushionList.GroupBy(x => x.UnderlyingName).Select(y => y.First());

                var result = (from objRuleList in DistinctItems
                              where objRuleList.UnderlyingName.ToLower().StartsWith(term.ToLower())
                              select objRuleList);

                return Json(result);

            }
            catch (Exception ex)
            {
                Session["ErrorData"] = ex.Message;
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "AutoCompleteUnderlyingIDGolden", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        #endregion

        #region Fixed Plus PR
        [HttpGet]
        public ActionResult FixedPlusPRList(string Status)
        {
            LoginController objLoginController = new LoginController();
            try
            {
                if (ValidateSession())
                {
                    FixedPlusPR objFixedPlusPR = new FixedPlusPR();

                    #region Menu Access By on Role

                    Int32 intResult = 0;
                    // bool PPorNonPP = false;

                    UserMaster objUserMaster = new UserMaster();
                    objUserMaster = (UserMaster)Session["LoggedInUser"];

                    var Result = objSP_PRICINGEntities.VALIDATE_MENU_BY_ROLE(objUserMaster.RoleID, "BFPPL");
                    intResult = Convert.ToInt32(Result.SingleOrDefault());

                    if (intResult == 0)
                        return RedirectToAction("UserNotAuthorize", "Login");

                    #endregion

                    ObjectResult<LookupResult> objLookupResult;
                    List<LookupResult> LookupResultList;
                    List<LookupMaster> StatusList = new List<LookupMaster>();

                    objLookupResult = objSP_PRICINGEntities.SP_FETCH_LOOKUP_VALUES("SFSM", false);
                    LookupResultList = objLookupResult.ToList();

                    if (LookupResultList != null && LookupResultList.Count > 0)
                    {
                        foreach (var LookupResult in LookupResultList)
                        {
                            LookupMaster objLookupMaster = new LookupMaster();
                            General.ReflectSingleData(objLookupMaster, LookupResult);

                            StatusList.Add(objLookupMaster);
                        }
                    }

                    objFixedPlusPR.StatusList = StatusList;

                    //--Set Status--Added by Shweta on 27th May 2016------------START--------------------
                    if (Status != null && Status != "")
                    {
                        LookupMaster objStatus = objFixedPlusPR.StatusList.Find(delegate(LookupMaster oLookupMaster) { return oLookupMaster.LookupID == Convert.ToInt32(Status); });
                        objFixedPlusPR.FilterStatus = Convert.ToString(objStatus.LookupID);
                    }
                    //--Set Status--Added by Shweta on 27th May 2016------------END----------------------

                    return View(objFixedPlusPR);
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
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "FixedPlusPRList", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        public JsonResult FetchFixedPlusPRList(string ProductID, string PP, string Status, string OptionTenure, string ProductTenure, string Underlying, string QuoteType, string Strike1, string Strike2, string FixedCoupon, string FixedCouponIRR, string Participation, string Distributor, string FromDate, string ToDate, string Sales, string Trading)
        {
            try
            {
                List<FixedPlusPR> FixedPlusPRList = new List<FixedPlusPR>();

                if (ProductID == "" || ProductID == "--Select--")
                    ProductID = "ALL";

                if (ProductTenure == "" || ProductTenure == "0" || ProductTenure == "--Select--")
                    ProductTenure = "ALL";

                if (FixedCoupon == "" || FixedCoupon == "0" || FixedCoupon == "--Select--")
                    FixedCoupon = "ALL";

                if (Strike1 == "" || Strike1 == "--Select--")
                    Strike1 = "ALL";

                if (Strike2 == "" || Strike2 == "--Select--")
                    Strike2 = "ALL";

                if (Underlying == "" || Underlying == "0" || Underlying == "--Select--")
                    Underlying = "ALL";

                if (OptionTenure == "" || OptionTenure == "0" || OptionTenure == "--Select--")
                    OptionTenure = "ALL";

                if (Status == "" || Status == "0" || Status == "--Select--")
                    Status = "ALL";

                if (FixedCouponIRR == "" || FixedCouponIRR == "--Select--")
                    FixedCouponIRR = "ALL";

                if (Participation == "" || Participation == "0" || Participation == "--Select--")
                    Participation = "ALL";

                if (PP == "--Select--")
                    PP = "ALL";

                if (FromDate == "")
                    FromDate = "1900-01-01";
                else
                    FromDate = FromDate.Substring(6, 4) + '-' + FromDate.Substring(0, 2) + '-' + FromDate.Substring(3, 2);

                if (ToDate == "")
                    ToDate = "2900-01-01";
                else
                    ToDate = ToDate.Substring(6, 4) + '-' + ToDate.Substring(0, 2) + '-' + ToDate.Substring(3, 2);

                DateTime dtFromDate = Convert.ToDateTime(FromDate);
                DateTime dtToDate = Convert.ToDateTime(ToDate);

                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("FETCH_FIXED_PLUS_PR", ProductID, PP, Status, OptionTenure, ProductTenure, Underlying, QuoteType, Strike1, Strike2, FixedCoupon, FixedCouponIRR, Participation, Distributor, dtFromDate, dtToDate, Sales, Trading);

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in dsResult.Tables[0].Rows)
                    {
                        FixedPlusPR obj = new FixedPlusPR();

                        obj.ProductID = Convert.ToString(dr["ProductID"]);
                        obj.UnderlyingName = Convert.ToString(dr["UnderlyingID"]);
                        obj.ConfirmedOn = Convert.ToString(dr["ConfirmedOn"]);
                        obj.OptionTenureMonth = Convert.ToInt32(dr["OptionTenureMonth"]);
                        obj.ProductTenure = Convert.ToString(dr["ProductTenure"]);
                        obj.Strike1 = Convert.ToDouble(dr["Strike1"]);
                        obj.Strike2 = Convert.ToDouble(dr["Strike2"]);
                        obj.FixedCouponIRRValue = Convert.ToString(dr["FixedCouponIRR"]);
                        obj.Participation = Convert.ToDouble(dr["Participation"]);
                        obj.Distributor = Convert.ToString(dr["Distributor"]);
                        obj.DistributorBuiltIn = Convert.ToDouble(dr["DistributorBuiltIn"]);
                        obj.EdelweissBuiltIn = Convert.ToDouble(dr["EdelweissBuiltIn"]);
                        obj.Status = Convert.ToString(dr["Status"]);
                        obj.InitialAveragingMonth = Convert.ToInt32(dr["InitialAveragingMonth"]);
                        obj.FinalAveragingMonth = Convert.ToInt32(dr["FinalAveragingMonth"]);
                        obj.SalesComments = Convert.ToString(dr["SalesComment"]);
                        obj.TradingComments = Convert.ToString(dr["TradingComment"]);
                        obj.CouponScenario1 = Convert.ToString(dr["CouponScenario1"]);
                        obj.IsFavourite = Convert.ToBoolean(dr["IsFavourite"]);

                        FixedPlusPRList.Add(obj);
                    }
                }

                var FixedPlusPRListData = FixedPlusPRList.ToList();
                return Json(FixedPlusPRListData, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "FetchFixedPlusPRList", objUserMaster.UserID);
                return Json("");
            }
        }

        public ActionResult AutoCompleteProductIDFixedPlus(string term)
        {
            try
            {
                List<FixedPlusPR> FixedPlusPRList = new List<FixedPlusPR>();

                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("FETCH_FIXED_PLUS_PR", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "", Convert.ToDateTime("1900-01-01"), Convert.ToDateTime("2900-01-01"), "", "");

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in dsResult.Tables[0].Rows)
                    {
                        FixedPlusPR obj = new FixedPlusPR();
                        obj.ProductID = Convert.ToString(dr["ProductID"]);
                        FixedPlusPRList.Add(obj);
                    }
                }

                var DistinctItems = FixedPlusPRList.GroupBy(x => x.ProductID).Select(y => y.First());

                var result = (from objRuleList in DistinctItems
                              where objRuleList.ProductID.ToLower().StartsWith(term.ToLower())
                              select objRuleList);

                return Json(result);

            }
            catch (Exception ex)
            {
                Session["ErrorData"] = ex.Message;

                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "AutoCompleteProductIDFixedPlus", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        public ActionResult AutoCompleteDistributorFixedPlus(string term)
        {
            try
            {
                List<FixedPlusPR> FixedPlusPRList = new List<FixedPlusPR>();

                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("FETCH_FIXED_PLUS_PR", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "", Convert.ToDateTime("1900-01-01"), Convert.ToDateTime("2900-01-01"), "", "");

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in dsResult.Tables[0].Rows)
                    {
                        FixedPlusPR obj = new FixedPlusPR();
                        obj.Distributor = Convert.ToString(dr["Distributor"]);
                        FixedPlusPRList.Add(obj);
                    }
                }

                var DistinctItems = FixedPlusPRList.GroupBy(x => x.Distributor).Select(y => y.First());

                var result = (from objRuleList in DistinctItems
                              where objRuleList.Distributor.ToLower().StartsWith(term.ToLower())
                              select objRuleList);

                return Json(result);

            }
            catch (Exception ex)
            {
                Session["ErrorData"] = ex.Message;
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "AutoCompleteDistributorFixedPlus", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        public ActionResult AutoCompleteUnderlyingIDFixedPlus(string term)
        {
            try
            {
                List<FixedPlusPR> FixedPlusPRList = new List<FixedPlusPR>();

                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("FETCH_FIXED_PLUS_PR", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "", Convert.ToDateTime("1900-01-01"), Convert.ToDateTime("2900-01-01"), "", "");

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in dsResult.Tables[0].Rows)
                    {
                        FixedPlusPR obj = new FixedPlusPR();
                        obj.UnderlyingName = Convert.ToString(dr["UnderlyingID"]);
                        FixedPlusPRList.Add(obj);
                    }
                }

                var DistinctItems = FixedPlusPRList.GroupBy(x => x.UnderlyingName).Select(y => y.First());

                var result = (from objRuleList in DistinctItems
                              where objRuleList.UnderlyingName.ToLower().StartsWith(term.ToLower())
                              select objRuleList);

                return Json(result);

            }
            catch (Exception ex)
            {
                Session["ErrorData"] = ex.Message;
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "AutoCompleteUnderlyingIDFixedPlus", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        #endregion

        #region Fixed OR PR
        [HttpGet]
        public ActionResult FixedOrPRList(string Status)
        {
            LoginController objLoginController = new LoginController();
            try
            {
                if (ValidateSession())
                {
                    FixedOrPR objFixedOrPR = new FixedOrPR();

                    #region Menu Access By on Role

                    Int32 intResult = 0;
                    // bool PPorNonPP = false;

                    UserMaster objUserMaster = new UserMaster();
                    objUserMaster = (UserMaster)Session["LoggedInUser"];

                    var Result = objSP_PRICINGEntities.VALIDATE_MENU_BY_ROLE(objUserMaster.RoleID, "BFOPL");
                    intResult = Convert.ToInt32(Result.SingleOrDefault());

                    if (intResult == 0)
                        return RedirectToAction("UserNotAuthorize", "Login");

                    #endregion

                    ObjectResult<LookupResult> objLookupResult;
                    List<LookupResult> LookupResultList;
                    List<LookupMaster> StatusList = new List<LookupMaster>();

                    objLookupResult = objSP_PRICINGEntities.SP_FETCH_LOOKUP_VALUES("SFSM", false);
                    LookupResultList = objLookupResult.ToList();

                    if (LookupResultList != null && LookupResultList.Count > 0)
                    {
                        foreach (var LookupResult in LookupResultList)
                        {
                            LookupMaster objLookupMaster = new LookupMaster();
                            General.ReflectSingleData(objLookupMaster, LookupResult);

                            StatusList.Add(objLookupMaster);
                        }
                    }

                    objFixedOrPR.StatusList = StatusList;

                    //--Set Status--Added by Shweta on 27th May 2016------------START--------------------
                    if (Status != null && Status != "")
                    {
                        LookupMaster objStatus = objFixedOrPR.StatusList.Find(delegate(LookupMaster oLookupMaster) { return oLookupMaster.LookupID == Convert.ToInt32(Status); });
                        objFixedOrPR.FilterStatus = Convert.ToString(objStatus.LookupID);
                    }
                    //--Set Status--Added by Shweta on 27th May 2016------------END----------------------

                    return View(objFixedOrPR);
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
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "FixedOrPRList", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        public JsonResult FetchFixedOrPRList(string ProductID, string PP, string Status, string OptionTenure, string ProductTenure, string Underlying, string QuoteType, string Strike1, string Strike2, string FixedCoupon, string FixedCouponIRR, string Participation, string Distributor, string FromDate, string ToDate, string Sales, string Trading)
        {
            try
            {
                List<FixedOrPR> FixedOrPRList = new List<FixedOrPR>();

                if (ProductID == "" || ProductID == "--Select--")
                    ProductID = "ALL";

                if (ProductTenure == "" || ProductTenure == "0" || ProductTenure == "--Select--")
                    ProductTenure = "ALL";

                if (FixedCoupon == "" || FixedCoupon == "0" || FixedCoupon == "--Select--")
                    FixedCoupon = "ALL";

                if (FixedCouponIRR == "" || FixedCouponIRR == "--Select--")
                    FixedCouponIRR = "ALL";

                if (Strike1 == "" || Strike1 == "--Select--")
                    Strike1 = "ALL";

                if (Participation == "" || Participation == "0" || Participation == "--Select--")
                    Participation = "ALL";

                if (Strike2 == "" || Strike2 == "--Select--")
                    Strike2 = "ALL";

                if (Underlying == "" || Underlying == "0" || Underlying == "--Select--")
                    Underlying = "ALL";

                if (OptionTenure == "" || OptionTenure == "0" || OptionTenure == "--Select--")
                    OptionTenure = "ALL";

                if (Status == "" || Status == "0" || Status == "--Select--")
                    Status = "ALL";

                if (PP == "--Select--")
                    PP = "ALL";

                if (FromDate == "")
                    FromDate = "1900-01-01";
                else
                    FromDate = FromDate.Substring(6, 4) + '-' + FromDate.Substring(0, 2) + '-' + FromDate.Substring(3, 2);

                if (ToDate == "")
                    ToDate = "2900-01-01";
                else
                    ToDate = ToDate.Substring(6, 4) + '-' + ToDate.Substring(0, 2) + '-' + ToDate.Substring(3, 2);

                DateTime dtFromDate = Convert.ToDateTime(FromDate);
                DateTime dtToDate = Convert.ToDateTime(ToDate);

                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("FETCH_FIXED_OR_PR", ProductID, PP, Status, OptionTenure, ProductTenure, Underlying, QuoteType, Strike1, Strike2, FixedCoupon, FixedCouponIRR, Participation, Distributor, dtFromDate, dtToDate, Sales, Trading);

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in dsResult.Tables[0].Rows)
                    {
                        FixedOrPR obj = new FixedOrPR();

                        obj.ProductID = Convert.ToString(dr["ProductID"]);
                        obj.UnderlyingName = Convert.ToString(dr["UnderlyingID"]);
                        obj.ConfirmedOn = Convert.ToString(dr["ConfirmedOn"]);
                        obj.OptionTenureMonth = Convert.ToInt32(dr["OptionTenureMonth"]);
                        obj.ProductTenure = Convert.ToString(dr["ProductTenure"]);
                        obj.Strike1 = Convert.ToDouble(dr["Strike1"]);
                        obj.Strike2 = Convert.ToDouble(dr["Strike2"]);
                        obj.FixedCouponIRRValue = Convert.ToString(dr["FixedCouponIRR"]);
                        obj.Participation = Convert.ToDouble(dr["Participation"]);
                        obj.Distributor = Convert.ToString(dr["Distributor"]);
                        obj.DistributorBuiltIn = Convert.ToDouble(dr["DistributorBuiltIn"]);
                        obj.EdelweissBuiltIn = Convert.ToDouble(dr["EdelweissBuiltIn"]);
                        obj.Status = Convert.ToString(dr["Status"]);
                        obj.InitialAveragingMonth = Convert.ToInt32(dr["InitialAveragingMonth"]);
                        obj.FinalAveragingMonth = Convert.ToInt32(dr["FinalAveragingMonth"]);
                        obj.SalesComments = Convert.ToString(dr["SalesComment"]);
                        obj.TradingComments = Convert.ToString(dr["TradingComment"]);
                        obj.CouponScenario1 = Convert.ToString(dr["CouponScenario1"]);
                        obj.IsFavourite = Convert.ToBoolean(dr["IsFavourite"]);

                        FixedOrPRList.Add(obj);
                    }
                }

                var FixedPlusPRListData = FixedOrPRList.ToList();
                return Json(FixedPlusPRListData, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "FetchFixedOrPRList", objUserMaster.UserID);
                return Json("");
            }
        }

        public ActionResult AutoCompleteProductIDFixedOrPR(string term)
        {
            try
            {
                List<FixedOrPR> FixedOrPRList = new List<FixedOrPR>();

                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("FETCH_FIXED_OR_PR", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "", Convert.ToDateTime("1900-01-01"), Convert.ToDateTime("2900-01-01"), "", "");

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in dsResult.Tables[0].Rows)
                    {
                        FixedOrPR obj = new FixedOrPR();
                        obj.ProductID = Convert.ToString(dr["ProductID"]);
                        FixedOrPRList.Add(obj);
                    }
                }

                var DistinctItems = FixedOrPRList.GroupBy(x => x.ProductID).Select(y => y.First());

                var result = (from objRuleList in DistinctItems
                              where objRuleList.ProductID.ToLower().StartsWith(term.ToLower())
                              select objRuleList);

                return Json(result);

            }
            catch (Exception ex)
            {
                Session["ErrorData"] = ex.Message;
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "AutoCompleteProductIDFixedOrPR", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        public ActionResult AutoCompleteDistributorFixedOrPR(string term)
        {
            try
            {
                List<FixedOrPR> FixedOrPRList = new List<FixedOrPR>();

                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("FETCH_FIXED_OR_PR", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "", Convert.ToDateTime("1900-01-01"), Convert.ToDateTime("2900-01-01"), "", "");

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in dsResult.Tables[0].Rows)
                    {
                        FixedOrPR obj = new FixedOrPR();
                        obj.Distributor = Convert.ToString(dr["Distributor"]);
                        FixedOrPRList.Add(obj);
                    }
                }

                var DistinctItems = FixedOrPRList.GroupBy(x => x.Distributor).Select(y => y.First());

                var result = (from objRuleList in DistinctItems
                              where objRuleList.Distributor.ToLower().StartsWith(term.ToLower())
                              select objRuleList);

                return Json(result);

            }
            catch (Exception ex)
            {
                Session["ErrorData"] = ex.Message;
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "AutoCompleteDistributorFixedOrPR", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        public ActionResult AutoCompleteUnderlyingIDFixedOrPR(string term)
        {
            try
            {
                List<FixedOrPR> FixedOrPRList = new List<FixedOrPR>();

                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("FETCH_FIXED_OR_PR", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "", Convert.ToDateTime("1900-01-01"), Convert.ToDateTime("2900-01-01"), "", "");

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in dsResult.Tables[0].Rows)
                    {
                        FixedOrPR obj = new FixedOrPR();
                        obj.UnderlyingName = Convert.ToString(dr["UnderlyingID"]);
                        FixedOrPRList.Add(obj);
                    }
                }

                var DistinctItems = FixedOrPRList.GroupBy(x => x.UnderlyingName).Select(y => y.First());

                var result = (from objRuleList in DistinctItems
                              where objRuleList.UnderlyingName.ToLower().StartsWith(term.ToLower())
                              select objRuleList);

                return Json(result);

            }
            catch (Exception ex)
            {
                Session["ErrorData"] = ex.Message;
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "AutoCompleteUnderlyingIDFixedOrPR", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        #endregion

        public void ExportStrikeGrid(string StrikeHTML)
        {
            try
            {
                Response.AppendHeader("content-disposition", "attachment;filename=ExportedHtml.xls");
                Response.Charset = "";
                Response.Cache.SetCacheability(HttpCacheability.NoCache);
                Response.ContentType = "application/vnd.ms-excel";
                //this.EnableViewState = false;
                Response.Write(StrikeHTML);
                Response.End();
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "ExportStrikeGrid", objUserMaster.UserID);
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
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                objLoginController.LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "ValidateSession", objUserMaster.UserID);
                return false;
            }
        }

        public void ClearFixedCouponSession()
        {
            Session.Remove("FixedCouponCopyQuote");
            Session.Remove("FixedCouponChildQuote");
            Session.Remove("CancelQuote");
            Session.Remove("IsChildQuote");
        }

        public void ClearFixedCouponMLDSession()
        {
            Session.Remove("FixedCouponMLDCopyQuote");
            Session.Remove("FixedCouponMLDChildQuote");
            Session.Remove("CancelQuote");
            Session.Remove("IsChildQuoteMLD");
        }

        public void ClearFixedPlusPRSession()
        {
            Session.Remove("FixedPlusPRCopyQuote");
            Session.Remove("FixedPlusPRChildQuote");
            Session.Remove("CancelQuote");
            Session.Remove("IsChildQuoteFixedPlus");
        }

        public void ClearFixedOrPRSession()
        {
            Session.Remove("FixedOrPRCopyQuote");
            Session.Remove("FixedOrPRChildQuote");
            Session.Remove("CancelQuote");
            Session.Remove("IsChildQuoteFixedOr");
        }

        public void ClearGoldenCushionSession()
        {
            Session.Remove("GoldenCushionCopyQuote");
            Session.Remove("GoldenCushionChildQuote");
            Session.Remove("CancelQuote");
            Session.Remove("IsChildQuoteGoldenCushion");
        }

        public void ClearCallBinarySession()
        {
            Session.Remove("CallBinaryCopyQuote");
            Session.Remove("CallBinaryChildQuote");
            Session.Remove("CancelQuote");
            Session.Remove("IsChildQuoteCallBinary");
        }

        public void ClearPutBinarySession()
        {
            Session.Remove("PutBinaryCopyQuote");
            Session.Remove("PutBinaryChildQuote");
            Session.Remove("CancelQuote");
            Session.Remove("IsChildQuotePutBinary");
        }

        public double TruncateDecimal(double value, int precision)
        {
            double step = (double)Math.Pow(10, precision);
            int tmp = (int)Math.Truncate(step * value);
            return tmp / step;
        }
    }
}