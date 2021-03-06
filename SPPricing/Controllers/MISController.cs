﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using SPPricing.Models;
using System.Data.Objects;
using System.IO;
using System.Web.UI.WebControls;
using OfficeOpenXml;
using System.Data;
using CRYPTOGRAPHY;
using System.Threading;
using System.Data.SqlClient;

namespace SPPricing.Controllers
{
    public class MISController : Controller
    {
        System.Object lockThis = new System.Object();
        SP_PRICINGEntities objSP_PRICINGEntities = new SP_PRICINGEntities();
        //
        // GET: /MCPricers/

        public ActionResult Index()
        {
            return View();
        }

        [HttpGet]
        public ActionResult Autocall(string ProductID)
        {
            LoginController objLoginController = new LoginController();
            try
            {
                if (ValidateSession())
                {
                    Autocall objAutoCall = new Autocall();

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
                    objAutoCall.UnderlyingList = UnderlyingList;
                    #endregion

                    #region Bind Implied Volatility List
                    List<LookupMaster> LookupMasterList = new List<LookupMaster>();

                    ObjectResult<LookupResult> objLookupResult;
                    List<LookupResult> ImpliedVolatilityList = new List<LookupResult>();

                    objLookupResult = objSP_PRICINGEntities.SP_FETCH_LOOKUP_VALUES("AIV", true);

                    ImpliedVolatilityList = objLookupResult.ToList();

                    foreach (LookupResult oLookupResult in ImpliedVolatilityList)
                    {
                        LookupMaster objLookupMaster = new LookupMaster();
                        General.ReflectSingleData(objLookupMaster, oLookupResult);

                        LookupMasterList.Add(objLookupMaster);
                    }
                    objAutoCall.ImpliedVolatilityList = LookupMasterList;
                    #endregion

                    #region Bind Observation Frequency List
                    List<LookupMaster> LookupMasterListForFrequency = new List<LookupMaster>();

                    ObjectResult<LookupResult> objLookupResultForFrequency;
                    List<LookupResult> ObservationFrequencyList = new List<LookupResult>();

                    objLookupResultForFrequency = objSP_PRICINGEntities.SP_FETCH_LOOKUP_VALUES("OF", true);

                    ObservationFrequencyList = objLookupResultForFrequency.ToList();

                    foreach (LookupResult oLookupResult in ObservationFrequencyList)
                    {
                        LookupMaster objLookupMaster = new LookupMaster();
                        General.ReflectSingleData(objLookupMaster, oLookupResult);

                        LookupMasterListForFrequency.Add(objLookupMaster);
                    }
                    objAutoCall.ObservationFrequencyList = LookupMasterListForFrequency;
                    #endregion

                    if (ProductID != "" && ProductID != null)
                    {
                        ObjectResult<AutoCallEditResult> objAutoCallEditResult = objSP_PRICINGEntities.FETCH_AUTOCALL_EDIT_DETAILS(ProductID);
                        List<AutoCallEditResult> AutoCallEditResultList = objAutoCallEditResult.ToList();

                        General.ReflectSingleData(objAutoCall, AutoCallEditResultList[0]);
                    }

                    if (Session["AutoCallCopyQuote"] != null)
                    {
                        ViewBag.Message = "Quote is copied";
                        objAutoCall = (Autocall)Session["AutoCallCopyQuote"];

                        ObjectResult<AutoCallEditResult> objAutoCallEditResult = objSP_PRICINGEntities.FETCH_AUTOCALL_EDIT_DETAILS("");
                        List<AutoCallEditResult> AutoCallEditResultList = objAutoCallEditResult.ToList();
                        Autocall oAutoCall = new Autocall();
                        if (AutoCallEditResultList != null && AutoCallEditResultList.Count > 0)
                            General.ReflectSingleData(oAutoCall, AutoCallEditResultList[0]);

                        objAutoCall.ParentProductID = objAutoCall.ProductID;
                        objAutoCall.ProductID = "";
                        objAutoCall.Status = oAutoCall.Status;
                        objAutoCall.SaveStatus = oAutoCall.SaveStatus;
                    }

                    else if (Session["AutoCallChildQuote"] != null)
                    {
                        ViewBag.Message = true;
                        objAutoCall = (Autocall)Session["AutoCallChildQuote"];

                        ObjectResult<AutoCallEditResult> objAutoCallEditResult = objSP_PRICINGEntities.FETCH_AUTOCALL_EDIT_DETAILS("");
                        List<AutoCallEditResult> AutoCallEditResultList = objAutoCallEditResult.ToList();
                        Autocall oAutoCall = new Autocall();
                        if (AutoCallEditResultList != null && AutoCallEditResultList.Count > 0)
                            General.ReflectSingleData(oAutoCall, AutoCallEditResultList[0]);

                        objAutoCall.ParentProductID = objAutoCall.ProductID;
                        objAutoCall.ProductID = "";
                        objAutoCall.Status = oAutoCall.Status;
                        objAutoCall.SaveStatus = oAutoCall.SaveStatus;
                    }
                    else if (Session["CancelQuote"] != null)
                    {
                        ViewBag.Message = "Quote is copied";
                        objAutoCall = (Autocall)Session["CancelQuote"];

                        ObjectResult<AutoCallEditResult> objAutoCallEditResult = objSP_PRICINGEntities.FETCH_AUTOCALL_EDIT_DETAILS(objAutoCall.ProductID);
                        List<AutoCallEditResult> AutoCallEditResultList = objAutoCallEditResult.ToList();
                        Autocall oAutoCall = new Autocall();
                        if (AutoCallEditResultList != null && AutoCallEditResultList.Count > 0)
                            General.ReflectSingleData(oAutoCall, AutoCallEditResultList[0]);

                        objAutoCall.Status = oAutoCall.Status;
                        objAutoCall.SaveStatus = oAutoCall.SaveStatus;

                        Session.Remove("CancelQuote");
                    }
                    else
                    {
                        Session.Remove("IsChildQuoteAutoCall");
                        Session.Remove("ParentProductID");
                    }
                    if (Session["AutoCallChildQuote"] == null && Session["AutoCallCopyQuote"] == null)
                        objAutoCall.SaveStatus = "";

                    if (Session["AutoCallCopyQuote"] != null)
                        Session.Remove("AutoCallCopyQuote");

                    if (Session["AutoCallChildQuote"] != null)
                        Session.Remove("AutoCallChildQuote");

                    if (ProductID == null)
                    {
                        objAutoCall.isGraphActive = false;
                        return View(objAutoCall);
                    }
                    else
                    {
                        return View(objAutoCall);
                        //GenerateGoldenCushionGraph(objAutoCall);
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
                LogError(ex.Message, ex.StackTrace, "MCPricersController", "Autocall Get", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        [HttpPost, ValidateInput(false)]
        public ActionResult Autocall(string Command, Autocall objAutocall, FormCollection objFormCollection)
        {
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
                    objAutocall.UnderlyingList = UnderlyingList;
                    #endregion

                    #region Bind Implied Volatility List
                    List<LookupMaster> LookupMasterList = new List<LookupMaster>();

                    ObjectResult<LookupResult> objLookupResult;
                    List<LookupResult> ImpliedVolatilityList = new List<LookupResult>();

                    objLookupResult = objSP_PRICINGEntities.SP_FETCH_LOOKUP_VALUES("AIV", true);

                    ImpliedVolatilityList = objLookupResult.ToList();

                    foreach (LookupResult oLookupResult in ImpliedVolatilityList)
                    {
                        LookupMaster objLookupMaster = new LookupMaster();
                        General.ReflectSingleData(objLookupMaster, oLookupResult);

                        LookupMasterList.Add(objLookupMaster);
                    }
                    objAutocall.ImpliedVolatilityList = LookupMasterList;
                    #endregion

                    #region Bind Observation Frequency List
                    List<LookupMaster> LookupMasterListForFrequency = new List<LookupMaster>();

                    ObjectResult<LookupResult> objLookupResultForFrequency;
                    List<LookupResult> ObservationFrequencyList = new List<LookupResult>();

                    objLookupResultForFrequency = objSP_PRICINGEntities.SP_FETCH_LOOKUP_VALUES("OF", true);

                    ObservationFrequencyList = objLookupResultForFrequency.ToList();

                    foreach (LookupResult oLookupResult in ObservationFrequencyList)
                    {
                        LookupMaster objLookupMaster = new LookupMaster();
                        General.ReflectSingleData(objLookupMaster, oLookupResult);

                        LookupMasterListForFrequency.Add(objLookupMaster);
                    }
                    objAutocall.ObservationFrequencyList = LookupMasterListForFrequency;
                    #endregion

                    if (Command == "FullDump")
                    {
                        objAutocall.Count = -1;
                        ExportScenarioDetails(objAutocall, false);

                        return RedirectToAction("Autocall");
                    }
                    else if (Command == "ScenarioDump")
                    {
                        ExportScenarioDetails(objAutocall, false);

                        return RedirectToAction("Autocall");
                    }
                    else if (Command == "ObservationDump")
                    {
                        objAutocall.Count = -1;
                        ExportScenarioDetails(objAutocall, true);

                        return RedirectToAction("Autocall");
                    }
                    else if (Command == "ExportToExcel")
                    {
                        ExportAutocall(objAutocall, objFormCollection);

                        return RedirectToAction("Autocall");
                    }
                    else if (Command == "CopyQuote")
                    {
                        Session["AutoCallCopyQuote"] = objAutocall;

                        return RedirectToAction("Autocall");
                    }
                    else if (Command == "CreateChildQuote")
                    {
                        Session.Remove("ParentProductID");
                        Session.Remove("IsChildQuoteAutoCall");

                        Session["ParentProductID"] = objAutocall.ProductID;

                        objAutocall.IsChildQuote = true;

                        Session["AutoCallChildQuote"] = objAutocall;
                        Session["IsChildQuoteAutoCall"] = objAutocall.IsChildQuote;

                        return RedirectToAction("Autocall");
                    }
                    else if (Command == "GenerateGraph")
                    {
                        objAutocall.isGraphActive = true;
                        return RedirectToAction("Autocall");
                        // return GenerateGoldenCushionGraph(objAutocall);
                    }
                    else if (Command == "AddNewProduct")
                    {
                        var productID = objAutocall.ProductID;
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
                        Session["CancelQuote"] = objAutocall;

                        return RedirectToAction("FixedCoupon");
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
                LogError(ex.Message, ex.StackTrace, "MCPricersController", "Autocall Post", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        public virtual void ExportAutocall(Autocall objAutocall, FormCollection objFormCollection)
        {
            try
            {
                string strTemplateFilePath = Server.MapPath("~/Templates");
                string strTemplateFileName = strTemplateFilePath + "\\AutocallTemplate.xlsx";

                string strTargetFilePath = Server.MapPath("~/OutputFiles");
                string strTargetFileName = strTargetFilePath + "\\" + objAutocall.ProductID + "_Autocall.xlsx";

                if (System.IO.File.Exists(strTargetFileName))
                    System.IO.File.Delete(strTargetFileName);

                FileInfo objTemplateFileInfo = new FileInfo(strTemplateFileName);
                objTemplateFileInfo.CopyTo(strTargetFileName);

                FileInfo objTargetFileInfo = new FileInfo(strTargetFileName);

                Underlying objUnderlying = objAutocall.UnderlyingList.Find(delegate(Underlying oUnderlying) { return oUnderlying.UnderlyingID == objAutocall.UnderlyingID; });
                LookupMaster objImpliedVolatility = objAutocall.ImpliedVolatilityList.Find(delegate(LookupMaster oLookupMaster) { return oLookupMaster.LookupID == objAutocall.ImpliedVolatilityID; });
                LookupMaster objObservationFrequency = objAutocall.ObservationFrequencyList.Find(delegate(LookupMaster oLookupMaster) { return oLookupMaster.LookupID == objAutocall.ObservationFrequencyID; });

                using (var xlPackage = new ExcelPackage(objTargetFileInfo))
                {
                    var worksheet = xlPackage.Workbook.Worksheets["Autocall"];

                    worksheet.Cell(1, 2).Value = objAutocall.ProductID.ToString();
                    worksheet.Cell(1, 4).Value = objAutocall.Distributor.ToString().ToUpper();
                    worksheet.Cell(1, 6).Value = objUnderlying.UnderlyingShortName;//Convert.ToString(objFormCollection["FilterUnderlying"]);

                    worksheet.Cell(3, 2).Formula = "=" + objAutocall.EdelweissBuiltIn.ToString() + "%";
                    worksheet.Cell(3, 4).Formula = "=" + objAutocall.DistributorBuiltIn.ToString() + "%";

                    worksheet.Cell(4, 2).Formula = "=" + objAutocall.FixedCoupon.ToString() + "%";
                    worksheet.Cell(4, 4).Formula = "=((POWER((1+B4),(12/B7))-1)*100) %";

                    worksheet.Cell(6, 2).Value = objAutocall.OptionTenure.ToString();

                    worksheet.Cell(7, 2).Formula = "=INT(D7/30.417)";
                    worksheet.Cell(7, 4).Value = objAutocall.RedemptionPeriodDays.ToString();

                    worksheet.Cell(9, 2).Formula = "=" + objAutocall.DeploymentRate.ToString() + "%";
                    worksheet.Cell(9, 4).Formula = "=" + objAutocall.CustomDeploymentRate.ToString() + "%";

                    worksheet.Cell(10, 2).Value = objAutocall.IsPrincipalProtected.ToString();
                    worksheet.Cell(10, 4).Formula = "=" + objAutocall.NonPPLevel.ToString() + "%";

                    worksheet.Cell(11, 2).Value = objAutocall.IsDiscountingApplicable.ToString();
                    worksheet.Cell(11, 4).Formula = "=" + objAutocall.RollCost.ToString() + "%";

                    worksheet.Cell(12, 2).Value = objImpliedVolatility.LookupDescription;//Convert.ToString(objFormCollection["FilterImpliedVolatility"]);
                    worksheet.Cell(12, 4).Value = objObservationFrequency.LookupDescription;//Convert.ToString(objFormCollection["FilterObservationFrequency"]);

                    worksheet.Cell(13, 2).Value = objAutocall.ObservationStart.ToString();
                    worksheet.Cell(13, 4).Value = objAutocall.ObservationEnd.ToString();

                    worksheet.Cell(15, 2).Value = objAutocall.EarlyRedemptionPaymentGap.ToString();
                    worksheet.Cell(15, 4).Value = objAutocall.ExpectedTimeToMaturity.ToString();
                    worksheet.Cell(15, 6).Formula = "=" + objAutocall.BondPriceCalculation.ToString() + "%";

                    worksheet.Cell(16, 2).Formula = "=" + objAutocall.InterestRateHit.ToString() + "%";
                    worksheet.Cell(16, 4).Formula = "=" + objAutocall.AverageDeploymentRate.ToString() + "%";
                    worksheet.Cell(16, 6).Formula = "=" + objAutocall.OptionPrice.ToString() + "%";

                    worksheet.Cell(17, 2).Formula = "=" + objAutocall.AutocallLevel.ToString() + "%";
                    worksheet.Cell(17, 4).Formula = "=" + objAutocall.ExpectedBondPrice.ToString() + "%";
                    worksheet.Cell(17, 6).Formula = "=" + objAutocall.TotalBuiltIn.ToString() + "%";

                    //worksheet.Cell(18, 2).Value = objAutocall.VarNonPPLevel.ToString();
                    worksheet.Cell(18, 2).Formula = "=" + objAutocall.CouponIfHit.ToString() + "%";
                    worksheet.Cell(18, 4).Formula = "=" + objAutocall.BondPrice.ToString() + "%";
                    worksheet.Cell(18, 6).Formula = "=" + objAutocall.Remaining.ToString() + "%"; //"=(100-(B3+D3)*100)-(100*(1+B4))/(POWER((1+(IF(D6>0,D6,B6))),(B5/12)))";

                    //worksheet.Cell(19, 2).Value = objAutocall.CouponIfHit.ToString();
                    worksheet.Cell(19, 4).Formula = "=" + objAutocall.InterestRateHitCalculation.ToString() + "%";
                    // worksheet.Cell(19, 6).Value = objAutocall.Coupon.ToString();

                    //worksheet.Cell(21, 2).Value = objAutocall.SalesComments.ToString();
                    //worksheet.Cell(22, 2).Value = objAutocall.TradingComments.ToString();

                    if (objAutocall.SalesComments == null)
                        worksheet.Cell(21, 2).Value = "";
                    else
                        worksheet.Cell(21, 2).Value = objAutocall.SalesComments.ToString();

                    if (objAutocall.TradingComments == null)
                        worksheet.Cell(22, 2).Value = "";
                    else
                        worksheet.Cell(22, 2).Value = objAutocall.TradingComments.ToString(); ;

                    if (objAutocall.CouponScenario1 == null)
                        worksheet.Cell(23, 2).Value = "";
                    else
                        worksheet.Cell(23, 2).Value = objAutocall.CouponScenario1;

                    if (objAutocall.CouponScenario2 == null)
                        worksheet.Cell(24, 2).Value = "";
                    else
                        worksheet.Cell(24, 2).Value = objAutocall.CouponScenario2;

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
                LogError(ex.Message, ex.StackTrace, "MCPricersController", "ExportAutocall", objUserMaster.UserID);
                //return RedirectToAction("ErrorPage", "Login");
            }
        }

        public virtual void ExportScenarioDetails(Autocall objAutocall, bool blnIsObservationDump)
        {
            ObjectResult<SimulationDumpResult> objSimulationDumpResult;
            List<SimulationDumpResult> SimulationDumpResultList = new List<SimulationDumpResult>();

            objSimulationDumpResult = objSP_PRICINGEntities.SP_FETCH_AUTOCALL_SIMULATIONS_DUMP(objAutocall.ProductID, objAutocall.Count, blnIsObservationDump);
            SimulationDumpResultList = objSimulationDumpResult.ToList();

            Response.Clear();
            Response.Buffer = true;
            Response.ContentType = "application/vnd.ms-excel";
            Response.Charset = "";
            System.IO.StringWriter tw = new System.IO.StringWriter();
            System.Web.UI.HtmlTextWriter hw = new System.Web.UI.HtmlTextWriter(tw);

            GridView gv = new GridView();
            gv.DataSource = SimulationDumpResultList;
            gv.DataBind();

            gv.RenderControl(hw);
            Response.AddHeader("Content-Disposition", "attachment;filename=" + objAutocall.ProductID + "_Autocall.xls");
            Response.Write(tw.ToString());
            Response.Flush();
            Response.End();
        }

        public ActionResult Lookback()
        {
            return View();
        }

        public ActionResult Generic()
        {
            return View();
        }

        //public ActionResult ManageAutoCall(string ProductID, string Distributor, string Underlying, string EdelweissBuiltIn, string DistributorBuiltIn, string FixedCoupon,
        //    string IRR, string OptionTenureMonth, string RedemptionPeriodDays, string DeploymentRate,
        //    string CustomerDeploymentRate, string NonPPLevel, string RollCost, string ImpliedVolatility, string ObservationFrequency, string ObservationStart, string ObservationEnd,
        //    string EarlyRedemptionPaymentGap, string ExpectedTimeToMaturity, string BondPrice, string InterestRateHit, string AverageDeploymentRate, string OptionPrice, string AutoCallLevel, string ExpectedBondPrice, string TotalBuiltIn, string NonPPLevel1, string BondPriceCalculation, string Remaining, string CouponIfHit, string InterestRateHitCalculation, string NoOfSimulation, string Count,
        //    string SalesComments, string TradingComments, string CouponScenario1, string CouponScenario2, string IsPrincipalProtected, string IsApplyDiscounting, string CopyProductID)
        //{
        //    try
        //    {
        //        if (ValidateSession())
        //        {
        //            UserMaster objUserMaster = new UserMaster();
        //            objUserMaster = (UserMaster)Session["LoggedInUser"];

        //            double RedemptionPeriodMonth = Convert.ToInt32(RedemptionPeriodDays) / 30.417;

        //            if (EdelweissBuiltIn == "")
        //                EdelweissBuiltIn = "0";

        //            if (DistributorBuiltIn == "")
        //                DistributorBuiltIn = "0";

        //            if (FixedCoupon == "")
        //                FixedCoupon = "0";

        //            if (IRR == "")
        //                IRR = "0";

        //            if (RedemptionPeriodDays == "")
        //                RedemptionPeriodDays = "0";

        //            if (DeploymentRate == "")
        //                DeploymentRate = "0";

        //            if (CustomerDeploymentRate == "")
        //                CustomerDeploymentRate = "0";

        //            if (NonPPLevel == "")
        //                NonPPLevel = "0";

        //            if (RollCost == "")
        //                RollCost = "0";

        //            if (ImpliedVolatility == "")
        //                ImpliedVolatility = "0";

        //            if (ObservationFrequency == "")
        //                ObservationFrequency = "0";

        //            if (EarlyRedemptionPaymentGap == "")
        //                EarlyRedemptionPaymentGap = "0";

        //            if (ExpectedTimeToMaturity == "")
        //                ExpectedTimeToMaturity = "0";

        //            if (BondPrice == "")
        //                BondPrice = "0";

        //            if (InterestRateHit == "")
        //                InterestRateHit = "0";

        //            if (OptionPrice == "")
        //                OptionPrice = "0";

        //            if (AutoCallLevel == "")
        //                AutoCallLevel = "0";

        //            if (ExpectedBondPrice == "")
        //                ExpectedBondPrice = "0";

        //            if (TotalBuiltIn == "")
        //                TotalBuiltIn = "0";

        //            if (NonPPLevel1 == "")
        //                NonPPLevel1 = "0";

        //            if (BondPriceCalculation == "")
        //                BondPriceCalculation = "0";

        //            if (Remaining == "")
        //                Remaining = "0";

        //            if (CouponIfHit == "")
        //                CouponIfHit = "0";

        //            if (InterestRateHitCalculation == "")
        //                InterestRateHitCalculation = "0";

        //            if (NoOfSimulation == "")
        //                NoOfSimulation = "0";

        //            if (Count == "")
        //                Count = "0";

        //            if (IsPrincipalProtected == "")
        //                IsPrincipalProtected = "false";

        //            if (IsApplyDiscounting == "")
        //                IsApplyDiscounting = "false";

        //            string ParentProductID = "";
        //            if (Session["ParentProductID"] != null)
        //                ParentProductID = (string)Session["ParentProductID"];

        //            double dblCoupon = 0;

        //            ObjectResult<ManageAutoCallResult> objManageAutoCallResult = objSP_PRICINGEntities.SP_MANAGE_AUTOCALL_DETAILS(ProductID, ParentProductID, Distributor,
        //                Convert.ToInt32(Underlying), Convert.ToDouble(EdelweissBuiltIn), Convert.ToDouble(DistributorBuiltIn), Convert.ToDouble(FixedCoupon), Convert.ToDouble(IRR),
        //                Convert.ToInt32(OptionTenureMonth), Convert.ToDouble(RedemptionPeriodDays), Convert.ToDouble(RedemptionPeriodMonth), Convert.ToDouble(DeploymentRate),
        //                Convert.ToDouble(CustomerDeploymentRate), Convert.ToDouble(NonPPLevel), Convert.ToDouble(RollCost), Convert.ToDouble(ImpliedVolatility),
        //                Convert.ToDouble(ObservationFrequency), Convert.ToInt32(ObservationStart), Convert.ToInt32(ObservationEnd), Convert.ToDouble(EarlyRedemptionPaymentGap),
        //                Convert.ToDouble(ExpectedTimeToMaturity), Convert.ToDouble(BondPrice), Convert.ToDouble(InterestRateHit), Convert.ToDouble(AverageDeploymentRate),
        //                Convert.ToDouble(OptionPrice), Convert.ToDouble(AutoCallLevel), Convert.ToDouble(ExpectedBondPrice), Convert.ToDouble(TotalBuiltIn), Convert.ToDouble(NonPPLevel1),
        //                Convert.ToDouble(BondPriceCalculation), Convert.ToDouble(Remaining), Convert.ToDouble(CouponIfHit), Convert.ToDouble(InterestRateHitCalculation),
        //                Convert.ToDouble(NoOfSimulation), Convert.ToDouble(Count), dblCoupon, SalesComments, TradingComments, CouponScenario1, CouponScenario2,
        //                Convert.ToBoolean(IsPrincipalProtected), Convert.ToBoolean(IsApplyDiscounting), objUserMaster.UserID, CopyProductID);

        //            List<ManageAutoCallResult> ManageAutoCallResultList = objManageAutoCallResult.ToList();

        //            return Json(ManageAutoCallResultList[0].ProductID);
        //        }
        //        else
        //        {
        //            return RedirectToAction("Login", "Login");
        //        }

        //    }
        //    catch (Exception ex)
        //    {
        //        UserMaster objUserMaster = new UserMaster();
        //        objUserMaster = (UserMaster)Session["LoggedInUser"];
        //        LogError(ex.Message, ex.StackTrace, "MCPricersController", "ManageAutocall", objUserMaster.UserID);
        //        return RedirectToAction("ErrorPage", "Login");
        //    }
        //}

        public JsonResult FetchUnderlyingList()
        {
            List<Underlying> UnderlyingList = new List<Underlying>();

            ObjectResult<UnderlyingListResult> objUnderlyingListResult;
            //List<UnderlyingListResult> UnderlyingListResultList;

            objUnderlyingListResult = objSP_PRICINGEntities.SP_FETCH_UNDERLYING_DETAILS();


            var UnderlyingListData = objUnderlyingListResult.ToList();
            return Json(UnderlyingListData, JsonRequestBehavior.AllowGet);
        }

        public JsonResult FetchAutocallSimulation(string ProductID, string SimulationCount)
        {
            ObjectResult<AutocallSimulationResult> objAutocallSimulationResult;
            List<AutocallSimulationResult> AutocallSimulationResultList = new List<AutocallSimulationResult>();

            objAutocallSimulationResult = objSP_PRICINGEntities.SP_FETCH_AUTOCALL_SIMULATIONS(ProductID, Convert.ToInt32(SimulationCount));

            var AutocallSimulationResultData = objAutocallSimulationResult.ToList();
            return Json(AutocallSimulationResultData, JsonRequestBehavior.AllowGet);
        }

        //public JsonResult GenerateAutocallSimulation(string ProductID, string ObservationStart, string ObservationEnd, string AutocallLevel, string EarlyRedemptionPaymentGap, string CouponIfHit, string InterestRateHit, string ObservationFrequency, string IV, string RFR, string NoOfSimulation)
        //{
        //    ThreadStart testThread1Start = new ThreadStart(() => SimulationThread1("", "", "", "", "", "", "", "", "", "", "", ""));
        //    ThreadStart testThread2Start = new ThreadStart(() => SimulationThread1("", "", "", "", "", "", "", "", "", "", "", ""));

        //    Thread[] testThread = new Thread[2];
        //    testThread[0] = new Thread(testThread1Start);
        //    testThread[1] = new Thread(testThread2Start);

        //    foreach (Thread myThread in testThread)
        //    {
        //        myThread.Start();
        //    }

        //    var Result = objSP_PRICINGEntities.SP_GENERATE_AUTOCALL_SIMULATIONS(ProductID, Convert.ToInt32(ObservationStart), Convert.ToInt32(ObservationEnd),
        //            Convert.ToDouble(AutocallLevel), Convert.ToInt32(EarlyRedemptionPaymentGap), Convert.ToDouble(CouponIfHit), Convert.ToDouble(InterestRateHit),
        //            Convert.ToInt32(ObservationFrequency), Convert.ToDouble(IV), Convert.ToDouble(RFR), Convert.ToInt32(NoOfSimulation));

        //    ObjectResult<AutocallDetailsResult> objAutocallDetailsResult;
        //    List<AutocallDetailsResult> AutocallDetailsResultList = new List<AutocallDetailsResult>();

        //    objAutocallDetailsResult = objSP_PRICINGEntities.SP_CALCULATE_AUTOCALL_DETAILS(ProductID);
        //    AutocallDetailsResultList = objAutocallDetailsResult.ToList();

        //    //var AutocallDetailsResultData = AutocallDetailsResultList;
        //    return Json(AutocallDetailsResultList, JsonRequestBehavior.AllowGet);
        //}

        //public void SimulationThread1(string ProductID, string ObservationStart, string ObservationEnd, string AutocallLevel, string EarlyRedemptionPaymentGap, string CouponIfHit, string InterestRateHit, string ObservationFrequency, string IV, string RFR, string NoOfSimulation, string StartCount)
        //{
        //    var Result = objSP_PRICINGEntities.SP_GENERATE_AUTOCALL_SIMULATIONS(ProductID, Convert.ToInt32(ObservationStart), Convert.ToInt32(ObservationEnd),
        //            Convert.ToDouble(AutocallLevel), Convert.ToInt32(EarlyRedemptionPaymentGap), Convert.ToDouble(CouponIfHit), Convert.ToDouble(InterestRateHit),
        //            Convert.ToInt32(ObservationFrequency), Convert.ToDouble(IV), Convert.ToDouble(RFR), Convert.ToInt32(NoOfSimulation));
        //}

        //public void SimulationThread2(string ProductID, string ObservationStart, string ObservationEnd, string AutocallLevel, string EarlyRedemptionPaymentGap, string CouponIfHit, string InterestRateHit, string ObservationFrequency, string IV, string RFR, string NoOfSimulation, string StartCount)
        //{
        //    var Result = objSP_PRICINGEntities.SP_GENERATE_AUTOCALL_SIMULATIONS(ProductID, Convert.ToInt32(ObservationStart), Convert.ToInt32(ObservationEnd),
        //            Convert.ToDouble(AutocallLevel), Convert.ToInt32(EarlyRedemptionPaymentGap), Convert.ToDouble(CouponIfHit), Convert.ToDouble(InterestRateHit),
        //            Convert.ToInt32(ObservationFrequency), Convert.ToDouble(IV), Convert.ToDouble(RFR), Convert.ToInt32(NoOfSimulation));
        //}

        public JsonResult FetchImpliedVolatilityList()
        {
            ObjectResult<LookupResult> objLookupResult;
            //List<LookupMaster> ImpliedVilatilityList = new List<LookupMaster>();

            objLookupResult = objSP_PRICINGEntities.SP_FETCH_LOOKUP_VALUES("AIV", true);

            var ImpliedVolatilityData = objLookupResult.ToList();
            return Json(ImpliedVolatilityData, JsonRequestBehavior.AllowGet);
        }

        public JsonResult FetchPricingDeploymentRate(string Tenure, string Entity, string IsSecured)
        {
            string strDeploymentRate = "";

            var DeploymentRate = objSP_PRICINGEntities.SP_FETCH_PRICING_DEPLOYMENT_RATE(Convert.ToInt32(Tenure), Convert.ToInt32(Entity), Convert.ToInt32(IsSecured));
            strDeploymentRate = Convert.ToString(DeploymentRate.SingleOrDefault());

            return Json(strDeploymentRate);
        }

        public JsonResult FetchObservationFrequencyList()
        {
            ObjectResult<LookupResult> objLookupResult;
            //List<LookupMaster> ImpliedVilatilityList = new List<LookupMaster>();

            objLookupResult = objSP_PRICINGEntities.SP_FETCH_LOOKUP_VALUES("OF", true);

            var ObservationFrequencyData = objLookupResult.ToList();
            return Json(ObservationFrequencyData, JsonRequestBehavior.AllowGet);
        }

        public JsonResult FetchPricerStatus(string PricerType, string ProductID)
        {
            string strStatus = "";

            if (PricerType == "A")
            {
                ObjectResult<AutoCallEditResult> objAutoCallEditResult = objSP_PRICINGEntities.FETCH_AUTOCALL_EDIT_DETAILS(ProductID);
                List<AutoCallEditResult> AutoCallEditResultList = objAutoCallEditResult.ToList();

                Autocall oAutocall = new Autocall();
                General.ReflectSingleData(oAutocall, AutoCallEditResultList[0]);

                strStatus = oAutocall.Status;
            }
            return Json(strStatus);
        }

        [HttpGet]
        public ActionResult AutocallList()
        {
            LoginController objLoginController = new LoginController();
            try
            {
                if (ValidateSession())
                {
                    Autocall objAutoCall = new Autocall();

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

                    objAutoCall.StatusList = StatusList;

                    return View(objAutoCall);
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
                LogError(ex.Message, ex.StackTrace, "MCPricersController", "AutocallList", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        public ActionResult FetchAutoCallList(string ProductID, string ProductTenure, string DeploymentRate, string Distributor, string SalesComment, string TradingComment, string UnderlyingID, string OptionTenure, string Status, string CouponIfHit, string AutoCallLevel, string ObservationFrequency, string ObservationStart, string ObservationEnd, string FromDate, string ToDate)
        {
            try
            {
                if (ValidateSession())
                {
                    List<Autocall> AutocallList = new List<Autocall>();

                    if (ProductID == "" || ProductID == "--Select--")
                        ProductID = "ALL";

                    if (ProductTenure == "" || ProductTenure == "0" || ProductTenure == "--Select--")
                        ProductTenure = "ALL";

                    if (DeploymentRate == "0")
                        DeploymentRate = "ALL";

                    if (Distributor == "" || Distributor == "--Select--")
                        Distributor = "ALL";

                    if (SalesComment == "" || SalesComment == "--Select--")
                        SalesComment = "ALL";

                    if (TradingComment == "" || TradingComment == "--Select--")
                        TradingComment = "ALL";

                    if (CouponIfHit == "" || CouponIfHit == "0" || CouponIfHit == "--Select--")
                        CouponIfHit = "ALL";

                    if (AutoCallLevel == "" || AutoCallLevel == "0" || AutoCallLevel == "--Select--")
                        AutoCallLevel = "ALL";

                    if (ObservationFrequency == "" || ObservationFrequency == "0" || ObservationFrequency == "--Select--")
                        ObservationFrequency = "ALL";

                    if (UnderlyingID == "" || UnderlyingID == "0" || UnderlyingID == "--Select--")
                        UnderlyingID = "ALL";

                    if (OptionTenure == "" || OptionTenure == "0" || OptionTenure == "--Select--")
                        OptionTenure = "ALL";

                    if (Status == "" || Status == "0" || Status == "--Select--")
                        Status = "ALL";

                    if (FromDate == "")
                        FromDate = "1900-01-01";

                    if (ToDate == "")
                        ToDate = "2900-01-01";


                    if (ObservationStart == "" || ObservationStart == "0" || ObservationStart == "--Select--")
                        ObservationStart = "ALL";

                    if (ObservationEnd == "" || ObservationEnd == "0" || ObservationEnd == "--Select--")
                        ObservationEnd = "ALL";

                    DataSet dsResult = new DataSet();
                    dsResult = General.ExecuteDataSet("FETCH_AUTO_CALL", ProductID, ProductTenure, DeploymentRate, Distributor, SalesComment, TradingComment, UnderlyingID, OptionTenure, Status, CouponIfHit, AutoCallLevel, ObservationFrequency, ObservationStart, ObservationEnd, FromDate, ToDate);

                    if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                    {
                        foreach (DataRow dr in dsResult.Tables[0].Rows)
                        {
                            Autocall obj = new Autocall();

                            obj.ProductID = Convert.ToString(dr["ProductID"]);
                            obj.UnderlyingName = Convert.ToString(dr["UnderlyingID"]);
                            obj.ProductTenure = Convert.ToString(dr["ProductTenure"]);
                            obj.OptionTenure = Convert.ToInt32(dr["OptionTenureMonth"]);
                            obj.DeploymentRate = Convert.ToDouble(dr["DeploymentRate"]);
                            obj.EdelweissBuiltIn = Convert.ToDouble(dr["EdelweissBuiltIn"]);
                            obj.DistributorBuiltIn = Convert.ToDouble(dr["DistributorBuiltIn"]);
                            obj.CouponScenario1 = Convert.ToString(dr["CouponScenario"]);
                            obj.ConfirmedOn = Convert.ToString(dr["ConfirmedOn"]);
                            obj.Distributor = Convert.ToString(dr["Distributor"]);
                            obj.AutocallLevel = Convert.ToInt32(dr["AutocallLevel"]);
                            obj.CouponIfHit = Convert.ToInt32(dr["CouponIfHit"]);
                            obj.ObservationStart = Convert.ToInt32(dr["ObservationStartDate"]);
                            obj.ObservationEnd = Convert.ToInt32(dr["ObservationEndDate"]);
                            obj.ObservationFrequencyID = Convert.ToInt32(dr["ObservationFrequencyID"]);
                            obj.EarlyRedemptionPaymentGap = Convert.ToInt32(dr["EarlyRedemptionPaymentGap"]);
                            obj.NonPPLevel = Convert.ToInt32(dr["NonPPLevel"]);
                            obj.Status = Convert.ToString(dr["Status"]);
                            obj.SalesComments = Convert.ToString(dr["SalesComment"]);
                            obj.TradingComments = Convert.ToString(dr["TradingComment"]);

                            AutocallList.Add(obj);
                        }
                    }

                    var AutocallListData = AutocallList.ToList();
                    return Json(AutocallListData, JsonRequestBehavior.AllowGet);
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
                LogError(ex.Message, ex.StackTrace, "MCPricersController", "FetchAutoCallList", objUserMaster.UserID);
                return Json("");
            }
        }

        public ActionResult AutoCompleteProductID(string term)
        {
            try
            {
                List<Autocall> AutocallList = new List<Autocall>();

                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("FETCH_AUTO_CALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "1900-01-01", "2900-01-01");

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in dsResult.Tables[0].Rows)
                    {
                        Autocall obj = new Autocall();

                        obj.ProductID = Convert.ToString(dr["ProductID"]);
                        obj.UnderlyingName = Convert.ToString(dr["UnderlyingID"]);
                        obj.ProductTenure = Convert.ToString(dr["ProductTenure"]);
                        obj.OptionTenure = Convert.ToInt32(dr["OptionTenureMonth"]);
                        obj.DeploymentRate = Convert.ToDouble(dr["DeploymentRate"]);
                        obj.EdelweissBuiltIn = Convert.ToDouble(dr["EdelweissBuiltIn"]);
                        obj.DistributorBuiltIn = Convert.ToDouble(dr["DistributorBuiltIn"]);
                        obj.CouponScenario1 = Convert.ToString(dr["CouponScenario"]);
                        obj.ConfirmedOn = Convert.ToString(dr["ConfirmedOn"]);

                        obj.AutocallLevel = Convert.ToInt32(dr["AutocallLevel"]);
                        obj.CouponIfHit = Convert.ToInt32(dr["CouponIfHit"]);
                        obj.ObservationStart = Convert.ToInt32(dr["ObservationStartDate"]);
                        obj.ObservationEnd = Convert.ToInt32(dr["ObservationEndDate"]);
                        obj.ObservationFrequencyID = Convert.ToInt32(dr["ObservationFrequencyID"]);
                        obj.EarlyRedemptionPaymentGap = Convert.ToInt32(dr["EarlyRedemptionPaymentGap"]);
                        obj.NonPPLevel = Convert.ToInt32(dr["NonPPLevel"]);


                        AutocallList.Add(obj);
                    }
                }

                var result = (from objRuleList in AutocallList
                              where objRuleList.ProductID.ToLower().StartsWith(term.ToLower())
                              select objRuleList);

                return Json(result);

            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "MCPricersController", "AutoCompleteProductID", objUserMaster.UserID);
                Session["ErrorData"] = ex.Message;
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        public ActionResult AutoCompleteDistributor(string term)
        {
            try
            {

                List<Autocall> AutocallList = new List<Autocall>();

                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("FETCH_AUTO_CALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "1900-01-01", "2900-01-01");

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in dsResult.Tables[0].Rows)
                    {
                        Autocall obj = new Autocall();

                        obj.ProductID = Convert.ToString(dr["ProductID"]);
                        obj.UnderlyingName = Convert.ToString(dr["UnderlyingID"]);
                        obj.ProductTenure = Convert.ToString(dr["ProductTenure"]);
                        obj.OptionTenure = Convert.ToInt32(dr["OptionTenureMonth"]);
                        obj.DeploymentRate = Convert.ToDouble(dr["DeploymentRate"]);
                        obj.EdelweissBuiltIn = Convert.ToDouble(dr["EdelweissBuiltIn"]);
                        obj.DistributorBuiltIn = Convert.ToDouble(dr["DistributorBuiltIn"]);
                        obj.CouponScenario1 = Convert.ToString(dr["CouponScenario"]);
                        obj.ConfirmedOn = Convert.ToString(dr["ConfirmedOn"]);

                        obj.AutocallLevel = Convert.ToInt32(dr["AutocallLevel"]);
                        obj.CouponIfHit = Convert.ToInt32(dr["CouponIfHit"]);
                        obj.ObservationStart = Convert.ToInt32(dr["ObservationStartDate"]);
                        obj.ObservationEnd = Convert.ToInt32(dr["ObservationEndDate"]);
                        obj.ObservationFrequencyID = Convert.ToInt32(dr["ObservationFrequencyID"]);
                        obj.EarlyRedemptionPaymentGap = Convert.ToInt32(dr["EarlyRedemptionPaymentGap"]);
                        obj.NonPPLevel = Convert.ToInt32(dr["NonPPLevel"]);


                        AutocallList.Add(obj);
                    }
                }

                var DistinctItems = AutocallList.GroupBy(x => x.Distributor).Select(y => y.First());

                var result = (from objRuleList in DistinctItems
                              where objRuleList.Distributor.ToLower().StartsWith(term.ToLower())
                              select objRuleList);

                return Json(result);

            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "MCPricersController", "AutoCompleteDistributor", objUserMaster.UserID);
                Session["ErrorData"] = ex.Message;
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        public ActionResult AutoCompleteUnderlyingID(string term)
        {
            try
            {
                List<Autocall> AutocallList = new List<Autocall>();

                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("FETCH_AUTO_CALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "1900-01-01", "2900-01-01");

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in dsResult.Tables[0].Rows)
                    {
                        Autocall obj = new Autocall();

                        obj.ProductID = Convert.ToString(dr["ProductID"]);
                        obj.UnderlyingName = Convert.ToString(dr["UnderlyingID"]);
                        obj.ProductTenure = Convert.ToString(dr["ProductTenure"]);
                        obj.OptionTenure = Convert.ToInt32(dr["OptionTenureMonth"]);
                        obj.DeploymentRate = Convert.ToDouble(dr["DeploymentRate"]);
                        obj.EdelweissBuiltIn = Convert.ToDouble(dr["EdelweissBuiltIn"]);
                        obj.DistributorBuiltIn = Convert.ToDouble(dr["DistributorBuiltIn"]);
                        obj.CouponScenario1 = Convert.ToString(dr["CouponScenario"]);
                        obj.ConfirmedOn = Convert.ToString(dr["ConfirmedOn"]);

                        obj.AutocallLevel = Convert.ToInt32(dr["AutocallLevel"]);
                        obj.CouponIfHit = Convert.ToInt32(dr["CouponIfHit"]);
                        obj.ObservationStart = Convert.ToInt32(dr["ObservationStartDate"]);
                        obj.ObservationEnd = Convert.ToInt32(dr["ObservationEndDate"]);
                        obj.ObservationFrequencyID = Convert.ToInt32(dr["ObservationFrequencyID"]);
                        obj.EarlyRedemptionPaymentGap = Convert.ToInt32(dr["EarlyRedemptionPaymentGap"]);
                        obj.NonPPLevel = Convert.ToInt32(dr["NonPPLevel"]);


                        AutocallList.Add(obj);
                    }
                }

                var DistinctItems = AutocallList.GroupBy(x => x.UnderlyingName).Select(y => y.First());

                var result = (from objRuleList in DistinctItems
                              where objRuleList.UnderlyingName.ToLower().StartsWith(term.ToLower())
                              select objRuleList);

                return Json(result);

            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "MCPricersController", "AutoCompleteUnderlyingID", objUserMaster.UserID);
                Session["ErrorData"] = ex.Message;
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        public JsonResult ManagePricerStatusLog(string PricerType, string ProductID, string StatusCode)
        {
            try
            {
                Int32 intResult = 0;
                bool PPorNonPP = false;

                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];

                var Result = objSP_PRICINGEntities.SP_MANAGE_PRICER_STATUS_LOG(PricerType, ProductID, objUserMaster.UserID, StatusCode);
                intResult = Convert.ToInt32(Result.SingleOrDefault());

                if (StatusCode == "AP")
                {
                    var productID = ProductID;


                    EncryptDecrypt obj = new EncryptDecrypt();
                    var encryptedpaswd = obj.Encrypt(objUserMaster.Password, "SPPricing", CryptographyEngine.AlgorithmType.DES);

                    var isPrincipalProtected = objSP_PRICINGEntities.SP_FETCH_IS_PROTECTED(PricerType, ProductID);
                    PPorNonPP = Convert.ToBoolean(isPrincipalProtected.SingleOrDefault());

                    var ProductType = "PP";
                    if (PPorNonPP == false)
                    {
                        ProductType = "NonPP";
                    }

                    var Url = "http://edemumnewuatvm4:63400/Login.aspx?UserId=" + objUserMaster.LoginName + "&Key=" + encryptedpaswd + "&ProductId=" + productID + "&ProductType=" + ProductType;
                    return Json(Url);
                }


                return Json(intResult);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "MCPricersController", "ManagePricerStatusLog", objUserMaster.UserID);
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
                LogError(ex.Message, ex.StackTrace, "MCPricersController", "SendDistributorMail", objUserMaster.UserID);
                return Json("");
            }
        }

        //public JsonResult GenerateAutocallSimulation(string ProductID, string ObservationStart, string ObservationEnd, string AutocallLevel, string EarlyRedemptionPaymentGap, string CouponIfHit, string InterestRateHit, string ObservationFrequency, string IV, string RFR, string NoOfSimulation, string RedemptionPeriodDays, string NonPPLevel)
        //{
        //    Thread[] testThread;

        //    var strProductID = objSP_PRICINGEntities.CLEAR_EXISTING_AUTOCALL_SIMULATIONS(ProductID);
        //    Int32 intNoOfSimulations = Convert.ToInt32(NoOfSimulation);

        //    if (intNoOfSimulations <= 500)
        //    {
        //        testThread = new Thread[1];
        //        ThreadStart testThread1Start = new ThreadStart(() => SimulationThread1(ProductID, ObservationStart, ObservationEnd, AutocallLevel, EarlyRedemptionPaymentGap, CouponIfHit, InterestRateHit, ObservationFrequency, IV, RFR, intNoOfSimulations, 1, RedemptionPeriodDays, NonPPLevel));
        //        testThread[0] = new Thread(testThread1Start);

        //        testThread[0].Start();
        //    }
        //    else
        //    {
        //        Int32 intThreadCount = intNoOfSimulations / 5;
        //        Int32 intRemainingCount = intNoOfSimulations % 5;

        //        ThreadStart testThread1Start;
        //        ThreadStart testThread2Start;
        //        ThreadStart testThread3Start;
        //        ThreadStart testThread4Start;
        //        ThreadStart testThread5Start;
        //        ThreadStart testThread6Start;

        //        if (intRemainingCount > 0)
        //            testThread = new Thread[6];
        //        else
        //            testThread = new Thread[5];

        //        testThread1Start = new ThreadStart(() => SimulationThread1(ProductID, ObservationStart, ObservationEnd, AutocallLevel, EarlyRedemptionPaymentGap, CouponIfHit, InterestRateHit, ObservationFrequency, IV, RFR, intThreadCount, 1, RedemptionPeriodDays, NonPPLevel));
        //        testThread[0] = new Thread(testThread1Start);

        //        testThread2Start = new ThreadStart(() => SimulationThread1(ProductID, ObservationStart, ObservationEnd, AutocallLevel, EarlyRedemptionPaymentGap, CouponIfHit, InterestRateHit, ObservationFrequency, IV, RFR, (intThreadCount * 2), (intThreadCount + 1), RedemptionPeriodDays, NonPPLevel));
        //        testThread[1] = new Thread(testThread2Start);

        //        testThread3Start = new ThreadStart(() => SimulationThread1(ProductID, ObservationStart, ObservationEnd, AutocallLevel, EarlyRedemptionPaymentGap, CouponIfHit, InterestRateHit, ObservationFrequency, IV, RFR, (intThreadCount * 3), (intThreadCount * 2) + 1, RedemptionPeriodDays, NonPPLevel));
        //        testThread[2] = new Thread(testThread3Start);

        //        testThread4Start = new ThreadStart(() => SimulationThread1(ProductID, ObservationStart, ObservationEnd, AutocallLevel, EarlyRedemptionPaymentGap, CouponIfHit, InterestRateHit, ObservationFrequency, IV, RFR, (intThreadCount * 4), (intThreadCount * 3) + 1, RedemptionPeriodDays, NonPPLevel));
        //        testThread[3] = new Thread(testThread4Start);

        //        testThread5Start = new ThreadStart(() => SimulationThread1(ProductID, ObservationStart, ObservationEnd, AutocallLevel, EarlyRedemptionPaymentGap, CouponIfHit, InterestRateHit, ObservationFrequency, IV, RFR, (intThreadCount * 5), (intThreadCount * 4) + 1, RedemptionPeriodDays, NonPPLevel));
        //        testThread[4] = new Thread(testThread5Start);

        //        if (intRemainingCount > 0)
        //        {
        //            testThread6Start = new ThreadStart(() => SimulationThread1(ProductID, ObservationStart, ObservationEnd, AutocallLevel, EarlyRedemptionPaymentGap, CouponIfHit, InterestRateHit, ObservationFrequency, IV, RFR, (intThreadCount * 5) + 1 + intRemainingCount, (intThreadCount * 5) + 1, RedemptionPeriodDays, NonPPLevel));
        //            testThread[5] = new Thread(testThread6Start);
        //        }

        //        foreach (Thread myThread in testThread)
        //        {
        //            myThread.Start();
        //        }
        //    }

        //    Int32 intMaxThreadCount = testThread.Length;

        //    for (Int32 i = 0; i < intMaxThreadCount; i++)
        //    {
        //        while (testThread[i].IsAlive)
        //        {
        //            System.Threading.Thread.Sleep(5);
        //        }
        //    }

        //    ObjectResult<AutocallDetailsResult> objAutocallDetailsResult;
        //    List<AutocallDetailsResult> AutocallDetailsResultList = new List<AutocallDetailsResult>();

        //    objAutocallDetailsResult = objSP_PRICINGEntities.SP_CALCULATE_AUTOCALL_DETAILS(ProductID);
        //    AutocallDetailsResultList = objAutocallDetailsResult.ToList();

        //    return Json(AutocallDetailsResultList, JsonRequestBehavior.AllowGet);
        //}


        //public void SimulationThread1(string ProductID, string ObservationStart, string ObservationEnd, string AutocallLevel, string EarlyRedemptionPaymentGap, string CouponIfHit, string InterestRateHit, string ObservationFrequency, string IV, string RFR, Int32 NoOfSimulation, Int32 StartCount, string RedemptionPeriodDays, string NonPPLevel)
        //{
        //    ObjectResult<AutocallSimulationDataResult> objAutocallSimulationDataResult;
        //    List<AutocallSimulationDataResult> AutocallSimulationDataResultList = new List<AutocallSimulationDataResult>();

        //    objAutocallSimulationDataResult = objSP_PRICINGEntities.SP_GENERATE_AUTOCALL_SIMULATIONS(ProductID, Convert.ToInt32(ObservationStart), Convert.ToInt32(ObservationEnd),
        //            Convert.ToDouble(AutocallLevel), Convert.ToInt32(EarlyRedemptionPaymentGap), Convert.ToDouble(CouponIfHit), Convert.ToDouble(InterestRateHit),
        //            Convert.ToInt32(ObservationFrequency), Convert.ToDouble(IV), Convert.ToDouble(RFR), NoOfSimulation, StartCount, Convert.ToInt32(RedemptionPeriodDays),
        //            Convert.ToDouble(NonPPLevel));

        //    AutocallSimulationDataResultList = objAutocallSimulationDataResult.ToList();

        //    InsertAutocallSimulationData(AutocallSimulationDataResultList);
        //}


        public void InsertAutocallSimulationData(List<AutocallSimulationDataResult> AutocallSimulationDataResultList)
        {
            lock (lockThis)
            {
                #region Source and Destination Column
                string strSourceColumn = "AUTOCALL_ID|SIMULATION_ID|MONTH|MONTH_DAYS|REDEMPTION_MONTH|AUTOCALL_LEVEL|BUFFER|UNDERLYING_LEVEL|IS_AUTOCALLED|DEPLOYMENT_RATE|INTEREST_RATE_HIT|COUPON_IF_HIT|ROW_NUM";
                string[] arrSourceColumn = null;
                if (strSourceColumn != "")
                    arrSourceColumn = strSourceColumn.Split('|');

                DataTable dtData = new DataTable();

                for (int i = 0; i < arrSourceColumn.Length; i++)
                {
                    dtData.Columns.Add(arrSourceColumn[i]);
                }

                string strDestinationColumn = "AUTOCALL_ID|SIMULATION_ID|MONTH|MONTH_DAYS|REDEMPTION_MONTH|AUTOCALL_LEVEL|BUFFER|UNDERLYING_LEVEL|IS_AUTOCALLED|DEPLOYMENT_RATE|INTEREST_RATE_HIT|COUPON_IF_HIT|ROW_NUM";
                string[] arrDestinationColumn = null;

                if (strDestinationColumn != "")
                    arrDestinationColumn = strDestinationColumn.Split('|');
                #endregion

                DataTable dtResult = new DataTable();

                dtResult.Columns.Add("AUTOCALL_ID");
                dtResult.Columns.Add("SIMULATION_ID");
                dtResult.Columns.Add("MONTH");
                dtResult.Columns.Add("MONTH_DAYS");
                dtResult.Columns.Add("REDEMPTION_MONTH");
                dtResult.Columns.Add("AUTOCALL_LEVEL");
                dtResult.Columns.Add("BUFFER");
                dtResult.Columns.Add("UNDERLYING_LEVEL");
                dtResult.Columns.Add("IS_AUTOCALLED");
                dtResult.Columns.Add("DEPLOYMENT_RATE");
                dtResult.Columns.Add("INTEREST_RATE_HIT");
                dtResult.Columns.Add("COUPON_IF_HIT");
                dtResult.Columns.Add("ROW_NUM");

                foreach (AutocallSimulationDataResult objAutocallSimulationDataResult in AutocallSimulationDataResultList)
                {
                    DataRow dr = dtResult.NewRow();

                    //General.ReflectSingleData(dr, objAutocallSimulationDataResult);
                    dr["AUTOCALL_ID"] = objAutocallSimulationDataResult.AUTOCALL_ID;
                    dr["SIMULATION_ID"] = objAutocallSimulationDataResult.SIMULATION_ID;
                    dr["MONTH"] = objAutocallSimulationDataResult.MONTH;
                    dr["MONTH_DAYS"] = objAutocallSimulationDataResult.MONTH_DAYS;
                    dr["REDEMPTION_MONTH"] = objAutocallSimulationDataResult.REDEMPTION_MONTH;
                    dr["AUTOCALL_LEVEL"] = objAutocallSimulationDataResult.AUTOCALL_LEVEL;
                    dr["BUFFER"] = objAutocallSimulationDataResult.BUFFER;
                    dr["UNDERLYING_LEVEL"] = objAutocallSimulationDataResult.UNDERLYING_LEVEL;
                    dr["IS_AUTOCALLED"] = objAutocallSimulationDataResult.IS_AUTOCALLED;
                    dr["DEPLOYMENT_RATE"] = objAutocallSimulationDataResult.DEPLOYMENT_RATE;
                    dr["INTEREST_RATE_HIT"] = objAutocallSimulationDataResult.INTEREST_RATE_HIT;
                    dr["COUPON_IF_HIT"] = objAutocallSimulationDataResult.COUPON_IF_HIT;
                    dr["ROW_NUM"] = objAutocallSimulationDataResult.ROW_NUM;

                    dtResult.Rows.Add(dr);
                }

                string strMyConnection = Convert.ToString(System.Configuration.ConfigurationManager.ConnectionStrings["SP_PRICINGConnectionString"]);

                if (arrSourceColumn != null && arrDestinationColumn != null && arrSourceColumn.Length == arrDestinationColumn.Length)
                {
                    using (SqlBulkCopy bulkCopy = new SqlBulkCopy(strMyConnection))
                    {
                        bulkCopy.DestinationTableName = "TBL_AUTOCALL_SIMULATION_DETAILS";

                        for (int i = 0; i < arrSourceColumn.Length; i++)
                        {
                            bulkCopy.ColumnMappings.Add(arrSourceColumn[i], arrDestinationColumn[i]);
                        }
                        bulkCopy.WriteToServer(dtResult);
                    }
                }
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
                objLoginController.LogError(ex.Message, ex.StackTrace, "MCPricersController", "ValidateSession", -1);
                return false;
            }
        }

        public void LogError(string strErrorDescription, string strStackTrace, string strClassName, string strMethodName, Int32 intUserId)
        {
            SP_PRICINGEntities objSP_PRICINGEntities = new SP_PRICINGEntities();
            var Count = objSP_PRICINGEntities.SP_ERROR_LOG(strErrorDescription, strStackTrace, strClassName, strMethodName, intUserId);
        }
    }
}
