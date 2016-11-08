using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using SPPricing.Models;
using System.Data;
using System.Data.Objects;
using System.Xml;
using Newtonsoft.Json;
using System.Web.UI.WebControls;
using System.IO;
using System.Web.UI;

namespace SPPricing.Controllers
{
    public class PeriodicValuationController : Controller
    {
        //
        // GET: /PeriodicValuation/
        SP_PRICINGEntities objSP_PRICINGEntities = new SP_PRICINGEntities();

        public ActionResult Index()
        {
            return View();
        }


        [HttpGet]
        public ActionResult PeriodicValuation(string ProductID, string WeightedAveragePrice)
        {
            try
            {
                if (ValidateSession())
                {
                    PeriodicValuation obj = new PeriodicValuation();
                    obj.ValuationDate = DateTime.Now.Date;

                    return View(obj);
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

                LogError(ex.Message, ex.StackTrace, "PeriodicValuationController", "PeriodicValuation Get", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        [HttpPost, ValidateInput(false)]
        public ActionResult PeriodicValuation(string Command, PeriodicValuation objPeriodicValuation, FormCollection objFormCollection)
        {
            try
            {
                if (ValidateSession())
                {
                    if (Command == "Save")
                    {

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

                LogError(ex.Message, ex.StackTrace, "PeriodicValuationController", "PeriodicValuation Post", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        [HttpGet]
        public ActionResult WeeklyValuation()
        {
            return View();
        }

        public JsonResult FetchWeeklyValuationList(string ValuationDate, string TenureMultiplier, string ValuationMultiplier, string ValuationStrikeThreshold, string ValuationDivisor)
        {
            UserMaster objUserMaster = new UserMaster();
            objUserMaster = (UserMaster)Session["LoggedInUser"];

            try
            {
                //{
                //    if (ValuationDate == "")
                //        ValuationDate = "2015-12-31";
                //objSPNoteMTM.ReportDate = Convert.ToDateTime(strReportDate.Substring(6, 4) + "-" + strReportDate.Substring(0, 2) + "-" + strReportDate.Substring(3, 2));

                if (ValuationDate == "")
                    ValuationDate = "1900-01-01";

                ObjectResult<WealthValuationResult> objWealthValuationResult = objSP_PRICINGEntities.FETCH_SP_WEALTH_VALUATION(Convert.ToDateTime(ValuationDate), objUserMaster.UserID, Convert.ToDouble(TenureMultiplier), Convert.ToDouble(ValuationMultiplier), Convert.ToDouble(ValuationStrikeThreshold), Convert.ToDouble(ValuationDivisor));
                List<WealthValuationResult> WealthValuationResultList = objWealthValuationResult.ToList();

                return Json(WealthValuationResultList, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                LogError(ex.Message, ex.StackTrace, "PeriodicValuationController", "FetchWeeklyValuationList", objUserMaster.UserID);
                return Json("");
            }
        }

        public void ExportWeeklyToExcel(string ValuationDate, string TenureMultiplier, string ValuationMultiplier, string ValuationStrikeThreshold, string ValuationDivisor)
        {

            UserMaster objUserMaster = new UserMaster();
            objUserMaster = (UserMaster)Session["LoggedInUser"];

            DataSet dsReportAutoCall = General.ExecuteDataSet("FETCH_SP_WEALTH_VALUATION", Convert.ToDateTime(ValuationDate), objUserMaster.UserID, Convert.ToDouble(TenureMultiplier), Convert.ToDouble(ValuationMultiplier), Convert.ToDouble(ValuationStrikeThreshold), Convert.ToDouble(ValuationDivisor));

            dsReportAutoCall.Tables[0].Columns.Remove("RowNum");
            dsReportAutoCall.Tables[0].Columns.Remove("LastTradePrice");
            dsReportAutoCall.Tables[0].Columns.Remove("LastTradeValue");
            dsReportAutoCall.Tables[0].Columns.Remove("TotalTradeValue");
            dsReportAutoCall.Tables[0].Columns.Remove("LastTradeYield");
            dsReportAutoCall.Tables[0].Columns.Remove("WeightedAveragePrice");
            dsReportAutoCall.Tables[0].Columns.Remove("UpdatedWeightedAveragePrice");
            dsReportAutoCall.Tables[0].Columns.Remove("OptionValue");
            dsReportAutoCall.Tables[0].Columns.Remove("BondDiscount");
            dsReportAutoCall.Tables[0].Columns.Remove("DiscountedOption");
            dsReportAutoCall.Tables[0].Columns.Remove("IsHaircut");
            dsReportAutoCall.Tables[0].Columns.Remove("HaircutRate");
            dsReportAutoCall.Tables[0].Columns.Remove("HaircutCalculation");
            dsReportAutoCall.Tables[0].Columns.Remove("IsIRR");
            dsReportAutoCall.Tables[0].Columns.Remove("IRRValue");
            dsReportAutoCall.Tables[0].Columns.Remove("ValueAfterIRR");
            dsReportAutoCall.Tables[0].Columns.Remove("IsAmortization");
            dsReportAutoCall.Tables[0].Columns.Remove("AmortizationValue");
            dsReportAutoCall.Tables[0].Columns.Remove("IsCustomValue");
            dsReportAutoCall.Tables[0].Columns.Remove("CustomValue");
            dsReportAutoCall.Tables[0].Columns.Remove("CreatedBy");


            GridView gv = new GridView();
            gv.DataSource = dsReportAutoCall;
            gv.DataBind();

            Response.ClearContent();
            Response.Buffer = true;
            Response.AddHeader("content-disposition", "attachment; filename=" + "EXPORT_WEEKLY" + ".xls");
            Response.ContentType = "application/ms-excel";
            Response.Charset = "";
            StringWriter sw = new StringWriter();
            HtmlTextWriter htw = new HtmlTextWriter(sw);
            gv.RenderControl(htw);
            Response.Output.Write(sw.ToString());
            Response.Flush();
            Response.End();
        }

        [HttpGet]
        public ActionResult FortnightlyValuation()
        {
            return View();
        }

        public JsonResult FetchFortnightValuationList(string ValuationDate, string TenureMultiplier, string ValuationMultiplier, string ValuationStrikeThreshold, string ValuationDivisor)
        {
            UserMaster objUserMaster = new UserMaster();
            objUserMaster = (UserMaster)Session["LoggedInUser"];

            try
            {
                if (ValuationDate == "")
                    ValuationDate = "1900-01-01";

                ObjectResult<ForthnightValuationResult> objForthnightValuationResult = objSP_PRICINGEntities.FETCH_SP_FORTNIGHT_VALUATION(Convert.ToDateTime(ValuationDate), objUserMaster.UserID, Convert.ToDouble(TenureMultiplier), Convert.ToDouble(ValuationMultiplier), Convert.ToDouble(ValuationStrikeThreshold), Convert.ToDouble(ValuationDivisor));
                List<ForthnightValuationResult> ForthnightValuationResultList = objForthnightValuationResult.ToList();

                return Json(ForthnightValuationResultList, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                LogError(ex.Message, ex.StackTrace, "PeriodicValuationController", "FetchFortnightValuationList", objUserMaster.UserID);
                return Json("");
            }
        }

        public void ExportFortnightToExcel(string ValuationDate, string TenureMultiplier, string ValuationMultiplier, string ValuationStrikeThreshold, string ValuationDivisor)
        {
            UserMaster objUserMaster = new UserMaster();
            objUserMaster = (UserMaster)Session["LoggedInUser"];

            DataSet dsReportAutoCall = General.ExecuteDataSet("FETCH_SP_FORTNIGHT_VALUATION", Convert.ToDateTime(ValuationDate), objUserMaster.UserID, Convert.ToDouble(TenureMultiplier), Convert.ToDouble(ValuationMultiplier), Convert.ToDouble(ValuationStrikeThreshold), Convert.ToDouble(ValuationDivisor));

            dsReportAutoCall.Tables[0].Columns.Remove("RowNum");
            dsReportAutoCall.Tables[0].Columns.Remove("LastTradePrice");
            dsReportAutoCall.Tables[0].Columns.Remove("LastTradeValue");
            dsReportAutoCall.Tables[0].Columns.Remove("TotalTradeValue");
            dsReportAutoCall.Tables[0].Columns.Remove("LastTradeYield");
            dsReportAutoCall.Tables[0].Columns.Remove("SPMTMSurfaceValuation");
            dsReportAutoCall.Tables[0].Columns.Remove("UpdatedSPMTMSurfaceValuation");
            dsReportAutoCall.Tables[0].Columns.Remove("OptionValue");
            dsReportAutoCall.Tables[0].Columns.Remove("BondDiscount");
            dsReportAutoCall.Tables[0].Columns.Remove("DiscountedOption");
            dsReportAutoCall.Tables[0].Columns.Remove("IsHaircut");
            dsReportAutoCall.Tables[0].Columns.Remove("HaircutRate");
            dsReportAutoCall.Tables[0].Columns.Remove("HaircutCalculation");
            dsReportAutoCall.Tables[0].Columns.Remove("IsIRR");
            dsReportAutoCall.Tables[0].Columns.Remove("IRRValue");
            dsReportAutoCall.Tables[0].Columns.Remove("ValueAfterIRR");
            dsReportAutoCall.Tables[0].Columns.Remove("IsAmortization");
            dsReportAutoCall.Tables[0].Columns.Remove("AmortizationValue");
            dsReportAutoCall.Tables[0].Columns.Remove("IsCustomValue");
            dsReportAutoCall.Tables[0].Columns.Remove("CustomValue");
            dsReportAutoCall.Tables[0].Columns.Remove("CreatedBy");

            GridView gv = new GridView();
            gv.DataSource = dsReportAutoCall;
            gv.DataBind();

            Response.ClearContent();
            Response.Buffer = true;
            Response.AddHeader("content-disposition", "attachment; filename=" + "EXPORT_FORTNIGHT" + ".xls");
            Response.ContentType = "application/ms-excel";
            Response.Charset = "";
            StringWriter sw = new StringWriter();
            HtmlTextWriter htw = new HtmlTextWriter(sw);
            gv.RenderControl(htw);
            Response.Output.Write(sw.ToString());
            Response.Flush();
            Response.End();
        }

        [HttpGet]
        public ActionResult QuarterlyValuation()
        {
            return View();
        }

        public JsonResult FetchQuaterlyValuationList(string ValuationDate, string TenureMultiplier, string ValuationMultiplier, string ValuationStrikeThreshold, string ValuationDivisor)
        {
            UserMaster objUserMaster = new UserMaster();
            objUserMaster = (UserMaster)Session["LoggedInUser"];

            try
            {
                if (ValuationDate == "")
                    ValuationDate = "1900-01-01";

                ObjectResult<QuaterlyValuationResult> objQuaterlyValuationResult = objSP_PRICINGEntities.FETCH_SP_QUATERLY_VALUATION(Convert.ToDateTime(ValuationDate), objUserMaster.UserID, Convert.ToDouble(TenureMultiplier), Convert.ToDouble(ValuationMultiplier), Convert.ToDouble(ValuationStrikeThreshold), Convert.ToDouble(ValuationDivisor));
                List<QuaterlyValuationResult> QuaterlyValuationResultList = objQuaterlyValuationResult.ToList();

                return Json(QuaterlyValuationResultList, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                LogError(ex.Message, ex.StackTrace, "PeriodicValuationController", "FetchQuaterlyValuationList", objUserMaster.UserID);
                return Json("");
            }
        }

        public void ExportQuaterlyToExcel(string ValuationDate, string TenureMultiplier, string ValuationMultiplier, string ValuationStrikeThreshold, string ValuationDivisor)
        {
            UserMaster objUserMaster = new UserMaster();
            objUserMaster = (UserMaster)Session["LoggedInUser"];

            DataSet dsReportAutoCall = General.ExecuteDataSet("FETCH_SP_QUATERLY_VALUATION", Convert.ToDateTime(ValuationDate), objUserMaster.UserID, Convert.ToDouble(TenureMultiplier), Convert.ToDouble(ValuationMultiplier), Convert.ToDouble(ValuationStrikeThreshold), Convert.ToDouble(ValuationDivisor));

            dsReportAutoCall.Tables[0].Columns.Remove("RowNum");
            dsReportAutoCall.Tables[0].Columns.Remove("LastTradePrice");
            dsReportAutoCall.Tables[0].Columns.Remove("LastTradeValue");
            dsReportAutoCall.Tables[0].Columns.Remove("TotalTradeValue");
            dsReportAutoCall.Tables[0].Columns.Remove("LastTradeYield");
            dsReportAutoCall.Tables[0].Columns.Remove("NAV");
            dsReportAutoCall.Tables[0].Columns.Remove("UpdatedNAV");
            dsReportAutoCall.Tables[0].Columns.Remove("OptionValue");
            dsReportAutoCall.Tables[0].Columns.Remove("BondDiscount");
            dsReportAutoCall.Tables[0].Columns.Remove("DiscountedOption");
            dsReportAutoCall.Tables[0].Columns.Remove("IsHaircut");
            dsReportAutoCall.Tables[0].Columns.Remove("HaircutRate");
            dsReportAutoCall.Tables[0].Columns.Remove("HaircutCalculation");
            dsReportAutoCall.Tables[0].Columns.Remove("IsIRR");
            dsReportAutoCall.Tables[0].Columns.Remove("IRRValue");
            dsReportAutoCall.Tables[0].Columns.Remove("ValueAfterIRR");
            dsReportAutoCall.Tables[0].Columns.Remove("IsAmortization");
            dsReportAutoCall.Tables[0].Columns.Remove("AmortizationValue");
            dsReportAutoCall.Tables[0].Columns.Remove("IsCustomValue");
            dsReportAutoCall.Tables[0].Columns.Remove("CustomValue");
            dsReportAutoCall.Tables[0].Columns.Remove("CreatedBy");

            GridView gv = new GridView();
            gv.DataSource = dsReportAutoCall;
            gv.DataBind();

            Response.ClearContent();
            Response.Buffer = true;
            Response.AddHeader("content-disposition", "attachment; filename=" + "EXPORT_QUATERLY" + ".xls");
            Response.ContentType = "application/ms-excel";
            Response.Charset = "";
            StringWriter sw = new StringWriter();
            HtmlTextWriter htw = new HtmlTextWriter(sw);
            gv.RenderControl(htw);
            Response.Output.Write(sw.ToString());
            Response.Flush();
            Response.End();
        }

        public JsonResult UpdateWeeklyValuation(string gridData, string FromDate, string TenureMultiplier, string ValuationMultiplier, string ValuationStrikeThreshold, string ValuationDivisor)
        {

            //XmlDocument xdoc = JsonConvert.DeserializeXmlNode(gridData);

            var rootJson = "{root1:" + gridData + "}";
            XmlDocument xdoc = JsonConvert.DeserializeXmlNode(rootJson, "root2");

            UserMaster objUserMaster = new UserMaster();
            objUserMaster = (UserMaster)Session["LoggedInUser"];

            General.ExecuteDataSet("SP_WEEKLY_PERIODIC_INSERT_XML", xdoc.InnerXml, objUserMaster.UserID, Convert.ToDateTime(FromDate), Convert.ToDouble(TenureMultiplier), Convert.ToDouble(ValuationMultiplier), Convert.ToDouble(ValuationStrikeThreshold), Convert.ToDouble(ValuationDivisor));

            //objSP_PRICINGEntities.SP_Insert_XML(xdoc.InnerXml, objUserMaster.UserID);
            return Json("");
        }

        public JsonResult UpdateFortnightValuation(string gridData, string FromDate, string TenureMultiplier, string ValuationMultiplier, string ValuationStrikeThreshold, string ValuationDivisor)
        {

            //XmlDocument xdoc = JsonConvert.DeserializeXmlNode(gridData);

            var rootJson = "{root1:" + gridData + "}";
            XmlDocument xdoc = JsonConvert.DeserializeXmlNode(rootJson, "root2");

            UserMaster objUserMaster = new UserMaster();
            objUserMaster = (UserMaster)Session["LoggedInUser"];

            General.ExecuteDataSet("SP_FORTNIGHT_XML", xdoc.InnerXml, objUserMaster.UserID, Convert.ToDateTime(FromDate), Convert.ToDouble(TenureMultiplier), Convert.ToDouble(ValuationMultiplier), Convert.ToDouble(ValuationStrikeThreshold), Convert.ToDouble(ValuationDivisor));

            return Json("");
        }

        public JsonResult UpdateQuaterlyValuation(string gridData, string FromDate, string TenureMultiplier, string ValuationMultiplier, string ValuationStrikeThreshold, string ValuationDivisor)
        {

            //XmlDocument xdoc = JsonConvert.DeserializeXmlNode(gridData);

            var rootJson = "{root1:" + gridData + "}";
            XmlDocument xdoc = JsonConvert.DeserializeXmlNode(rootJson, "root2");

            UserMaster objUserMaster = new UserMaster();
            objUserMaster = (UserMaster)Session["LoggedInUser"];

            General.ExecuteDataSet("SP_QUATERLY_XML", xdoc.InnerXml, objUserMaster.UserID, Convert.ToDateTime(FromDate), Convert.ToDouble(TenureMultiplier), Convert.ToDouble(ValuationMultiplier), Convert.ToDouble(ValuationStrikeThreshold), Convert.ToDouble(ValuationDivisor));

            return Json("");
        }

        public JsonResult OnHairCut(string TenureMultiplier, string ValuationMultiplier, string ValuationStrikeThreshold, string ValuationDivisor, string RedemptionDate, string ProductValue, string ValuationDate, string isChecked)
        {
            double dblTenureMultiplier = Convert.ToDouble(TenureMultiplier) / 100;
            double dblValuationMultiplier = Convert.ToDouble(ValuationMultiplier);
            double dblValuationStrikeThreshold = Convert.ToDouble(ValuationStrikeThreshold);
            double dblValuationDivisor = Convert.ToDouble(ValuationDivisor);
            double dblProductValue = Convert.ToDouble(ProductValue);

            DateTime dtRedemptionDate = Convert.ToDateTime(RedemptionDate);
            DateTime dtValuationDate = Convert.ToDateTime(ValuationDate);
            //DateTime dtValuationDate = DateTime.Now;

            TimeSpan DaysDiff = dtRedemptionDate - dtValuationDate;
            Int32 intDateDiffInDays = Convert.ToInt32(DaysDiff.TotalDays);

            var HaircutRate = Math.Max(0, (intDateDiffInDays * 1.00 / 365 * 1.00) * (dblTenureMultiplier + dblValuationMultiplier * Math.Max((dblProductValue * 1.00 / 100 * 1.00) - dblValuationStrikeThreshold, 0) * 1.00 / dblValuationDivisor * 1.00));
            var HaircutCalculation = 0.00;

            if (dblProductValue > 108)
                if (isChecked.ToUpper() == "TRUE")
                    HaircutCalculation = (dblProductValue + (dblProductValue * HaircutRate));
                else
                    HaircutCalculation = (dblProductValue - (dblProductValue * HaircutRate));
            else
                HaircutCalculation = dblProductValue;

            var FinalData = HaircutRate + "|" + HaircutCalculation;

            return Json(FinalData);
        }

        public JsonResult OnIRR(string TenureMultiplier, string ValuationMultiplier, string ValuationStrikeThreshold, string ValuationDivisor, string RedemptionDate, string ProductValue, string InitialFixingDates, string FixedCoupon, string ProductType)
        {
            double dblTenureMultiplier = Convert.ToDouble(TenureMultiplier);
            double dblValuationMultiplier = Convert.ToDouble(ValuationMultiplier);
            double dblValuationStrikeThreshold = Convert.ToDouble(ValuationStrikeThreshold);
            double dblValuationDivisor = Convert.ToDouble(ValuationDivisor);
            double dblProductValue = Convert.ToDouble(ProductValue);
            double dblFixedCoupon = Convert.ToDouble(FixedCoupon);

            DateTime dtRedemptionDate = Convert.ToDateTime(RedemptionDate);
            DateTime dtInitialFixingDates = Convert.ToDateTime(InitialFixingDates);
            DateTime dtValuationDate = DateTime.Now;
            TimeSpan DaysDiff = dtRedemptionDate - dtInitialFixingDates;
            Int32 intDateDiffInDays = Convert.ToInt32(DaysDiff.TotalDays);

            TimeSpan DaysDiff1 = dtValuationDate - dtInitialFixingDates;
            Int32 intDateDiffInDays1 = Convert.ToInt32(DaysDiff1.TotalDays);

            var IRR = Math.Pow((1 + dblFixedCoupon), (365 / (intDateDiffInDays)));
            var IRRValue = Math.Pow((1 + IRR), ((intDateDiffInDays1) / 365));
            var ValueAfterIRR = 0.00;

            if (ProductType == "Fixed Coupon" || ProductType == "Fixed MLD")
                ValueAfterIRR = dblProductValue - Convert.ToDouble(IRRValue);
            else
                ValueAfterIRR = dblProductValue;

            var FinalData = IRRValue + "|" + ValueAfterIRR;

            return Json(FinalData);
        }

        public JsonResult OnBuiltAMRT(string TenureMultiplier, string ValuationMultiplier, string ValuationStrikeThreshold, string ValuationDivisor, string RedemptionDate, string ProductValue, string InitialFixingDates, string InbuiltCharges)
        {
            double dblTenureMultiplier = Convert.ToDouble(TenureMultiplier);
            double dblValuationMultiplier = Convert.ToDouble(ValuationMultiplier);
            double dblValuationStrikeThreshold = Convert.ToDouble(ValuationStrikeThreshold);
            double dblValuationDivisor = Convert.ToDouble(ValuationDivisor);
            double dblProductValue = Convert.ToDouble(ProductValue);
            double dblInbuiltCharges = Convert.ToDouble(InbuiltCharges);

            DateTime dtRedemptionDate = Convert.ToDateTime(RedemptionDate);
            DateTime dtInitialFixingDates = Convert.ToDateTime(InitialFixingDates);
            DateTime dtValuationDate = DateTime.Now;
            TimeSpan DaysDiff = dtRedemptionDate - dtInitialFixingDates;
            Int32 intDateDiffInDays = Convert.ToInt32(DaysDiff.TotalDays);

            TimeSpan DaysDiff1 = dtValuationDate - dtInitialFixingDates;
            Int32 intDateDiffInDays1 = Convert.ToInt32(DaysDiff1.TotalDays);

            var BuiltAMRT = dblProductValue + dblInbuiltCharges - dblInbuiltCharges * Math.Min(1, (intDateDiffInDays1) / (intDateDiffInDays));

            return Json(BuiltAMRT);
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
                objLoginController.LogError(ex.Message, ex.StackTrace, "PeriodicValuationController", "ValidateSession", -1);
                return false;
            }
        }
    }
}
