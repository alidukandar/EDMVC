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
using System.Xml;
using Newtonsoft.Json.Linq;

namespace SPPricing.Controllers
{
    public class UnwindQuotesController : Controller
    {
        //
        // GET: /UnwindQuotes/

        SP_PRICINGEntities objSP_PRICINGEntities = new SP_PRICINGEntities();

        public ActionResult Index()
        {
            return View();
        }

        [HttpGet]
        public ActionResult UnwindQuotes(string ProductID, string ValDate)
        {
            try
            {
                if (ValidateSession())
                {
                    #region Menu Access By on Role

                    Int32 intResult = 0;
                    // bool PPorNonPP = false;

                    UserMaster objUserMaster = new UserMaster();
                    objUserMaster = (UserMaster)Session["LoggedInUser"];

                    var Result = objSP_PRICINGEntities.VALIDATE_MENU_BY_ROLE(objUserMaster.RoleID, "UQ");
                    intResult = Convert.ToInt32(Result.SingleOrDefault());

                    if (intResult == 0)
                        return RedirectToAction("UserNotAuthorize", "Login");

                    #endregion

                    Session.Remove("Approach1UnwindDetails");

                    UnwindQuotes objUnwindQuotes = new UnwindQuotes();

                    if (ProductID != "" && ProductID != null)
                    {
                        ObjectResult<UnwindQuotesEditResult> objUnwindQuoteDetailsResult = objSP_PRICINGEntities.FETCH_UNWIND_QUOTES_EDIT_DETAILS(ProductID, Convert.ToDateTime(ValDate));
                        List<UnwindQuotesEditResult> UnwindQuoteDetailsResultList = objUnwindQuoteDetailsResult.ToList();

                        General.ReflectSingleData(objUnwindQuotes, UnwindQuoteDetailsResultList[0]);

                        ViewBag.IsCalculate = "True";
                    }
                    else
                    {
                        objUnwindQuotes.ValDate = DateTime.Now.Date;
                    }
                    return View(objUnwindQuotes);
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
                LogError(ex.Message, ex.StackTrace, "UnwindQuotesController", "UnwindQuotes Get", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        [HttpPost, ValidateInput(false)]
        public ActionResult UnwindQuotes(UnwindQuotes objUnwindQuotes, string Command, string ProductID)
        {
            try
            {
                if (ValidateSession())
                {
                    if (Command == "ExportToExcel")
                    {
                        List<UnwindQuotes> UnwindQuotesApproach1List = new List<UnwindQuotes>();
                        List<UnwindQuotes> UnwindQuotesApproach2List = new List<UnwindQuotes>();

                        if (objUnwindQuotes.Approach1Data != "")
                        {
                            /* Parse the underlying array after removing the opening and closing braces */
                            var array = JArray.Parse(objUnwindQuotes.Approach1Data.Trim('{', '}'));

                            for (int i = 0; i < array.Count; i++)
                            {
                                UnwindQuotes oUnwindQuotes = new UnwindQuotes();
                                oUnwindQuotes = array[i].ToObject<UnwindQuotes>();
                                UnwindQuotesApproach1List.Add(oUnwindQuotes);
                            }
                        }

                        if (objUnwindQuotes.Approach2Data != "")
                        {
                            /* Parse the underlying array after removing the opening and closing braces */
                            var array = JArray.Parse(objUnwindQuotes.Approach2Data.Trim('{', '}'));

                            for (int i = 0; i < array.Count; i++)
                            {
                                UnwindQuotes oUnwindQuotes = new UnwindQuotes();
                                oUnwindQuotes = array[i].ToObject<UnwindQuotes>();
                                UnwindQuotesApproach2List.Add(oUnwindQuotes);
                            }
                        }

                        string strTemplateFilePath = Server.MapPath("~/Templates");
                        string strTemplateFileName = strTemplateFilePath + "\\UnwindQuotesTemplate.xlsx";

                        string strTargetFilePath = Server.MapPath("~/OutputFiles");
                        string strTargetFileName = strTargetFilePath + "\\" + objUnwindQuotes.ProductCode + "_UnwindQuotes.xlsx";

                        if (System.IO.File.Exists(strTargetFileName))
                            System.IO.File.Delete(strTargetFileName);

                        FileInfo objTemplateFileInfo = new FileInfo(strTemplateFileName);
                        objTemplateFileInfo.CopyTo(strTargetFileName);

                        FileInfo objTargetFileInfo = new FileInfo(strTargetFileName);

                        using (var xlPackage = new ExcelPackage(objTargetFileInfo))
                        {
                            var worksheet = xlPackage.Workbook.Worksheets["UnwindQuotes"];

                            worksheet.Cell(1, 2).Value = objUnwindQuotes.ProductCode.ToString();
                            worksheet.Cell(1, 4).Value = objUnwindQuotes.LendingRate.ToString();
                            worksheet.Cell(1, 6).Value = objUnwindQuotes.FixedCoupon.ToString();
                            worksheet.Cell(1, 8).Value = objUnwindQuotes.InitailDate.ToString();
                            worksheet.Cell(1, 10).Value = objUnwindQuotes.ExpiryDate.ToString();
                            worksheet.Cell(1, 12).Value = objUnwindQuotes.ProductType.ToString();

                            worksheet.Cell(3, 2).Value = objUnwindQuotes.PreMISValue.ToString();

                            if (objUnwindQuotes.SpotLevel == null)
                                worksheet.Cell(3, 4).Value = "";
                            else
                                worksheet.Cell(3, 4).Value = objUnwindQuotes.SpotLevel.ToString();

                            if (objUnwindQuotes.LowerSpotLevel == null)
                                worksheet.Cell(3, 6).Value = "";
                            else
                                worksheet.Cell(3, 6).Value = objUnwindQuotes.LowerSpotLevel.ToString();

                            if (objUnwindQuotes.HighSpotLevel == null)
                                worksheet.Cell(3, 8).Value = "";
                            else
                                worksheet.Cell(3, 8).Value = objUnwindQuotes.HighSpotLevel.ToString();

                            if (objUnwindQuotes.MonthsLeft == null)
                                worksheet.Cell(3, 10).Value = "";
                            else
                                worksheet.Cell(3, 10).Value = objUnwindQuotes.MonthsLeft.ToString();

                            if (objUnwindQuotes.TenureMulti == null)
                                worksheet.Cell(3, 12).Value = "";
                            else
                                worksheet.Cell(3, 12).Value = objUnwindQuotes.TenureMulti.ToString();


                            if (objUnwindQuotes.ValuationMulti == null)
                                worksheet.Cell(5, 2).Value = "";
                            else
                                worksheet.Cell(5, 2).Value = objUnwindQuotes.ValuationMulti.ToString();

                            if (objUnwindQuotes.StrikeThreshold == null)
                                worksheet.Cell(5, 4).Value = "";
                            else
                                worksheet.Cell(5, 4).Value = objUnwindQuotes.StrikeThreshold.ToString();

                            if (objUnwindQuotes.ValuationDivisor == null)
                                worksheet.Cell(5, 6).Value = "";
                            else
                                worksheet.Cell(5, 6).Value = objUnwindQuotes.ValuationDivisor.ToString();

                            if (objUnwindQuotes.Haircut == null)
                                worksheet.Cell(5, 8).Value = "";
                            else
                                worksheet.Cell(5, 8).Value = objUnwindQuotes.Haircut.ToString();

                            worksheet.Cell(5, 10).Value = objUnwindQuotes.Approach1.ToString();
                            worksheet.Cell(5, 12).Value = objUnwindQuotes.Approach2.ToString();

                            worksheet.Cell(7, 2).Value = objUnwindQuotes.MinimumValue.ToString();
                            worksheet.Cell(7, 4).Value = objUnwindQuotes.DaysPassed.ToString();
                            worksheet.Cell(7, 6).Value = objUnwindQuotes.DaysLeft.ToString();
                            worksheet.Cell(7, 8).Value = objUnwindQuotes.MonthsLeft.ToString();
                            worksheet.Cell(7, 10).Value = objUnwindQuotes.ValuationDate.ToString();
                            worksheet.Cell(7, 12).Value = objUnwindQuotes.TenureExpired.ToString();

                            worksheet.Cell(11, 2).Value = objUnwindQuotes.BondValueAp1.ToString();
                            worksheet.Cell(11, 4).Value = objUnwindQuotes.OptionValueAp1.ToString();
                            worksheet.Cell(11, 6).Value = objUnwindQuotes.AUMAp1.ToString();
                            worksheet.Cell(11, 8).Value = objUnwindQuotes.TotalAp1.ToString();

                            Int32 intStart = 14;

                            if (UnwindQuotesApproach1List != null && UnwindQuotesApproach1List.Count > 0)
                            {
                                foreach (UnwindQuotes oUnwindQuotes in UnwindQuotesApproach1List)
                                {
                                    worksheet.Cell(intStart, 1).Value = oUnwindQuotes.ProductCode.ToString();
                                    worksheet.Cell(intStart, 2).Value = oUnwindQuotes.EnumDescription.ToString();
                                    worksheet.Cell(intStart, 3).Value = oUnwindQuotes.InstrumentType.ToString();
                                    worksheet.Cell(intStart, 4).Value = oUnwindQuotes.Underlying.ToString();
                                    worksheet.Cell(intStart, 5).Value = oUnwindQuotes.ExpiryDateString.ToString();
                                    worksheet.Cell(intStart, 6).Value = oUnwindQuotes.Quantity.ToString();
                                    worksheet.Cell(intStart, 7).Value = oUnwindQuotes.Strike.ToString();
                                    worksheet.Cell(intStart, 8).Value = oUnwindQuotes.Spot.ToString();
                                    worksheet.Cell(intStart, 9).Value = oUnwindQuotes.Price.ToString();
                                    worksheet.Cell(intStart, 10).Value = oUnwindQuotes.OptionValue.ToString();
                                    worksheet.Cell(intStart, 11).Value = oUnwindQuotes.Days.ToString();
                                    worksheet.Cell(intStart, 12).Value = oUnwindQuotes.AutoCall.ToString();
                                    worksheet.Cell(intStart, 13).Value = oUnwindQuotes.BarrierLevel.ToString();
                                    worksheet.Cell(intStart, 14).Value = oUnwindQuotes.Payout.ToString();
                                    worksheet.Cell(intStart, 15).Value = oUnwindQuotes.AppliedIV.ToString();
                                    // worksheet.Cell(intStart, 16).Value = objUnwindQuotes.CustomIV.ToString();
                                    worksheet.Cell(intStart, 17).Value = oUnwindQuotes.AppliedRC.ToString();
                                    // worksheet.Cell(intStart, 18).Value = objUnwindQuotes.CustomRC.ToString();
                                    worksheet.Cell(intStart, 19).Value = oUnwindQuotes.Status.ToString();

                                    intStart = intStart + 1;
                                }
                            }

                            intStart = intStart - 14;

                            worksheet.Cell(16 + intStart, 1).Value = "MIS Value (After IV Adjustment)";
                            worksheet.Cell(16 + intStart, 2).Value = objUnwindQuotes.MISValueAp1.ToString();
                            worksheet.Cell(16 + intStart, 3).Value = "Haircut";
                            worksheet.Cell(16 + intStart, 4).Value = objUnwindQuotes.HaircutAp1.ToString();
                            worksheet.Cell(16 + intStart, 5).Value = "Buyback Quote";
                            worksheet.Cell(16 + intStart, 6).Value = objUnwindQuotes.BuybackQuoteAp1.ToString();

                            worksheet.Cell(18 + intStart, 1).Value = "Approach 2 Details";

                            worksheet.Cell(20 + intStart, 1).Value = "Bond Value";
                            worksheet.Cell(20 + intStart, 2).Value = objUnwindQuotes.BondValueAp2.ToString();
                            worksheet.Cell(20 + intStart, 3).Value = "Option Value";
                            worksheet.Cell(20 + intStart, 4).Value = objUnwindQuotes.OptionValueAp2.ToString();
                            worksheet.Cell(20 + intStart, 5).Value = "AUM(Crs)";
                            worksheet.Cell(20 + intStart, 6).Value = objUnwindQuotes.AUMAp2.ToString();
                            worksheet.Cell(20 + intStart, 7).Value = "Total";
                            worksheet.Cell(20 + intStart, 8).Value = objUnwindQuotes.TotalAp2.ToString();

                            worksheet.Cell(22 + intStart, 1).Value = "Product Code";
                            worksheet.Cell(22 + intStart, 2).Value = "Instrument";
                            worksheet.Cell(22 + intStart, 3).Value = "Underlying";
                            worksheet.Cell(22 + intStart, 4).Value = "Expiry";
                            worksheet.Cell(22 + intStart, 5).Value = "Quantity";
                            worksheet.Cell(22 + intStart, 6).Value = "Strike";
                            worksheet.Cell(22 + intStart, 7).Value = "Spot";
                            worksheet.Cell(22 + intStart, 8).Value = "Price";
                            worksheet.Cell(22 + intStart, 9).Value = "Option Value";
                            worksheet.Cell(22 + intStart, 10).Value = "Days";
                            worksheet.Cell(22 + intStart, 11).Value = "Autocall Level";
                            worksheet.Cell(22 + intStart, 12).Value = "Barrier Level";
                            worksheet.Cell(22 + intStart, 13).Value = "Payout";
                            worksheet.Cell(22 + intStart, 14).Value = "Applied IV";
                            worksheet.Cell(22 + intStart, 15).Value = "Applied RC";
                            worksheet.Cell(22 + intStart, 16).Value = "Status";

                            if (UnwindQuotesApproach2List != null && UnwindQuotesApproach2List.Count > 0)
                            {
                                foreach (UnwindQuotes oUnwindQuotes2 in UnwindQuotesApproach2List)
                                {
                                    worksheet.Cell(23 + intStart, 1).Value = oUnwindQuotes2.ProductCode.ToString();
                                    worksheet.Cell(23 + intStart, 2).Value = oUnwindQuotes2.InstrumentType.ToString();
                                    worksheet.Cell(23 + intStart, 3).Value = oUnwindQuotes2.Underlying.ToString();
                                    worksheet.Cell(23 + intStart, 4).Value = oUnwindQuotes2.ExpiryDateString.ToString();
                                    worksheet.Cell(23 + intStart, 5).Value = oUnwindQuotes2.Quantity.ToString();
                                    worksheet.Cell(23 + intStart, 6).Value = oUnwindQuotes2.Strike.ToString();
                                    worksheet.Cell(23 + intStart, 7).Value = oUnwindQuotes2.Spot.ToString();
                                    worksheet.Cell(23 + intStart, 8).Value = oUnwindQuotes2.Price.ToString();
                                    worksheet.Cell(23 + intStart, 9).Value = oUnwindQuotes2.OptionValue.ToString();
                                    worksheet.Cell(23 + intStart, 10).Value = oUnwindQuotes2.Days.ToString();
                                    worksheet.Cell(23 + intStart, 11).Value = oUnwindQuotes2.AutoCall.ToString();
                                    worksheet.Cell(23 + intStart, 12).Value = oUnwindQuotes2.BarrierLevel.ToString();
                                    worksheet.Cell(23 + intStart, 13).Value = oUnwindQuotes2.Payout.ToString();
                                    worksheet.Cell(23 + intStart, 14).Value = oUnwindQuotes2.AppliedIV.ToString();
                                    worksheet.Cell(23 + intStart, 15).Value = oUnwindQuotes2.AppliedRC.ToString();
                                    worksheet.Cell(23 + intStart, 16).Value = oUnwindQuotes2.Status.ToString();

                                    intStart = intStart + 1;
                                }
                            }

                            worksheet.Cell(25 + intStart, 1).Value = "MIS Value (No Adjustment)";
                            worksheet.Cell(25 + intStart, 2).Value = objUnwindQuotes.MISValueAp2.ToString();
                            worksheet.Cell(25 + intStart, 3).Value = "Interest Rate Hit";
                            worksheet.Cell(25 + intStart, 4).Value = objUnwindQuotes.InterestHitAp2.ToString();
                            worksheet.Cell(25 + intStart, 5).Value = "Haircut";
                            worksheet.Cell(25 + intStart, 6).Value = objUnwindQuotes.HaircutAp2.ToString();
                            worksheet.Cell(25 + intStart, 7).Value = "Buyback Quote";
                            worksheet.Cell(25 + intStart, 8).Value = objUnwindQuotes.BuybackQuoteAp2.ToString();

                            worksheet.Cell(27 + intStart, 1).Value = "Attribute";
                            worksheet.Cell(27 + intStart, 2).Value = "Bond";
                            worksheet.Cell(27 + intStart, 3).Value = "Buyback Quote";

                            worksheet.Cell(28 + intStart, 1).Value = "Current Value";
                            worksheet.Cell(28 + intStart, 2).Value = objUnwindQuotes.CurrentValueBond.ToString();
                            worksheet.Cell(28 + intStart, 3).Value = objUnwindQuotes.CurrentValueBBC.ToString();

                            worksheet.Cell(29 + intStart, 1).Value = "Remaining Tenure";
                            worksheet.Cell(29 + intStart, 2).Value = objUnwindQuotes.RemTenureBond.ToString();
                            worksheet.Cell(29 + intStart, 3).Value = objUnwindQuotes.RemTenureBBC.ToString();

                            worksheet.Cell(30 + intStart, 1).Value = "Rate";
                            worksheet.Cell(30 + intStart, 2).Value = objUnwindQuotes.RateBond.ToString();
                            worksheet.Cell(30 + intStart, 3).Value = objUnwindQuotes.RateBBC.ToString();

                            worksheet.Cell(31 + intStart, 1).Value = "Interest";
                            worksheet.Cell(31 + intStart, 2).Value = objUnwindQuotes.InterestBond.ToString();
                            worksheet.Cell(31 + intStart, 3).Value = objUnwindQuotes.InterestBBC.ToString();

                            worksheet.Cell(32 + intStart, 1).Value = "Difference";
                            worksheet.Cell(32 + intStart, 2).Value = objUnwindQuotes.DifferenceBond.ToString();
                            worksheet.Cell(32 + intStart, 3).Value = objUnwindQuotes.DifferenceBBC.ToString();

                            worksheet.Cell(34 + intStart, 1).Value = "Frequency";
                            worksheet.Cell(34 + intStart, 2).Value = "Date";
                            worksheet.Cell(34 + intStart, 3).Value = "Quote";
                            worksheet.Cell(34 + intStart, 4).Value = "Ref Spot";

                            if (objUnwindQuotes.Frequency == null)
                                worksheet.Cell(35, 1).Value = "";
                            else
                                worksheet.Cell(35, 1).Value = Convert.ToString(objUnwindQuotes.Frequency);

                            worksheet.Cell(35, 2).Value = objUnwindQuotes.ValuationDate.ToString();
                            worksheet.Cell(35, 3).Value = objUnwindQuotes.Quote.ToString();
                            worksheet.Cell(35, 4).Value = objUnwindQuotes.RefSpot.ToString();

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
                }

                return View();
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnwindQuotesController", "UnwindQuotes Post", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        public ActionResult AutoCompleteQuoteID(string term)
        {
            try
            {
                if (ValidateSession())
                {
                    List<UnwindQuotes> UnwindQuoteList = new List<UnwindQuotes>();

                    DataSet dsResult = new DataSet();
                    dsResult = General.ExecuteDataSet("FETCH_UNWIND_QUOTES_PRODUCTS");

                    if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                    {
                        foreach (DataRow dr in dsResult.Tables[0].Rows)
                        {
                            UnwindQuotes obj = new UnwindQuotes();

                            obj.ProductCode = Convert.ToString(dr["PRODUCT_CODE"]);
                            UnwindQuoteList.Add(obj);
                        }
                    }

                    var DistinctItems = UnwindQuoteList.GroupBy(x => x.ProductCode).Select(y => y.First());

                    var result = (from objRuleList in DistinctItems
                                  where objRuleList.ProductCode.ToLower().StartsWith(term.ToLower())
                                  select objRuleList);

                    return Json(result);
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
                LogError(ex.Message, ex.StackTrace, "UnwindQuotesController", "AutoCompleteQuoteID", objUserMaster.UserID);
                Session["ErrorData"] = ex.Message;
                return RedirectToAction("ErrorPage", "Login");
            }

        }

        public ActionResult ValDateCalc(string ValDate, string ExpiryDate, string InitialDate)
        {
            try
            {
                if (ValDate != "" && ExpiryDate != "" && InitialDate != "")
                {
                    DateTime valuation = Convert.ToDateTime(ValDate);
                    DateTime Expiry = Convert.ToDateTime(ExpiryDate);
                    DateTime Initial = Convert.ToDateTime(InitialDate);

                    TimeSpan DaysPassed = valuation - Initial;
                    TimeSpan DaysLeft = Expiry - valuation;


                    double TenureExpired = Convert.ToDouble(DaysPassed.TotalDays) / (30.417);
                    double MonthsLeft = Convert.ToDouble(DaysLeft.TotalDays) / (30.417);

                    UnwindQuotes obj = new UnwindQuotes();
                    obj.DaysPassed = Convert.ToInt32(DaysPassed.TotalDays);
                    obj.DaysLeft = Convert.ToInt32(DaysLeft.TotalDays);
                    obj.MonthsLeft = Math.Round(Convert.ToDouble(MonthsLeft), 2);
                    obj.TenureExpired = Math.Round(Convert.ToDouble(TenureExpired), 2);

                    return Json(obj, JsonRequestBehavior.AllowGet);
                }

                return Json("");
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnwindQuotesController", "ValDateCalc", objUserMaster.UserID);
                Session["ErrorData"] = ex.Message;
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        public ActionResult RedemptionDiscountdaysCalc(string RedemptionDate, string OptionDiscountdate)
        {
            try
            {
                if (RedemptionDate != "" && OptionDiscountdate != "")
                {
                    DateTime RedDate = Convert.ToDateTime(RedemptionDate);
                    DateTime OptionDiscDate = Convert.ToDateTime(OptionDiscountdate);

                    TimeSpan DaysDiff = RedDate - OptionDiscDate;


                    double RedemptionDisc = Convert.ToDouble(DaysDiff.TotalDays) / (365);

                    UnwindQuotes obj = new UnwindQuotes();
                    obj.DiffDatecalc = Convert.ToInt32(RedemptionDisc);

                    return Json(obj, JsonRequestBehavior.AllowGet);
                }

                return Json("");
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnwindQuotesController", "RedemptionDiscountdaysCalc", objUserMaster.UserID);
                Session["ErrorData"] = ex.Message;
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        public JsonResult FetchUnwindData(string ProductCode, string ValuationDate)
        {
            try
            {
                if (ValuationDate == "" || ValuationDate == null)
                    ValuationDate = "1900-01-01";

                ObjectResult<UnwindQuoteListResult> objUnwindQuoteListResult = objSP_PRICINGEntities.FETCH_UNWIND_QUOTES_LIST(ProductCode, Convert.ToDateTime(ValuationDate));
                List<UnwindQuoteListResult> UnwindQuoteListResultList = objUnwindQuoteListResult.ToList();


                return Json(UnwindQuoteListResultList, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnwindQuotesController", "FetchUnwindData", objUserMaster.UserID);
                return Json("");
            }
        }

        public JsonResult CalculateRedemptionMinusValuationDate(string RedemptionDate, string ValuationDate)
        {
            try
            {
                DateTime dtRedemptionDate = Convert.ToDateTime(RedemptionDate);

                //DateTime dtValuationDate = Convert.ToDateTime(ValuationDate.Substring(6, 4) + "-" + ValuationDate.Substring(0, 2) + "-" + ValuationDate.Substring(3, 2));

                DateTime dtValuationDate = Convert.ToDateTime(ValuationDate);

                TimeSpan DaysDiff = dtRedemptionDate - dtValuationDate;

                return Json(Convert.ToInt32(DaysDiff.TotalDays));
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnwindQuotesController", "CalculateRedemptionMinusValuationDate", objUserMaster.UserID);
                return Json("");
            }
        }

        public JsonResult FetchTotalOptionPrice(string ProductCode, string ValuationDate)
        {
            try
            {
                ObjectResult<TotalOptionPriceResult> objTotalOptionPriceResult = objSP_PRICINGEntities.FETCH_TOTAL_OPTION_PRICE(ProductCode, Convert.ToDateTime(ValuationDate));
                List<TotalOptionPriceResult> TotalOptionPriceResultList = objTotalOptionPriceResult.ToList();

                return Json(TotalOptionPriceResultList, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnwindQuotesController", "CalculateRedemptionMinusValuationDate", objUserMaster.UserID);
                return Json("");
            }
        }

        public JsonResult CalculateValuationMinusInitialDate(string ValuationDate, string InitialDate)
        {
            try
            {
                DateTime dtInitialDate = Convert.ToDateTime(InitialDate);
                DateTime dtValuationDate = Convert.ToDateTime(ValuationDate);

                TimeSpan DaysDiff = dtValuationDate - dtInitialDate;

                return Json(Convert.ToInt32(DaysDiff.TotalDays));
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnwindQuotesController", "CalculateRedemptionMinusValuationDate", objUserMaster.UserID);
                return Json("");
            }
        }

        public JsonResult CalculateRedemptionMinusDiscountOptionTillDate(string RedemptionDate, string DiscountOptionTillDate)
        {
            try
            {
                DateTime dtRedemptionDate = Convert.ToDateTime(RedemptionDate);

                //DateTime dtValuationDate = Convert.ToDateTime(ValuationDate.Substring(6, 4) + "-" + ValuationDate.Substring(0, 2) + "-" + ValuationDate.Substring(3, 2));

                DateTime dtValuationDate = Convert.ToDateTime(DiscountOptionTillDate);

                TimeSpan DaysDiff = dtRedemptionDate - dtValuationDate;

                return Json(Convert.ToInt32(DaysDiff.TotalDays));
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnwindQuotesController", "CalculateRedemptionMinusValuationDate", objUserMaster.UserID);
                return Json("");
            }
        }

        public JsonResult CalculateExpiryMinusValuationDate(string ExpiryDate, string ValuationDate)
        {
            try
            {
                DateTime dtExpiryDate = Convert.ToDateTime(ExpiryDate);
                DateTime dtValuationDate = Convert.ToDateTime(ValuationDate);

                TimeSpan DaysDiff = dtExpiryDate - dtValuationDate;

                return Json(Math.Abs(Convert.ToInt32(DaysDiff.TotalDays)));
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnwindQuotesController", "CalculateRedemptionMinusValuationDate", objUserMaster.UserID);
                return Json("");
            }
        }

        public JsonResult CalculateDiscountingRate(string ActualDeployemntRateAtInitiation, string RedemptionDate, string ValuationDate)
        {
            try
            {
                double dblActualDeployemntRateAtInitiation = Convert.ToDouble(ActualDeployemntRateAtInitiation);
                DateTime dtRedemptionDate = Convert.ToDateTime(RedemptionDate);
                DateTime dtValuationDate = Convert.ToDateTime(ValuationDate);

                TimeSpan DaysDiff = dtRedemptionDate - dtValuationDate;

                double dblActualDeployemntRate = 0;

                ObjectResult<FetchActualDeploymentRateResult> FetchActualDeploymentRateResult = objSP_PRICINGEntities.SP_FETCH_ACTUAL_DEPLOYMENT_RATE(Convert.ToInt32(DaysDiff.TotalDays), dtValuationDate);
                List<FetchActualDeploymentRateResult> FetchActualDeploymentRateResultList = FetchActualDeploymentRateResult.ToList();

                //var Count = objSP_PRICINGEntities.SP_FETCH_ACTUAL_DEPLOYMENT_RATE(Convert.ToInt32(DaysDiff.TotalDays), dtValuationDate);
                dblActualDeployemntRate = Convert.ToDouble(FetchActualDeploymentRateResultList[0].DeploymentRate);

                double dblDiscountingRate = 0;

                if (dblActualDeployemntRateAtInitiation > dblActualDeployemntRate)
                    dblDiscountingRate = dblActualDeployemntRateAtInitiation;
                else
                    dblDiscountingRate = dblActualDeployemntRate;

                return Json(dblDiscountingRate);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnwindQuotesController", "CalculateDiscountingRate", objUserMaster.UserID);
                return Json("");
            }
        }

        public JsonResult CalculateActualDeployemntRate(string DaysLeft)
        {
            try
            {
                double dblActualDeployemntRate = 0;

                ObjectResult<FetchActualDeploymentRateResult> FetchActualDeploymentRateResult = objSP_PRICINGEntities.SP_FETCH_ACTUAL_DEPLOYMENT_RATE(Convert.ToInt32(DaysLeft), DateTime.Now);
                List<FetchActualDeploymentRateResult> FetchActualDeploymentRateResultList = FetchActualDeploymentRateResult.ToList();

                dblActualDeployemntRate = Convert.ToDouble(FetchActualDeploymentRateResultList[0].DeploymentRate);

                return Json(dblActualDeployemntRate);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnwindQuotesController", "CalculateActualDeployemntRate", objUserMaster.UserID);
                return Json("");
            }
        }

        public JsonResult FetchUnwindApproach1(string ProductCode, string Spot, string ValDate)
        {
            try
            {
                DataSet dsResult = new DataSet();

                if (Session["Approach1UnwindDetails"] != null)
                    dsResult = (DataSet)Session["Approach1UnwindDetails"];
                else if (ProductCode != "" && Spot != "")
                    dsResult = General.ExecuteDataSet("FETCH_UNWIND_APPROACH1", ProductCode, Convert.ToDouble(Spot), Convert.ToDateTime(ValDate));

                List<UnwindQuotes> UnwindQuotesList = new List<UnwindQuotes>();

                if (dsResult != null && dsResult.Tables.Count == 2)
                {
                    foreach (DataRow dr in dsResult.Tables[1].Rows)
                    {
                        UnwindQuotes objUnwindQuotes = new UnwindQuotes();

                        General.ReflectSingleData(dr, objUnwindQuotes);

                        objUnwindQuotes.RowNum = Convert.ToInt32(dr["RowNum"]);
                        objUnwindQuotes.ProductCode = Convert.ToString(dr["ProductCode"]);
                        objUnwindQuotes.InstrumentType = Convert.ToString(dr["InstrumentType"]);
                        objUnwindQuotes.EnumDescription = Convert.ToString(dr["EnumDescription"]);
                        objUnwindQuotes.Underlying = Convert.ToString(dr["Underlying"]);
                        objUnwindQuotes.ExpiryDateString = Convert.ToString(dr["ExpiryDateString"]);
                        objUnwindQuotes.DirectionOfOption = Convert.ToString(dr["DirectionOfOption"]);
                        objUnwindQuotes.AUM = Convert.ToString(dr["AUM"]);
                        objUnwindQuotes.FinalFixingDateCount = Convert.ToString(dr["FinalFixingDateCount"]);
                        objUnwindQuotes.Participation = Convert.ToString(dr["Participation"]);
                        objUnwindQuotes.Quantity = Convert.ToString(dr["Quantity"]);
                        objUnwindQuotes.Price = Convert.ToString(dr["Price"]);
                        objUnwindQuotes.OptionValue = Convert.ToString(dr["OptionValue"]);
                        objUnwindQuotes.Strike = Convert.ToString(dr["Strike"]);
                        objUnwindQuotes.Spot = Convert.ToString(dr["Spot"]);
                        objUnwindQuotes.StrikeMultiplier = Convert.ToString(dr["StrikeMultiplier"]);
                        objUnwindQuotes.Days = Convert.ToString(dr["Days"]);
                        objUnwindQuotes.AutoCall = Convert.ToString(dr["AutoCall"]);
                        objUnwindQuotes.BarrierLevel = Convert.ToString(dr["BarrierLevel"]);
                        objUnwindQuotes.Payout = Convert.ToString(dr["Payout"]);
                        objUnwindQuotes.AppliedIV = Convert.ToString(dr["AppliedIV"]);
                        objUnwindQuotes.AppliedRC = Convert.ToString(dr["AppliedRC"]);
                        objUnwindQuotes.UpdatedCustomIV = "";
                        objUnwindQuotes.UpdatedCustomRC = "";
                        objUnwindQuotes.Status = Convert.ToString(dr["Status"]);

                        UnwindQuotesList.Add(objUnwindQuotes);
                    }
                }

                var UnwindApproachData = UnwindQuotesList.ToList();
                return Json(UnwindApproachData, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnwindQuotesController", "FetchUnwindApproach1", objUserMaster.UserID);
                return Json("");
            }
        }


        public JsonResult FetchUnwindApproach1Error(string ProductCode, string Spot, string ValDate)
        {
            try
            {
                if (ProductCode != "" && Spot != "")
                {

                    DataSet dsResult = General.ExecuteDataSet("FETCH_UNWIND_APPROACH1", ProductCode, Convert.ToDouble(Spot), Convert.ToDateTime(ValDate));

                    if (ProductCode != "")
                        if (dsResult != null && dsResult.Tables.Count == 2)
                            Session["Approach1UnwindDetails"] = dsResult;

                    List<UnwindQuotes> UnwindQuotesList = new List<UnwindQuotes>();

                    foreach (DataRow dr in dsResult.Tables[0].Rows)
                    {
                        UnwindQuotes objUnwindQuotes = new UnwindQuotes();
                        //General.ReflectSingleData(objUnwindQuotes, dr);

                        objUnwindQuotes.ProductCode = Convert.ToString(dr["ProductCode"]);
                        objUnwindQuotes.InstrumentType = Convert.ToString(dr["InstrumentType"]);
                        objUnwindQuotes.EnumDescription = Convert.ToString(dr["EnumDescription"]);
                        objUnwindQuotes.Underlying = Convert.ToString(dr["Underlying"]);
                        objUnwindQuotes.ExpiryDateString = Convert.ToString(dr["ExpiryDateString"]);
                        objUnwindQuotes.DirectionOfOption = Convert.ToString(dr["DirectionOfOption"]);

                        UnwindQuotesList.Add(objUnwindQuotes);
                    }

                    var UnwindApproachData = UnwindQuotesList.ToList();
                    return Json(UnwindApproachData, JsonRequestBehavior.AllowGet);
                }
                return Json("");
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnwindQuotesController", "FetchUnwindApproach1", objUserMaster.UserID);
                return Json("");
            }
        }

        public JsonResult FetchUnwindApproach2(string ProductCode, string Spot, string ValDate)
        {
            try
            {
                DataSet dsResult = new DataSet();

                if (Session["Approach1UnwindDetails"] != null)
                    dsResult = (DataSet)Session["Approach1UnwindDetails"];
                else if (ProductCode != "" && Spot != "")
                    dsResult = General.ExecuteDataSet("FETCH_UNWIND_APPROACH1", ProductCode, Convert.ToDouble(Spot), Convert.ToDateTime(ValDate));

                List<UnwindQuotes> UnwindQuotesList = new List<UnwindQuotes>();

                if (dsResult != null && dsResult.Tables.Count == 2)
                {
                    foreach (DataRow dr in dsResult.Tables[1].Rows)
                    {
                        UnwindQuotes objUnwindQuotes = new UnwindQuotes();

                        General.ReflectSingleData(dr, objUnwindQuotes);

                        objUnwindQuotes.RowNum = Convert.ToInt32(dr["RowNum"]);
                        objUnwindQuotes.ProductCode = Convert.ToString(dr["ProductCode"]);
                        objUnwindQuotes.InstrumentType = Convert.ToString(dr["InstrumentType"]);
                        objUnwindQuotes.EnumDescription = Convert.ToString(dr["EnumDescription"]);
                        objUnwindQuotes.Underlying = Convert.ToString(dr["Underlying"]);
                        objUnwindQuotes.ExpiryDateString = Convert.ToString(dr["ExpiryDateString"]);
                        objUnwindQuotes.DirectionOfOption = Convert.ToString(dr["DirectionOfOption"]);
                        objUnwindQuotes.AUM = Convert.ToString(dr["AUM"]);
                        objUnwindQuotes.FinalFixingDateCount = Convert.ToString(dr["FinalFixingDateCount"]);
                        objUnwindQuotes.Participation = Convert.ToString(dr["Participation"]);
                        objUnwindQuotes.Quantity = Convert.ToString(dr["Quantity"]);
                        objUnwindQuotes.Price = Convert.ToString(dr["Price"]);
                        objUnwindQuotes.OptionValue = Convert.ToString(dr["OptionValue"]);
                        objUnwindQuotes.Strike = Convert.ToString(dr["Strike"]);
                        objUnwindQuotes.Spot = Convert.ToString(dr["Spot"]);
                        objUnwindQuotes.StrikeMultiplier = Convert.ToString(dr["StrikeMultiplier"]);
                        objUnwindQuotes.Days = Convert.ToString(dr["Days"]);
                        objUnwindQuotes.AutoCall = Convert.ToString(dr["AutoCall"]);
                        objUnwindQuotes.BarrierLevel = Convert.ToString(dr["BarrierLevel"]);
                        objUnwindQuotes.Payout = Convert.ToString(dr["Payout"]);
                        objUnwindQuotes.AppliedIV = Convert.ToString(dr["AppliedIV"]);
                        objUnwindQuotes.AppliedRC = Convert.ToString(dr["AppliedRC"]);
                        objUnwindQuotes.UpdatedCustomIV = "";
                        objUnwindQuotes.UpdatedCustomRC = "";
                        objUnwindQuotes.Status = Convert.ToString(dr["Status"]);

                        UnwindQuotesList.Add(objUnwindQuotes);
                    }
                }

                var UnwindApproachData = UnwindQuotesList.ToList();
                return Json(UnwindApproachData, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnwindQuotesController", "FetchUnwindApproach1", objUserMaster.UserID);
                return Json("");
            }
        }


        public JsonResult FetchUnwindApproach2Error(string ProductCode, string Spot, string ValDate)
        {
            try
            {
                if (ProductCode != "" && Spot != "")
                {

                    DataSet dsResult = General.ExecuteDataSet("FETCH_UNWIND_APPROACH1", ProductCode, Convert.ToDouble(Spot), Convert.ToDateTime(ValDate));

                    if (dsResult != null && dsResult.Tables.Count == 2)
                        Session["Approach2UnwindDetails"] = dsResult;

                    List<UnwindQuotes> UnwindQuotesList = new List<UnwindQuotes>();

                    foreach (DataRow dr in dsResult.Tables[0].Rows)
                    {
                        UnwindQuotes objUnwindQuotes = new UnwindQuotes();
                        //General.ReflectSingleData(objUnwindQuotes, dr);

                        objUnwindQuotes.ProductCode = Convert.ToString(dr["ProductCode"]);
                        objUnwindQuotes.InstrumentType = Convert.ToString(dr["InstrumentType"]);
                        objUnwindQuotes.EnumDescription = Convert.ToString(dr["EnumDescription"]);
                        objUnwindQuotes.Underlying = Convert.ToString(dr["Underlying"]);
                        objUnwindQuotes.ExpiryDateString = Convert.ToString(dr["ExpiryDateString"]);
                        objUnwindQuotes.DirectionOfOption = Convert.ToString(dr["DirectionOfOption"]);

                        UnwindQuotesList.Add(objUnwindQuotes);
                    }

                    var UnwindApproachData = UnwindQuotesList.ToList();
                    return Json(UnwindApproachData, JsonRequestBehavior.AllowGet);
                }
                return Json("");
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnwindQuotesController", "FetchUnwindApproach1", objUserMaster.UserID);
                return Json("");
            }
        }

        public void LogError(string strErrorDescription, string strStackTrace, string strClassName, string strMethodName, Int32 intUserId)
        {
            SP_PRICINGEntities objSP_PRICINGEntities = new SP_PRICINGEntities();
            var Count = objSP_PRICINGEntities.SP_ERROR_LOG(strErrorDescription, strStackTrace, strClassName, strMethodName, intUserId);
        }

        public ActionResult ManageUnwind(string ID, string ProductCode, string ProductType, string FixedCoupon,
            string LendingRate, string PreMISValue, string InitailDate, string ExpiryDate,
            string SpotLevel, string HighSpotLevel, string LowerSpotLevel, string MonthLeft, string Haircut, string TenureMulti, string ValuationMulti,
            string StrikeThreshold, string ValuationDivisor, string Approach1, string Approach2, string MinimumValue, string DaysPassed, string DaysLeft, string MonthsLeft,
            string ValuationDate, string TenureExpired, string BondValueAp1, string OptionValueAp1, string AUMAp1, string TotalAp1,
            string BondValueAp2, string OptionValueAp2, string AUMAp2, string TotalAp2, string MISValueAp1, string HaircutAp1,
            string BuybackQuoteAp1, string MISValueAp2, string InterestHitAp2, string HaircutAp2,
            string BuybackQuoteAp2, string CurrentValueBond, string CurrentValueBBC, string RemTenureBond, string RemTenureBBC, string RateBond, string RateBBC, string InterestBond,
            string InterestBBC, string DifferenceBond, string DifferenceBBC, string Frequency, string ValDate, string Quote, string RefSpot, string GridData1, string GridData2)
        {
            LoginController objLoginController = new LoginController();
            try
            {
                if (ValidateSession())
                {

                    var rootJson = "{root1:" + GridData1 + "}";
                    XmlDocument xdoc = JsonConvert.DeserializeXmlNode(rootJson, "root2");

                    var rootJson1 = "{root1:" + GridData2 + "}";
                    XmlDocument xdoc1 = JsonConvert.DeserializeXmlNode(rootJson1, "root2");



                    if (LendingRate == "")
                        LendingRate = "0";
                    if (PreMISValue == "")
                        PreMISValue = "0";
                    if (SpotLevel == "")
                        SpotLevel = "0";
                    if (HighSpotLevel == "")
                        HighSpotLevel = "0";
                    if (LowerSpotLevel == "")
                        LowerSpotLevel = "0";
                    if (MonthLeft == "")
                        MonthLeft = "0";
                    if (Haircut == "")
                        Haircut = "0";
                    if (TenureMulti == "")
                        TenureMulti = "0";
                    if (ValuationMulti == "")
                        ValuationMulti = "0";
                    if (StrikeThreshold == "")
                        StrikeThreshold = "0";
                    if (ValuationDivisor == "")
                        ValuationDivisor = "0";
                    if (Approach1 == "")
                        Approach1 = "0";
                    if (Approach2 == "")
                        Approach2 = "0";
                    if (MinimumValue == "")
                        MinimumValue = "0";
                    if (DaysPassed == "")
                        DaysPassed = "0";
                    if (DaysLeft == "")
                        DaysLeft = "0";
                    if (MonthsLeft == "")
                        MonthsLeft = "0";
                    if (BondValueAp1 == "")
                        BondValueAp1 = "0";
                    if (TotalAp2 == "")
                        TotalAp2 = "0";
                    if (MISValueAp2 == "")
                        MISValueAp2 = "0";
                    if (InterestBond == "")
                        InterestBond = "0";
                    if (OptionValueAp1 == "")
                        OptionValueAp1 = "0";
                    if (MISValueAp1 == "")
                        MISValueAp1 = "0";
                    if (InterestHitAp2 == "")
                        InterestHitAp2 = "0";
                    if (InterestBBC == "")
                        InterestBBC = "0";
                    if (AUMAp1 == "")
                        AUMAp1 = "0";
                    if (HaircutAp1 == "")
                        HaircutAp1 = "0";
                    if (HaircutAp2 == "")
                        HaircutAp2 = "0";
                    if (DifferenceBond == "")
                        DifferenceBond = "0";
                    if (TotalAp1 == "")
                        TotalAp1 = "0";
                    if (BuybackQuoteAp1 == "")
                        BuybackQuoteAp1 = "0";
                    if (DifferenceBBC == "")
                        DifferenceBBC = "0";
                    if (BondValueAp1 == "")
                        BondValueAp1 = "0";
                    if (OptionValueAp1 == "")
                        OptionValueAp1 = "0";
                    if (AUMAp1 == "")
                        AUMAp1 = "0";
                    if (TotalAp1 == "")
                        TotalAp1 = "0";
                    if (CurrentValueBond == "")
                        CurrentValueBond = "0";
                    if (CurrentValueBBC == "")
                        CurrentValueBBC = "0";
                    if (RemTenureBond == "")
                        RemTenureBond = "0";
                    if (RemTenureBBC == "")
                        RemTenureBBC = "0";
                    if (RateBond == "")
                        RateBond = "0";
                    if (RateBBC == "")
                        RateBBC = "0";
                    if (Quote == "")
                        Quote = "0";
                    if (RefSpot == "")
                        RefSpot = "0";


                    UserMaster objUserMaster = new UserMaster();
                    objUserMaster = (UserMaster)Session["LoggedInUser"];

                    Int32 intResult = 0;
                    var Count = objSP_PRICINGEntities.SP_MANAGE_TBL_UNWIND_QUOTE(Convert.ToInt32(ID), ProductCode, ProductType, FixedCoupon, Convert.ToDouble(LendingRate), Convert.ToDouble(PreMISValue), Convert.ToDateTime(InitailDate), Convert.ToDateTime(ExpiryDate),
                                                 Convert.ToDouble(SpotLevel), Convert.ToDouble(HighSpotLevel), Convert.ToDouble(LowerSpotLevel), Convert.ToDouble(MonthLeft),
                                                Convert.ToDouble(Haircut), Convert.ToDouble(TenureMulti), Convert.ToDouble(ValuationMulti), Convert.ToDouble(StrikeThreshold), Convert.ToDouble(ValuationDivisor), Convert.ToDouble(Approach1),
                                                Convert.ToDouble(Approach2), Convert.ToDouble(MinimumValue), Convert.ToDouble(DaysPassed), Convert.ToDouble(DaysLeft), Convert.ToDouble(MonthsLeft), Convert.ToDateTime(ValuationDate), Convert.ToDouble(TenureExpired),
                                                Convert.ToDouble(BondValueAp1), Convert.ToDouble(OptionValueAp1), Convert.ToDouble(AUMAp1), Convert.ToDouble(TotalAp1), Convert.ToDouble(BondValueAp2), Convert.ToDouble(OptionValueAp2), Convert.ToDouble(AUMAp2),
                                                Convert.ToDouble(TotalAp2), Convert.ToDouble(MISValueAp1), Convert.ToDouble(HaircutAp1), Convert.ToDouble(BuybackQuoteAp1),
                                                Convert.ToDouble(MISValueAp2), Convert.ToDouble(InterestHitAp2), Convert.ToDouble(HaircutAp2), Convert.ToDouble(BuybackQuoteAp2), Convert.ToDouble(CurrentValueBond), Convert.ToDouble(CurrentValueBBC), Convert.ToDouble(RemTenureBond), Convert.ToDouble(RemTenureBBC), Convert.ToDouble(RateBond), Convert.ToDouble(RateBBC),
                                                Convert.ToDouble(InterestBond), Convert.ToDouble(InterestBBC), Convert.ToDouble(DifferenceBond), Convert.ToDouble(DifferenceBBC), Frequency, Convert.ToDateTime(ValDate), Convert.ToInt32(Quote), Convert.ToInt32(RefSpot), xdoc.InnerXml, xdoc1.InnerXml, objUserMaster.UserID);
                    intResult = Count.SingleOrDefault().Value;

                    return Json(intResult);
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
                LogError(ex.Message, ex.StackTrace, "UnwindQuotesController", "ManageUnwind Get", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
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
                objLoginController.LogError(ex.Message, ex.StackTrace, "UnwindQuotesController", "ValidateSession", -1);
                return false;
            }
        }

        public JsonResult FetchUnwindQuotesList(string ProductCode, string ProductType, string ValuationDate, string IsPreciousRecord)
        {
            try
            {
                List<UnwindApproachResult> UnwindApproachList = new List<UnwindApproachResult>();

                if (ProductCode == "" || ProductCode == "--Select--")
                    ProductCode = "ALL";

                if (ProductType == "" || ProductType == "--Select--")
                    ProductType = "ALL";

                if (ValuationDate == "")
                    ValuationDate = "1900-01-01";

                ObjectResult<UnwindQuoteDetailsResult> objUnwindQuoteDetailsResult = objSP_PRICINGEntities.FETCH_UNWIND_QUOTES_DETAILS(ProductCode, ProductType, Convert.ToDateTime(ValuationDate), Convert.ToBoolean(IsPreciousRecord));
                List<UnwindQuoteDetailsResult> UnwindQuoteDetailsResultList = objUnwindQuoteDetailsResult.ToList();

                return Json(UnwindQuoteDetailsResultList, JsonRequestBehavior.AllowGet);

            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "UnwindQuotesController", "FetchUnwindQuotesList", objUserMaster.UserID);
                return Json("");
            }


        }

        public ActionResult UnwindQuotesList()
        {
            try
            {
                if (ValidateSession())
                {
                    UnwindQuotes obj = new UnwindQuotes();

                    #region Menu Access By on Role

                    Int32 intResult = 0;
                    // bool PPorNonPP = false;

                    UserMaster objUserMaster = new UserMaster();
                    objUserMaster = (UserMaster)Session["LoggedInUser"];

                    var Result = objSP_PRICINGEntities.VALIDATE_MENU_BY_ROLE(objUserMaster.RoleID, "UQL");
                    intResult = Convert.ToInt32(Result.SingleOrDefault());

                    if (intResult == 0)
                        return RedirectToAction("UserNotAuthorize", "Login");

                    #endregion

                    obj.ValDate = DateTime.Now.Date;
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
                LogError(ex.Message, ex.StackTrace, "UnwindQuotesController", "UnwindQuotesList Get", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }
    }
}
