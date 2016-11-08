using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using SPPricing.Models;
using System.Data.Objects;
using Newtonsoft.Json;
using System.Net.Http;
using System.Data;
using System.Text.RegularExpressions;

namespace SPPricing.Controllers
{
    public class QuotronQuoteCreationController : Controller
    {
        //
        // GET: /QuotronQuoteCreation/

        SP_PRICINGEntities objSP_PRICINGEntities = new SP_PRICINGEntities();


        #region Request Create
        public ActionResult QuotronQuoteCreation(string ProductID, string IsQuotron)
        {
            LoginController objLoginController = new LoginController();
            try
            {
                if (ValidateSession())
                {
                    QuotronQuoteRequest objQuotronQuoteRequest = new QuotronQuoteRequest();

                    #region Menu Access By on Role

                    Int32 intResult = 0;
                    // bool PPorNonPP = false;

                    UserMaster objUserMaster = new UserMaster();
                    objUserMaster = (UserMaster)Session["LoggedInUser"];

                    var Result = objSP_PRICINGEntities.VALIDATE_MENU_BY_ROLE(objUserMaster.RoleID, "QQC");
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
                    objQuotronQuoteRequest.UnderlyingList = UnderlyingList;
                    #endregion

                    #region Line of Quotes List
                    List<QuotronQuoteRequest> LineofQuoteList = new List<QuotronQuoteRequest>();

                    ObjectResult<LineofQuoteResult> objLineofQuoteResult;
                    List<LineofQuoteResult> LineofQuoteResultList = new List<LineofQuoteResult>();

                    objLineofQuoteResult = objSP_PRICINGEntities.SP_FETCH_LINE_OF_QUOTES();
                    LineofQuoteResultList = objLineofQuoteResult.ToList();

                    foreach (LineofQuoteResult oLineofQuoteListResult in LineofQuoteResultList)
                    {
                        QuotronQuoteRequest objLOQ = new QuotronQuoteRequest();
                        General.ReflectSingleData(objLOQ, oLineofQuoteListResult);

                        LineofQuoteList.Add(objLOQ);
                    }
                    objQuotronQuoteRequest.LineOfQuotesList = LineofQuoteList;
                    #endregion

                    #region Product Type List
                    List<QuotronQuoteRequest> OptionTypeList = new List<QuotronQuoteRequest>();

                    ObjectResult<OptionTypeResult> objOptionTypeResult;
                    List<OptionTypeResult> OptionTypeResultList = new List<OptionTypeResult>();

                    objOptionTypeResult = objSP_PRICINGEntities.SP_FETCH_OPTION_TYPES();
                    OptionTypeResultList = objOptionTypeResult.ToList();

                    foreach (OptionTypeResult oOptionTypeListResult in OptionTypeResultList)
                    {
                        QuotronQuoteRequest objOptionType = new QuotronQuoteRequest();
                        General.ReflectSingleData(objOptionType, oOptionTypeListResult);

                        OptionTypeList.Add(objOptionType);
                    }
                    objQuotronQuoteRequest.OptionTypeList = OptionTypeList;
                    #endregion

                    if (ProductID != "" && ProductID != null)
                    {

                        ObjectResult<QuotronQuoteEditResult> objQuotronQuoteEditResult1 = objSP_PRICINGEntities.FETCH_QUOTRON_QUOTE_EDIT_DETAILS(ProductID);
                        List<QuotronQuoteEditResult> QuotronQuoteEditResultList1 = objQuotronQuoteEditResult1.ToList();

                        General.ReflectSingleData(objQuotronQuoteRequest, QuotronQuoteEditResultList1[0]);
                    }
                    else if (Session["CopyQuote"] != null)
                    {
                        ViewBag.Message = "Quote is copied";
                        objQuotronQuoteRequest = (QuotronQuoteRequest)Session["CopyQuote"];
                        objQuotronQuoteRequest.UnderlyingID = Convert.ToInt32(Session["UnderlyingID"]);
                        objQuotronQuoteRequest.LineOfQuotesId = Convert.ToInt32(Session["LineOfQuotesId"]);
                        objQuotronQuoteRequest.OptionTypeId = Convert.ToInt32(Session["OptionTypeId"]);


                        ObjectResult<QuotronQuoteEditResult> objQuotronQuoteEditResult1 = objSP_PRICINGEntities.FETCH_QUOTRON_QUOTE_EDIT_DETAILS("");
                        List<QuotronQuoteEditResult> QuotronQuoteEditResultList1 = objQuotronQuoteEditResult1.ToList();
                        QuotronQuoteRequest oQuotronQuoteRequest = new QuotronQuoteRequest();
                        if (QuotronQuoteEditResultList1 != null && QuotronQuoteEditResultList1.Count > 0)
                            General.ReflectSingleData(oQuotronQuoteRequest, QuotronQuoteEditResultList1[0]);

                        objQuotronQuoteRequest.ParentProductID = objQuotronQuoteRequest.ProductID;
                        objQuotronQuoteRequest.ProductID = "";
                        objQuotronQuoteRequest.Status = "";
                        objQuotronQuoteRequest.SaveStatus = oQuotronQuoteRequest.SaveStatus;
                        objQuotronQuoteRequest.IsCopyQuote = true;
                    }
                    else
                    {
                        Session.Remove("UnderlyingID");
                        Session.Remove("LineOfQuotesId");
                        Session.Remove("OptionTypeId");
                    }

                    if (Session["CopyQuote"] != null)
                        Session.Remove("CopyQuote");

                    return View(objQuotronQuoteRequest);
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
                LogError(ex.Message, ex.StackTrace, "QuotronQuoteCreationController", "QuotronQuoteCreation Get", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        [HttpPost]
        public ActionResult QuotronQuoteCreation(QuotronQuoteRequest objQuotronQuoteRequest, string Command)
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
                    objQuotronQuoteRequest.UnderlyingList = UnderlyingList;
                    #endregion

                    #region Line of Quotes List
                    List<QuotronQuoteRequest> LineofQuoteList = new List<QuotronQuoteRequest>();

                    ObjectResult<LineofQuoteResult> objLineofQuoteResult;
                    List<LineofQuoteResult> LineofQuoteResultList = new List<LineofQuoteResult>();

                    objLineofQuoteResult = objSP_PRICINGEntities.SP_FETCH_LINE_OF_QUOTES();
                    LineofQuoteResultList = objLineofQuoteResult.ToList();

                    foreach (LineofQuoteResult oLineofQuoteListResult in LineofQuoteResultList)
                    {
                        QuotronQuoteRequest objLOQ = new QuotronQuoteRequest();
                        General.ReflectSingleData(objLOQ, oLineofQuoteListResult);

                        LineofQuoteList.Add(objLOQ);
                    }
                    objQuotronQuoteRequest.LineOfQuotesList = LineofQuoteList;
                    #endregion

                    #region Product Type List
                    List<QuotronQuoteRequest> OptionTypeList = new List<QuotronQuoteRequest>();

                    ObjectResult<OptionTypeResult> objOptionTypeResult;
                    List<OptionTypeResult> OptionTypeResultList = new List<OptionTypeResult>();

                    objOptionTypeResult = objSP_PRICINGEntities.SP_FETCH_OPTION_TYPES();
                    OptionTypeResultList = objOptionTypeResult.ToList();

                    foreach (OptionTypeResult oOptionTypeListResult in OptionTypeResultList)
                    {
                        QuotronQuoteRequest objOptionType = new QuotronQuoteRequest();
                        General.ReflectSingleData(objOptionType, oOptionTypeListResult);

                        OptionTypeList.Add(objOptionType);
                    }
                    objQuotronQuoteRequest.OptionTypeList = OptionTypeList;
                    #endregion

                    if (Command == "CopyQuote")
                    {
                        Session["CopyQuote"] = objQuotronQuoteRequest;
                        Session["UnderlyingID"] = objQuotronQuoteRequest.UnderlyingID;
                        Session["LineOfQuotesId"] = objQuotronQuoteRequest.LineOfQuotesId;
                        Session["OptionTypeId"] = objQuotronQuoteRequest.OptionTypeId;

                        return RedirectToAction("QuotronQuoteCreation");
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
                LogError(ex.Message, ex.StackTrace, "QuotronQuoteCreationController", "QuotronQuoteCreation Get", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        public ActionResult ManageQuoteCreation(string ProductID, string Distributor, string LineOfQuotesId, string CapG, string UnderlyingID, string OptionTypeId, string OptionTenureMonth, string ProductTenure, string InitialAveragingDaysDiff, string InitialAveragingFrequency, string FinalAveragingDaysDiff, string FinalAveragingFrequency, string EdelweissBuiltIn, string DistributorBuiltIn, string FixedCoupon, string PaticipatoryRatio, string Strike, string BarrierLevel, string ObservationFrequency, string Rebate, string RangeAccrualLevel, string PayoutTime, string RequiredVariable, string RequiredPayoff, string SalesComments, string TradingComments)
        {
            LoginController objLoginController = new LoginController();
            try
            {
                if (ValidateSession())
                {
                    UserMaster objUserMaster = new UserMaster();
                    objUserMaster = (UserMaster)Session["LoggedInUser"];

                    ObjectResult<QuotronQuoteResult> objQuotronQuoteResult = objSP_PRICINGEntities.SP_MANAGE_QUOTRON_QUOTE(ProductID, Distributor, LineOfQuotesId, Convert.ToBoolean(CapG), UnderlyingID, OptionTypeId, OptionTenureMonth, ProductTenure, InitialAveragingDaysDiff, InitialAveragingFrequency, FinalAveragingDaysDiff, FinalAveragingFrequency, EdelweissBuiltIn, DistributorBuiltIn, FixedCoupon, PaticipatoryRatio, Strike, BarrierLevel, ObservationFrequency, Rebate, RangeAccrualLevel, PayoutTime, RequiredVariable, RequiredPayoff, SalesComments, TradingComments, objUserMaster.UserID, DateTime.Now);
                    List<QuotronQuoteResult> ManageQuotronQuoteResultList = objQuotronQuoteResult.ToList();

                    return Json(ManageQuotronQuoteResultList[0]);
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

                LogError(ex.Message, ex.StackTrace, "QuotronQuoteCreationController", "ManageFixedPlusPR Get", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }

        }

        public JsonResult ManagePricerStatusLog(string PricerType, string ProductID, string StatusCode)
        {
            try
            {
                Int32 intResult = 0;

                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];


                var Result = objSP_PRICINGEntities.SP_MANAGE_PRICER_STATUS_LOG(PricerType, ProductID, objUserMaster.UserID, StatusCode);
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

        public JsonResult FetchPricerStatus(string PricerType, string ProductID)
        {
            string strStatus = "";

            ObjectResult<QuotronQuoteEditResult> objQuotronQuoteEditResult = objSP_PRICINGEntities.FETCH_QUOTRON_QUOTE_EDIT_DETAILS(ProductID);
            List<QuotronQuoteEditResult> QuotronQuoteEditResultList = objQuotronQuoteEditResult.ToList();
            QuotronQuoteRequest oQuotronQuoteRequest = new QuotronQuoteRequest();
            General.ReflectSingleData(oQuotronQuoteRequest, QuotronQuoteEditResultList[0]);

            strStatus = oQuotronQuoteRequest.Status;

            return Json(strStatus);
        }

        #endregion

        #region Search Quote

        public ActionResult QuotronQuoteList(string Status)
        {
            LoginController objLoginController = new LoginController();
            try
            {
                if (ValidateSession())
                {
                    QuotronQuoteRequest objQuotronQuoteRequest = new QuotronQuoteRequest();

                    #region Menu Access By on Role

                    Int32 intResult = 0;
                    // bool PPorNonPP = false;

                    UserMaster objUserMaster = new UserMaster();
                    objUserMaster = (UserMaster)Session["LoggedInUser"];

                    var Result = objSP_PRICINGEntities.VALIDATE_MENU_BY_ROLE(objUserMaster.RoleID, "QQL");
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
                    objQuotronQuoteRequest.UnderlyingList = UnderlyingList;
                    #endregion

                    #region Line of Quotes List
                    List<QuotronQuoteRequest> LineofQuoteList = new List<QuotronQuoteRequest>();

                    ObjectResult<LineofQuoteResult> objLineofQuoteResult;
                    List<LineofQuoteResult> LineofQuoteResultList = new List<LineofQuoteResult>();

                    objLineofQuoteResult = objSP_PRICINGEntities.SP_FETCH_LINE_OF_QUOTES();
                    LineofQuoteResultList = objLineofQuoteResult.ToList();

                    foreach (LineofQuoteResult oLineofQuoteListResult in LineofQuoteResultList)
                    {
                        QuotronQuoteRequest objLOQ = new QuotronQuoteRequest();
                        General.ReflectSingleData(objLOQ, oLineofQuoteListResult);

                        LineofQuoteList.Add(objLOQ);
                    }
                    objQuotronQuoteRequest.LineOfQuotesList = LineofQuoteList;
                    #endregion

                    #region Product Type List
                    List<QuotronQuoteRequest> OptionTypeList = new List<QuotronQuoteRequest>();

                    ObjectResult<OptionTypeResult> objOptionTypeResult;
                    List<OptionTypeResult> OptionTypeResultList = new List<OptionTypeResult>();

                    objOptionTypeResult = objSP_PRICINGEntities.SP_FETCH_OPTION_TYPES();
                    OptionTypeResultList = objOptionTypeResult.ToList();

                    foreach (OptionTypeResult oOptionTypeListResult in OptionTypeResultList)
                    {
                        QuotronQuoteRequest objOptionType = new QuotronQuoteRequest();
                        General.ReflectSingleData(objOptionType, oOptionTypeListResult);

                        OptionTypeList.Add(objOptionType);
                    }
                    objQuotronQuoteRequest.OptionTypeList = OptionTypeList;
                    #endregion

                    #region Status List
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

                    objQuotronQuoteRequest.StatusList = StatusList;

                    //--Set Status--Added by Shweta on 27th May 2016------------START--------------------
                    if (Status != null && Status != "")
                    {
                        LookupMaster objStatus = objQuotronQuoteRequest.StatusList.Find(delegate(LookupMaster oLookupMaster) { return oLookupMaster.LookupID == Convert.ToInt32(Status); });
                        objQuotronQuoteRequest.FilterStatus = Convert.ToString(objStatus.LookupID);
                    }
                    //--Set Status--Added by Shweta on 27th May 2016------------END----------------------
                    #endregion

                    #region Quote Type List
                    List<LookupMaster> QuoteTypeList = new List<LookupMaster>();
                    objLookupResult = objSP_PRICINGEntities.SP_FETCH_LOOKUP_VALUES("QSTM", false);
                    LookupResultList = objLookupResult.ToList();

                    if (LookupResultList != null && LookupResultList.Count > 0)
                    {
                        foreach (var LookupResult in LookupResultList)
                        {
                            LookupMaster objLookupMaster = new LookupMaster();
                            General.ReflectSingleData(objLookupMaster, LookupResult);

                            QuoteTypeList.Add(objLookupMaster);
                        }
                    }

                    objQuotronQuoteRequest.QuoteTypeList = QuoteTypeList;
                    #endregion

                    return View(objQuotronQuoteRequest);
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
                LogError(ex.Message, ex.StackTrace, "QuotronQuoteCreationController", "QuotronQuoteList Get", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        public ActionResult RedirectToMethod(string ProductID)
        {
            LoginController objLoginController = new LoginController();
            try
            {
                if (ValidateSession())
                {
                    if (ProductID != "" && ProductID != null)
                    {

                        var PricerCode = Regex.Matches(ProductID, @"\D+|\d+")
                                 .Cast<Match>()
                                 .Select(m => m.Value)
                                 .ToArray();

                        ObjectResult<PricerCodeMasterList> objQuotronQuoteEditResult = objSP_PRICINGEntities.SP_FETCH_PRICER_MASTER();
                        List<PricerCodeMasterList> QuotronQuoteEditResultList = objQuotronQuoteEditResult.ToList();
                        List<PricerCodeMaster> test = new List<PricerCodeMaster>();

                        foreach (var a in QuotronQuoteEditResultList)
                        {
                            PricerCodeMaster obj = new PricerCodeMaster();
                            obj.PricerCode = a.PricerCode;
                            obj.ControllerName = a.ControllerName;
                            obj.MethodName = a.MethodName;

                            test.Add(obj);
                        }

                        PricerCodeMaster objPricerCodeMaster;

                        objPricerCodeMaster = test.Find(delegate(PricerCodeMaster oPricerCodeMaster) { return oPricerCodeMaster.PricerCode == "#" + PricerCode[0]; });
                        if (objPricerCodeMaster != null)
                        {
                            ProductID = "#" + ProductID;
                        }
                        else
                        {
                            objPricerCodeMaster = test.Find(delegate(PricerCodeMaster oPricerCodeMaster) { return oPricerCodeMaster.PricerCode == PricerCode[0]; });
                        }


                        return RedirectToAction(objPricerCodeMaster.MethodName, objPricerCodeMaster.ControllerName, new { ProductID = ProductID, IsQuotron = true });
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
                LogError(ex.Message, ex.StackTrace, "QuotronQuoteCreationController", "QuotronQuoteList Get", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        [AcceptVerbs(HttpVerbs.Get)]
        public ActionResult FetchLineOfQuoteByOptionType(int OptionTypeID)
        {
            //TO DO : get data from wherever you want. 
            LoginController objLoginController = new LoginController();

            List<KeyValuePair<int, string>> OptionTypeList = new List<KeyValuePair<int, string>>();

            try
            {
                if (ValidateSession())
                {
                    QuotronQuoteRequest objQuotronQuoteRequest = new QuotronQuoteRequest();

                    DataSet dsResult = new DataSet();
                    dsResult = General.ExecuteDataSet("FETCH_LINE_OF_QUOTE_BY_OPTION_TYPE", OptionTypeID);

                    List<LineOfQuoteModel> LineOfQuoteList = new List<LineOfQuoteModel>();

                    if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                    {
                        foreach (DataRow dr in dsResult.Tables[0].Rows)
                        {
                            LineOfQuoteModel objLineOfQuote = new LineOfQuoteModel();
                            //objCustomerDetails.ID = Convert.ToInt32(dr["ID"]);
                            objLineOfQuote.LineOfQuotes = Convert.ToString(dr["LineOfQuotes"]);
                            objLineOfQuote.ID = Convert.ToInt32(dr["ID"]);


                            LineOfQuoteList.Add(objLineOfQuote);
                        }
                    }
                    return Json(LineOfQuoteList, JsonRequestBehavior.AllowGet);
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
                LogError(ex.Message, ex.StackTrace, "QuotronQuoteCreationController", "QuotronQuoteList Get", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }




        }

        public JsonResult FetchQuotronQuoteList(string ProductID, string QuoteType, string Distributor, string LineOfQuotesName, string CapG, string UnderlyingName, string OptionTypeName, string OptionTenureMonth, string ProductTenure, string EdelweissBuiltIn, string DistributorBuiltIn, string FixedCoupon, string PaticipatoryRatio, string Strike, string BarrierLevel, string ObservationFrequency, string RangeAccrualLevel, string RequiredPayoff, string Status, string SalesComments, string TradingComments)
        {
            try
            {
                List<QuotronQuoteRequest> QuotronQuoteRequestList = new List<QuotronQuoteRequest>();

                if (ProductID == "" || ProductID == "--Select--")
                    ProductID = "ALL";

                if (QuoteType == "" || QuoteType == "--Select--" || QuoteType == "-1")
                    QuoteType = "ALL";
                
                if (Distributor == "" || Distributor == "0" || Distributor == "--Select--")
                    Distributor = "ALL";

                if (LineOfQuotesName == "" || LineOfQuotesName == "0" || LineOfQuotesName == "--Select--" || LineOfQuotesName == "-1")
                    LineOfQuotesName = "ALL";

                if (CapG == "--Select--")
                    CapG = "ALL";

                if (UnderlyingName == "" || UnderlyingName == "0" || UnderlyingName == "--Select--")
                    UnderlyingName = "ALL";

                if (OptionTypeName == "" || OptionTypeName == "0" || OptionTypeName == "--Select--")
                    OptionTypeName = "ALL";

                if (OptionTenureMonth == "" || OptionTenureMonth == "0" || OptionTenureMonth == "--Select--")
                    OptionTenureMonth = "ALL";

                if (ProductTenure == "" || ProductTenure == "0" || ProductTenure == "--Select--")
                    ProductTenure = "ALL";


                if (EdelweissBuiltIn == "" || EdelweissBuiltIn == "0" || EdelweissBuiltIn == "--Select--")
                    EdelweissBuiltIn = "ALL";

                if (DistributorBuiltIn == "" || DistributorBuiltIn == "0" || DistributorBuiltIn == "--Select--")
                    DistributorBuiltIn = "ALL";

                if (FixedCoupon == "" || FixedCoupon == "0" || FixedCoupon == "--Select--")
                    FixedCoupon = "ALL";

                if (PaticipatoryRatio == "" || PaticipatoryRatio == "0" || PaticipatoryRatio == "--Select--")
                    PaticipatoryRatio = "ALL";

                if (Strike == "" || Strike == "0" || Strike == "--Select--")
                    Strike = "ALL";

                if (BarrierLevel == "" || BarrierLevel == "0" || BarrierLevel == "--Select--")
                    BarrierLevel = "ALL";

                if (RequiredPayoff == "" || RequiredPayoff == "0" || RequiredPayoff == "--Select--")
                    RequiredPayoff = "ALL";

                if (RangeAccrualLevel == "" || RangeAccrualLevel == "0" || RangeAccrualLevel == "--Select--")
                    RangeAccrualLevel = "ALL";

                if (ObservationFrequency == "" || ObservationFrequency == "0" || ObservationFrequency == "--Select--")
                    ObservationFrequency = "ALL";

                if (Status == "" || Status == "0" || Status == "--Select--")
                    Status = "ALL";

                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("SP_FETCH_QUOTRON_QUOTE", ProductID, QuoteType, LineOfQuotesName, OptionTypeName, UnderlyingName, FixedCoupon, CapG, Distributor, RequiredPayoff, BarrierLevel, PaticipatoryRatio, ObservationFrequency, Strike, OptionTenureMonth, ProductTenure, EdelweissBuiltIn, DistributorBuiltIn, RangeAccrualLevel, Status, SalesComments, TradingComments);

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in dsResult.Tables[0].Rows)
                    {
                        QuotronQuoteRequest obj = new QuotronQuoteRequest();

                        obj.ProductID = Convert.ToString(dr["ProductID"]);
                        obj.RequiredVariable = Convert.ToString(dr["RequiredVariable"]);
                        obj.LineOfQuotesName = Convert.ToString(dr["LineOfQuotesName"]);
                        obj.OptionTypeName = Convert.ToString(dr["OptionTypeName"]);
                        obj.CapG = Convert.ToBoolean(dr["CapG"]);
                        obj.UnderlyingName = Convert.ToString(dr["UnderlyingName"]);
                        obj.OptionTenureMonth = Convert.ToString(dr["OptionTenureMonth"]);
                        obj.ProductTenure = Convert.ToString(dr["ProductTenure"]);
                        obj.InitialAveragingDaysDiff = Convert.ToString(dr["InitialAveragingDaysDiff"]);
                        obj.InitialAveragingFrequency = Convert.ToString(dr["InitialAveragingFrequency"]);
                        obj.FinalAveragingDaysDiff = Convert.ToString(dr["FinalAveragingDaysDiff"]);
                        obj.FinalAveragingFrequency = Convert.ToString(dr["FinalAveragingFrequency"]);
                        obj.EdelweissBuiltIn = Convert.ToString(dr["EdelweissBuiltIn"]);
                        obj.DistributorBuiltIn = Convert.ToString(dr["DistributorBuiltIn"]);
                        obj.FixedCoupon = Convert.ToString(dr["FixedCoupon"]);
                        obj.Strike = Convert.ToString(dr["Strike"]);
                        obj.PaticipatoryRatio = Convert.ToString(dr["PaticipatoryRatio"]);
                        obj.BarrierLevel = Convert.ToString(dr["BarrierLevel"]);
                        obj.Rebate = Convert.ToString(dr["Rebate"]);
                        obj.ObservationFrequency = Convert.ToString(dr["ObservationFrequency"]);
                        obj.RangeAccrualLevel = Convert.ToString(dr["RangeAccrualLevel"]);
                        obj.PayoutTime = Convert.ToString(dr["PayoutTime"]);
                        obj.RequiredPayoff = Convert.ToString(dr["RequiredPayoff"]);
                        obj.SalesComments = Convert.ToString(dr["SalesComments"]);
                        obj.TradingComments = Convert.ToString(dr["TradingComments"]);
                        obj.IsFavourite = Convert.ToBoolean(dr["IsFavourite"]);
                        obj.Distributor = Convert.ToString(dr["Distributor"]);
                        obj.Status = Convert.ToString(dr["Status"]);
                        obj.ConfirmedOn = Convert.ToString(dr["ConfirmedOn"]);

                        QuotronQuoteRequestList.Add(obj);
                    }
                }

                var QuotronQuoteRequestListData = QuotronQuoteRequestList.ToList();
                return Json(QuotronQuoteRequestListData, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "QuotronQuoteCreationController", "FetchQuotronQuoteList", objUserMaster.UserID);
                return Json("");
            }
        }

        public ActionResult AutoCompleteProductID(string term)
        {
            try
            {
                List<QuotronQuoteRequest> QuotronQuoteRequestList = new List<QuotronQuoteRequest>();

                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("SP_FETCH_QUOTRON_QUOTE", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL");

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in dsResult.Tables[0].Rows)
                    {
                        QuotronQuoteRequest obj = new QuotronQuoteRequest();

                        obj.ProductID = Convert.ToString(dr["ProductID"]);
                        obj.RequiredVariable = Convert.ToString(dr["RequiredVariable"]);
                        obj.LineOfQuotesName = Convert.ToString(dr["LineOfQuotesName"]);
                        obj.OptionTypeName = Convert.ToString(dr["OptionTypeName"]);
                        obj.CapG = Convert.ToBoolean(dr["CapG"]);
                        obj.UnderlyingName = Convert.ToString(dr["UnderlyingName"]);
                        obj.OptionTenureMonth = Convert.ToString(dr["OptionTenureMonth"]);
                        obj.ProductTenure = Convert.ToString(dr["ProductTenure"]);
                        obj.InitialAveragingDaysDiff = Convert.ToString(dr["InitialAveragingDaysDiff"]);
                        obj.InitialAveragingFrequency = Convert.ToString(dr["InitialAveragingFrequency"]);
                        obj.FinalAveragingDaysDiff = Convert.ToString(dr["FinalAveragingDaysDiff"]);
                        obj.FinalAveragingFrequency = Convert.ToString(dr["FinalAveragingFrequency"]);
                        obj.EdelweissBuiltIn = Convert.ToString(dr["EdelweissBuiltIn"]);
                        obj.DistributorBuiltIn = Convert.ToString(dr["DistributorBuiltIn"]);
                        obj.FixedCoupon = Convert.ToString(dr["FixedCoupon"]);
                        obj.Strike = Convert.ToString(dr["Strike"]);
                        obj.PaticipatoryRatio = Convert.ToString(dr["PaticipatoryRatio"]);
                        obj.BarrierLevel = Convert.ToString(dr["BarrierLevel"]);
                        obj.Rebate = Convert.ToString(dr["Rebate"]);
                        obj.ObservationFrequency = Convert.ToString(dr["ObservationFrequency"]);
                        obj.RangeAccrualLevel = Convert.ToString(dr["RangeAccrualLevel"]);
                        obj.PayoutTime = Convert.ToString(dr["PayoutTime"]);
                        obj.RequiredPayoff = Convert.ToString(dr["RequiredPayoff"]);
                        obj.SalesComments = Convert.ToString(dr["SalesComments"]);
                        obj.TradingComments = Convert.ToString(dr["TradingComments"]);
                        obj.Distributor = Convert.ToString(dr["Distributor"]);
                        obj.Status = Convert.ToString(dr["Status"]);
                        obj.ConfirmedOn = Convert.ToString(dr["ConfirmedOn"]);


                        QuotronQuoteRequestList.Add(obj);
                    }
                }

                var DistinctItems = QuotronQuoteRequestList.GroupBy(x => x.ProductID).Select(y => y.First());

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

        public ActionResult AutoCompleteUnderlyingID(string term)
        {
            try
            {
                List<QuotronQuoteRequest> QuotronQuoteRequestList = new List<QuotronQuoteRequest>();

                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("SP_FETCH_QUOTRON_QUOTE", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL");

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in dsResult.Tables[0].Rows)
                    {
                        QuotronQuoteRequest obj = new QuotronQuoteRequest();

                        obj.ProductID = Convert.ToString(dr["ProductID"]);
                        obj.RequiredVariable = Convert.ToString(dr["RequiredVariable"]);
                        obj.LineOfQuotesName = Convert.ToString(dr["LineOfQuotesName"]);
                        obj.OptionTypeName = Convert.ToString(dr["OptionTypeName"]);
                        obj.CapG = Convert.ToBoolean(dr["CapG"]);
                        obj.UnderlyingName = Convert.ToString(dr["UnderlyingName"]);
                        obj.OptionTenureMonth = Convert.ToString(dr["OptionTenureMonth"]);
                        obj.ProductTenure = Convert.ToString(dr["ProductTenure"]);
                        obj.InitialAveragingDaysDiff = Convert.ToString(dr["InitialAveragingDaysDiff"]);
                        obj.InitialAveragingFrequency = Convert.ToString(dr["InitialAveragingFrequency"]);
                        obj.FinalAveragingDaysDiff = Convert.ToString(dr["FinalAveragingDaysDiff"]);
                        obj.FinalAveragingFrequency = Convert.ToString(dr["FinalAveragingFrequency"]);
                        obj.EdelweissBuiltIn = Convert.ToString(dr["EdelweissBuiltIn"]);
                        obj.DistributorBuiltIn = Convert.ToString(dr["DistributorBuiltIn"]);
                        obj.FixedCoupon = Convert.ToString(dr["FixedCoupon"]);
                        obj.Strike = Convert.ToString(dr["Strike"]);
                        obj.PaticipatoryRatio = Convert.ToString(dr["PaticipatoryRatio"]);
                        obj.BarrierLevel = Convert.ToString(dr["BarrierLevel"]);
                        obj.Rebate = Convert.ToString(dr["Rebate"]);
                        obj.ObservationFrequency = Convert.ToString(dr["ObservationFrequency"]);
                        obj.RangeAccrualLevel = Convert.ToString(dr["RangeAccrualLevel"]);
                        obj.PayoutTime = Convert.ToString(dr["PayoutTime"]);
                        obj.RequiredPayoff = Convert.ToString(dr["RequiredPayoff"]);
                        obj.SalesComments = Convert.ToString(dr["SalesComments"]);
                        obj.TradingComments = Convert.ToString(dr["TradingComments"]);
                        obj.Distributor = Convert.ToString(dr["Distributor"]);
                        obj.Status = Convert.ToString(dr["Status"]);
                        obj.ConfirmedOn = Convert.ToString(dr["ConfirmedOn"]);


                        QuotronQuoteRequestList.Add(obj);
                    }
                }

                var DistinctItems = QuotronQuoteRequestList.GroupBy(x => x.UnderlyingName).Select(y => y.First());

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

        public ActionResult AutoCompleteDistributor(string term)
        {
            try
            {
                List<QuotronQuoteRequest> QuotronQuoteRequestList = new List<QuotronQuoteRequest>();

                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("SP_FETCH_QUOTRON_QUOTE", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL", "ALL");

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in dsResult.Tables[0].Rows)
                    {
                        QuotronQuoteRequest obj = new QuotronQuoteRequest();

                        obj.ProductID = Convert.ToString(dr["ProductID"]);
                        obj.RequiredVariable = Convert.ToString(dr["RequiredVariable"]);
                        obj.LineOfQuotesName = Convert.ToString(dr["LineOfQuotesName"]);
                        obj.OptionTypeName = Convert.ToString(dr["OptionTypeName"]);
                        obj.CapG = Convert.ToBoolean(dr["CapG"]);
                        obj.UnderlyingName = Convert.ToString(dr["UnderlyingName"]);
                        obj.OptionTenureMonth = Convert.ToString(dr["OptionTenureMonth"]);
                        obj.ProductTenure = Convert.ToString(dr["ProductTenure"]);
                        obj.InitialAveragingDaysDiff = Convert.ToString(dr["InitialAveragingDaysDiff"]);
                        obj.InitialAveragingFrequency = Convert.ToString(dr["InitialAveragingFrequency"]);
                        obj.FinalAveragingDaysDiff = Convert.ToString(dr["FinalAveragingDaysDiff"]);
                        obj.FinalAveragingFrequency = Convert.ToString(dr["FinalAveragingFrequency"]);
                        obj.EdelweissBuiltIn = Convert.ToString(dr["EdelweissBuiltIn"]);
                        obj.DistributorBuiltIn = Convert.ToString(dr["DistributorBuiltIn"]);
                        obj.FixedCoupon = Convert.ToString(dr["FixedCoupon"]);
                        obj.Strike = Convert.ToString(dr["Strike"]);
                        obj.PaticipatoryRatio = Convert.ToString(dr["PaticipatoryRatio"]);
                        obj.BarrierLevel = Convert.ToString(dr["BarrierLevel"]);
                        obj.Rebate = Convert.ToString(dr["Rebate"]);
                        obj.ObservationFrequency = Convert.ToString(dr["ObservationFrequency"]);
                        obj.RangeAccrualLevel = Convert.ToString(dr["RangeAccrualLevel"]);
                        obj.PayoutTime = Convert.ToString(dr["PayoutTime"]);
                        obj.RequiredPayoff = Convert.ToString(dr["RequiredPayoff"]);
                        obj.SalesComments = Convert.ToString(dr["SalesComments"]);
                        obj.TradingComments = Convert.ToString(dr["TradingComments"]);
                        obj.Distributor = Convert.ToString(dr["Distributor"]);
                        obj.Status = Convert.ToString(dr["Status"]);
                        obj.ConfirmedOn = Convert.ToString(dr["ConfirmedOn"]);


                        QuotronQuoteRequestList.Add(obj);
                    }
                }

                var DistinctItems = QuotronQuoteRequestList.GroupBy(x => x.Distributor).Select(y => y.First());

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


        #endregion

        public void LogError(string strErrorDescription, string strStackTrace, string strClassName, string strMethodName, Int32 intUserId)
        {
            SP_PRICINGEntities objSP_PRICINGEntities = new SP_PRICINGEntities();
            var Count = objSP_PRICINGEntities.SP_ERROR_LOG(strErrorDescription, strStackTrace, strClassName, strMethodName, intUserId);
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
                objLoginController.LogError(ex.Message, ex.StackTrace, "QuotronQuoteCreationController", "ValidateSession", -1);
                return false;
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
                LogError(ex.Message, ex.StackTrace, "QuotronQuoteCreationController", "FetchUnderlyingList", objUserMaster.UserID);
                return Json("");
            }

            //return Json(UnderlyingListData);
        }
    }
}