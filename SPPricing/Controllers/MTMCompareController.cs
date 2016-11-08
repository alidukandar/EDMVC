using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using SPPricing.Models;
using System.Data;
using System.Data.Objects;

namespace SPPricing.Controllers
{
    public class MTMCompareController : Controller
    {
        SP_PRICINGEntities objSP_PRICINGEntities = new SP_PRICINGEntities();
        //
        // GET: /MTMCompare/

        public ActionResult Index()
        {
            return View();
        }

        [HttpGet]
        public ActionResult MTMCompare()
        {
            try
            {
                if (ValidateSession())
                {
                    List<MTMCompare> MTMCompareList = new List<MTMCompare>();

                    #region Menu Access By on Role

                    Int32 intResult = 0;
                    // bool PPorNonPP = false;

                    UserMaster objUserMaster = new UserMaster();
                    objUserMaster = (UserMaster)Session["LoggedInUser"];

                    var Result = objSP_PRICINGEntities.VALIDATE_MENU_BY_ROLE(objUserMaster.RoleID, "MTMC");
                    intResult = Convert.ToInt32(Result.SingleOrDefault());

                    if (intResult == 0)
                        return RedirectToAction("UserNotAuthorize", "Login");

                    #endregion

                    string OldReportDate = "1900-01-01";

                    string NewReportDate = "2900-01-01";
                    ObjectResult<TopDateResult> objTopDateResult = objSP_PRICINGEntities.SP_FETCH_TOP2_DATE(Convert.ToDateTime(OldReportDate), Convert.ToDateTime(NewReportDate));
                    List<TopDateResult> TopDateResultList = objTopDateResult.ToList();

                    MTMCompare obj = new MTMCompare();

                    obj.OldReportDate = Convert.ToDateTime(TopDateResultList[0].OldDate);
                    obj.NewReportDate = Convert.ToDateTime(TopDateResultList[0].NewDate);
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
                LogError(ex.Message, ex.StackTrace, "MTMCompareController", "MTMCompare Get", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        [HttpGet]
        public ActionResult MTMCompare2(string ProductCode, string AUM, string OldReportDate, string NewReportDate)
        {
            try
            {
                if (ValidateSession())
                {
                    MTMCompareScreen2 obj = new MTMCompareScreen2();
                    obj.AUM = Convert.ToDouble(AUM);
                    obj.OldReportDate = Convert.ToDateTime(OldReportDate);
                    obj.NewReportDate = Convert.ToDateTime(NewReportDate);
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
                LogError(ex.Message, ex.StackTrace, "MTMCompareController", "MTMCompare2 Get", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        public JsonResult FetchMTMCompareList(string ProductCode, string OldReportDate, string NewReportDate)
        {
            try
            {
                List<MTMCompare> MTMCompareList = new List<MTMCompare>();

                if (OldReportDate == "")
                    OldReportDate = "1900-01-01";

                if (NewReportDate == "")
                    NewReportDate = "2900-01-01";

                if (ProductCode == "")
                    ProductCode = "ALL";              

                ObjectResult<MTMCompareScreen1Result> objMTMCompareScreen1Result = objSP_PRICINGEntities.SP_FETCH_MTM_COMPARE_SCREEN_1(ProductCode,Convert.ToDateTime(OldReportDate), Convert.ToDateTime(NewReportDate));
                List<MTMCompareScreen1Result> MTMCompareScreen1ResultList = objMTMCompareScreen1Result.ToList();

                return Json(MTMCompareScreen1ResultList, JsonRequestBehavior.AllowGet);

            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"]; 
                LogError(ex.Message, ex.StackTrace, "MTMCompareController", "FetchMTMCompareList", objUserMaster.UserID);
                return Json("");
            }
        }

        public JsonResult FetchMTMCompare2List(string ProductCode, string OldReportDate, string NewReportDate)
        {
            try
            {
                List<MTMCompareScreen2> MTMCompareScreen2List = new List<MTMCompareScreen2>();

                ObjectResult<MTMCompareScreen2Result> objMTMCompareScreen2Result = objSP_PRICINGEntities.SP_FETCH_MTM_COMPARE_SCREEN_2(Convert.ToDateTime(OldReportDate), Convert.ToDateTime(NewReportDate), ProductCode);
                List<MTMCompareScreen2Result> MTMCompareScreen2ResultList = objMTMCompareScreen2Result.ToList();
                var a = MTMCompareScreen2ResultList.Count();

                return Json(MTMCompareScreen2ResultList, JsonRequestBehavior.AllowGet);

            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "MTMCompareController", "FetchMTMCompare2List", objUserMaster.UserID);
                return Json("");
            }
        }

        public ActionResult AutoCompleteProductID(string term)
        {
            try
            {
                List<MTMCompare> MTMCompareList = new List<MTMCompare>();

                ObjectResult<MTMCompareScreen1Result> objMTMCompareScreen1Result = objSP_PRICINGEntities.SP_FETCH_MTM_COMPARE_SCREEN_1("ALL", Convert.ToDateTime("1900-01-01"), Convert.ToDateTime("2900-01-01"));
                List<MTMCompareScreen1Result> MTMCompareScreen1ResultList = objMTMCompareScreen1Result.ToList();

                var DistinctItems = MTMCompareScreen1ResultList.GroupBy(x => x.ProductCode).Select(y => y.First());

                var result = (from objRuleList in DistinctItems
                              where objRuleList.ProductCode.ToLower().StartsWith(term.ToLower())
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

        public JsonResult FetchBondMTM2CompareList(string ProductCode, string OldReportDate, string NewReportDate)
        {
            try
            {

                List<MTMCompareScreen2> MTMCompareScreen2List = new List<MTMCompareScreen2>();

                ObjectResult<BondMTMCompareScreen2Result> objMTMCompareScreen2Result = objSP_PRICINGEntities.SP_FETCH_BOND_MTM_COMPARE_SCREEN_2(Convert.ToDateTime(OldReportDate), Convert.ToDateTime(NewReportDate), ProductCode);
                List<BondMTMCompareScreen2Result> MTMCompareScreen2ResultList = objMTMCompareScreen2Result.ToList();

                return Json(MTMCompareScreen2ResultList, JsonRequestBehavior.AllowGet);

            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"]; 
                LogError(ex.Message, ex.StackTrace, "MTMCompareController", "FetchBondMTM2CompareList", objUserMaster.UserID);
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
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"]; 
                objLoginController.LogError(ex.Message, ex.StackTrace, "MTMCompareController", "ValidateSession", -1);
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
