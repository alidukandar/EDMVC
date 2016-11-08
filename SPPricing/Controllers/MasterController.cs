using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using SPPricing.Models;
using System.Data.Objects;

namespace SPPricing.Controllers
{
    public class MasterController : Controller
    {
        //
        // GET: /Master/

        SP_PRICINGEntities objSP_PRICINGEntities = new SP_PRICINGEntities();

        #region Product Type
        public ActionResult ProductTypeMaster()
        {
            LoginController objLoginController = new LoginController();
            try
            {
                if (ValidateSession())
                {
                    #region Menu Access By on Role

                    Int32 intResult = 0;
                    // bool PPorNonPP = false;

                    UserMaster objUserMaster = new UserMaster();
                    objUserMaster = (UserMaster)Session["LoggedInUser"];

                    var Result = objSP_PRICINGEntities.VALIDATE_MENU_BY_ROLE(objUserMaster.RoleID, "PTM");
                    intResult = Convert.ToInt32(Result.SingleOrDefault());

                    if (intResult == 0)
                        return RedirectToAction("UserNotAuthorize", "Login");

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
                LogError(ex.Message, ex.StackTrace, "MasterController", "ProductTypeMaster Get", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        public JsonResult AddProductType(string ProductType, string ProductCode)
        {
            try
            {
                Int32 intResult = 0;
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];

                var Result = objSP_PRICINGEntities.SP_MANAGE_PRODUCT_TYPES(ProductType, ProductCode,false, objUserMaster.UserID);
                intResult = Convert.ToInt32(Result.SingleOrDefault());

                return Json(intResult);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "MasterController", "AddProductType", objUserMaster.UserID);
                return Json("");
            }
        }

        public JsonResult DeleteProductType(string ProductType)
        {
            try
            {
                Int32 intResult = 0;
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];

                var Result = objSP_PRICINGEntities.SP_MANAGE_PRODUCT_TYPES(ProductType, "", true, objUserMaster.UserID);
                intResult = Convert.ToInt32(Result.SingleOrDefault());

                if (intResult == 2)
                    ViewBag.Message = "Data Deleted Successfully";
                return Json(intResult);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "MasterController", "DeleteProductType", objUserMaster.UserID);
                return Json("");
            }
        }


        public JsonResult FetchProductTypeList()
        {
            try
            {               

                List<Underlying> UnderlyingList = new List<Underlying>();

                ObjectResult<ProductTypeListResult> objProductTypeListResult;
                //List<UnderlyingListResult> UnderlyingListResultList;

                objProductTypeListResult = objSP_PRICINGEntities.SP_FETCH_PRODUCT_TYPES();


                var ProductTypeListData = objProductTypeListResult.ToList();
                return Json(ProductTypeListData, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "MasterController", "FetchProductTypeList", objUserMaster.UserID);
                return Json("");
            }
        }
        #endregion

        #region Line of Quote

        public ActionResult LineOfQuoteMaster()
        {
            LoginController objLoginController = new LoginController();
            try
            {
                if (ValidateSession())
                {
                    LineOfQuoteModel objLineOfQuoteModel = new LineOfQuoteModel();

                    #region Menu Access By on Role

                    Int32 intResult = 0;
                    // bool PPorNonPP = false;

                    UserMaster objUserMaster = new UserMaster();
                    objUserMaster = (UserMaster)Session["LoggedInUser"];

                    var Result = objSP_PRICINGEntities.VALIDATE_MENU_BY_ROLE(objUserMaster.RoleID, "LOQM");
                    intResult = Convert.ToInt32(Result.SingleOrDefault());

                    if (intResult == 0)
                        return RedirectToAction("UserNotAuthorize", "Login");

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
                    objLineOfQuoteModel.OptionTypeList = OptionTypeList;
                    #endregion
                    return View(objLineOfQuoteModel);
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
                LogError(ex.Message, ex.StackTrace, "MasterController", "LineOfQuoteMaster Get", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        public JsonResult FetchLineOfQuoteList()
        {
            try
            {

                ObjectResult<LineofQuoteResult> objLineofQuoteResult;
                List<LineofQuoteResult> LineofQuoteResultList = new List<LineofQuoteResult>();

                objLineofQuoteResult = objSP_PRICINGEntities.SP_FETCH_LINE_OF_QUOTES();
                LineofQuoteResultList = objLineofQuoteResult.ToList();

                return Json(LineofQuoteResultList, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "MasterController", "FetchLineOfQuoteList", objUserMaster.UserID);
                return Json("");
            }
        }

        public JsonResult AddLineOfQuote(string LineofQuote, string OptionType)
        {
            try
            {
                Int32 intResult = 0;
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];

                var Result = objSP_PRICINGEntities.SP_MANAGE_LINE_OF_QUOTES(LineofQuote,OptionType, false, objUserMaster.UserID);
                intResult = Convert.ToInt32(Result.SingleOrDefault());

                return Json(intResult);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "MasterController", "AddLineOfQuote", objUserMaster.UserID);
                return Json("");
            }
        }

        public JsonResult DeleteLineOfQuote(string LineofQuote, string OptionType)
        {
            try
            {
                Int32 intResult = 0;
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];

                var Result = objSP_PRICINGEntities.SP_MANAGE_LINE_OF_QUOTES(LineofQuote,OptionType, true, objUserMaster.UserID);
                intResult = Convert.ToInt32(Result.SingleOrDefault());

                if (intResult == 2)
                    ViewBag.Message = "Data Deleted Successfully";
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
                objLoginController.LogError(ex.Message, ex.StackTrace, "MasterController", "ValidateSession", -1);
                return false;
            }
        }
    }
}