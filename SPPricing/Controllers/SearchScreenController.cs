using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using SPPricing.Models;

namespace SPPricing.Controllers
{
    public class SearchScreenController : Controller
    {
        //
        // GET: /SearchScreen/

        public ActionResult Index()
        {
            return View();
        }

        public ActionResult SearchScreen()
        {
            try
            {
                if (ValidateSession())
                {
                    SearchCriteria objSearchCriteria = new SearchCriteria();

                    objSearchCriteria.QuoteTypeList = new List<LookupMaster>();
                    objSearchCriteria.StatusList = new List<LookupMaster>();

                    return View(objSearchCriteria);
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
                LogError(ex.Message, ex.StackTrace, "SearchScreenController", "SearchScreen Post", objUserMaster.UserID);
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
                objLoginController.LogError(ex.Message, ex.StackTrace, "SearchScreenController", "ValidateSession", -1);
                return false;
            }
        }
    }
}