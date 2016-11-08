using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.DirectoryServices;
using System.Data.Objects;
using SPPricing.Models;
using System.Web.UI;

namespace SPPricing.Controllers
{
    public class LoginController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        [HttpGet]
        public ActionResult Login(string ProductID)
        {
            if (ProductID != null && ProductID != "")
            {
                string strUserName = Request.ServerVariables["LOGON_USER"].ToUpper().Replace("EDELCAP" + "\\", "");
                string strData = windowAuthenticationSuccess(strUserName);

                if (strData != "")
                {
                    if (AuthenticateMailUser(strUserName, ""))
                    {
                        if (ProductID.Contains("FCM"))
                            return RedirectToAction("FixedCouponMLD", "BlackscholesPricers", new { ProductID = ProductID });
                        if (ProductID.Contains("FC"))
                            return RedirectToAction("FixedCoupon", "BlackscholesPricers", new { ProductID = ProductID });
                        if (ProductID.Contains("FPP"))
                            return RedirectToAction("FixedPlusPR", "BlackscholesPricers", new { ProductID = ProductID });
                        if (ProductID.Contains("FOP"))
                            return RedirectToAction("FixedOrPR", "BlackscholesPricers", new { ProductID = ProductID });
                        if (ProductID.Contains("GC"))
                            return RedirectToAction("GoldenCushion", "BlackscholesPricers", new { ProductID = ProductID });
                        if (ProductID.Contains("CB"))
                            return RedirectToAction("CallBinary", "BlackscholesPricers", new { ProductID = ProductID });
                        if (ProductID.Contains("PB"))
                            return RedirectToAction("PutBinary", "BlackscholesPricers", new { ProductID = ProductID });
                    }
                }
            }

            UserMaster objUserMaster = new UserMaster();
            return View(objUserMaster);
        }

        [HttpGet]
        public ActionResult UserNotAuthorize()
        {
            return View();
        }

        [HttpPost]
        public ActionResult Login(UserMaster objUserMaster)
        {
            if (AuthenticateUser(objUserMaster.UserName, objUserMaster.Password))
            {
                //string Role = Convert.ToString(Session["Role"]);
                //if (Role.ToUpper() == "SALES")
                //    return RedirectToAction("FixedCouponList", "BlackscholesPricers");
                //else if (Role.ToUpper() == "TRADING")
                //    return RedirectToAction("UnderlyingCreationList", "UnderlyingCreation");
                //else if (Role.ToUpper() == "MID OFFICE")
                //    return RedirectToAction("MTMReport", "SPNoteMTM");

                return RedirectToAction("Index", "Dashboard");
            }

            ViewBag.Message = "Invalid Credentials";
            return View(objUserMaster);//RedirectToAction("ErrorPage", "Login");

        }

        public bool AuthenticateUser(string UserName, string Password)
        {
            try
            {
                if (ValidateUser(UserName, Password))
                {
                    UserMaster objUserMaster = new UserMaster();

                    SP_PRICINGEntities objSP_PRICINGEntities = new SP_PRICINGEntities();

                    ObjectResult<ValidateUserResult> objUserListResult = objSP_PRICINGEntities.SP_VALIDATE_USER(UserName, Password);
                    List<ValidateUserResult> UserList = objUserListResult.ToList();

                    if (UserList != null && UserList.Count == 1)
                    {
                        General.ReflectSingleData(objUserMaster, UserList[0]);
                        objUserMaster.Password = Password;
                        Session["LoggedInUser"] = objUserMaster;
                        Session["Role"] = objUserMaster.RoleName;
                        Session["LoginName"] = objUserMaster.UserName;                       

                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "", "", objUserMaster.UserID);
                return false;
            }
        }

        public bool AuthenticateMailUser(string UserName, string Password)
        {
            try
            {
                UserMaster objUserMaster = new UserMaster();

                SP_PRICINGEntities objSP_PRICINGEntities = new SP_PRICINGEntities();

                ObjectResult<ValidateUserResult> objUserListResult = objSP_PRICINGEntities.SP_VALIDATE_USER(UserName, Password);
                List<ValidateUserResult> UserList = objUserListResult.ToList();

                if (UserList != null && UserList.Count == 1)
                {
                    General.ReflectSingleData(objUserMaster, UserList[0]);
                    objUserMaster.Password = Password;
                    Session["LoggedInUser"] = objUserMaster;
                    Session["Role"] = objUserMaster.RoleName;

                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "", "", objUserMaster.UserID);
                return false;
            }
        }

        public bool ValidateUser(string UserName, string Password)
        {
            DirectoryEntry Entry = new DirectoryEntry("LDAP://EDELCAP", UserName, Password);
            DirectorySearcher Search = new DirectorySearcher(Entry);

            try
            {
                SearchResult results = Search.FindOne();
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private string windowAuthenticationSuccess(string LoginID)
        {
            try
            {
                DirectorySearcher searcher = new DirectorySearcher("");
                string filter = "(&(objectCategory=person)(objectClass=user)(|(samaccountname=" + LoginID + "*)))";

                DirectorySearcher search = new DirectorySearcher(filter);
                SearchResult result = search.FindOne();

                if (result != null)
                {
                    DirectoryEntry entry = result.GetDirectoryEntry();
                    string strEmpNo = entry.Properties["samaccountname"].Value.ToString();

                    return strEmpNo;
                }
                else
                    return "";
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void LogError(string strErrorDescription, string strStackTrace, string strClassName, string strMethodName, Int32 intUserId)
        {
            SP_PRICINGEntities objSP_PRICINGEntities = new SP_PRICINGEntities();
            var Count = objSP_PRICINGEntities.SP_ERROR_LOG(strErrorDescription, strStackTrace, strClassName, strMethodName, intUserId);
        }

        public ActionResult ErrorPage()
        {
            return View();
        }

        public ActionResult Logout()
        {
            LoginController objLoginController = new LoginController();

            try
            {
                Session["LoggedInUser"] = null;

                return RedirectToAction("Login", "Login");
            }
            catch (Exception ex)
            {
                UserMaster objCustomerDetails = (UserMaster)Session["LoggedInUser"];
                objLoginController.LogError(ex.Message, ex.StackTrace, "LoginController", "Logout Get", Convert.ToInt32(objCustomerDetails.UserID));
                return RedirectToAction("ErrorPage", "Login");
            }
        }

    }
}