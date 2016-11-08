using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using SPPricing.Models;
using System.Data.Objects;
using System.Data;
using System.DirectoryServices;

namespace SPPricing.Controllers
{
    public class AdminModuleController : Controller
    {
        SP_PRICINGEntities objSP_PRICINGEntities = new SP_PRICINGEntities();

        public ActionResult Index()
        {
            return View();
        }

        [HttpGet]
        public ActionResult UserMaster()
        {
            UserMaster objUserMaster = new UserMaster();

            #region Menu Access By on Role

            Int32 intResult = 0;
            // bool PPorNonPP = false;

          
            objUserMaster = (UserMaster)Session["LoggedInUser"];

            var Result = objSP_PRICINGEntities.VALIDATE_MENU_BY_ROLE(objUserMaster.RoleID, "AMUM");
            intResult = Convert.ToInt32(Result.SingleOrDefault());

            if (intResult == 0)
                return RedirectToAction("UserNotAuthorize", "Login");

            #endregion

            List<RoleMaster> RoleMasterList = new List<RoleMaster>();

            ObjectResult<RoleResult> objRoleResult = objSP_PRICINGEntities.SP_FETCH_ROLE_MASTER_DETAILS();
            List<RoleResult> RoleResultList = objRoleResult.ToList();

            if (RoleResultList != null && RoleResultList.Count > 0)
            {
                foreach (RoleResult oRoleResult in RoleResultList)
                {
                    RoleMaster objRoleMaster = new RoleMaster();
                    General.ReflectSingleData(objRoleMaster, oRoleResult);
                    RoleMasterList.Add(objRoleMaster);
                }
            }

            objUserMaster.RoleList = RoleMasterList;

            return View(objUserMaster);
        }

        [HttpPost]
        public ActionResult UserMaster(UserMaster objUserMaster)
        {
            return View();
        }

        [HttpGet]
        public ActionResult RoleMaster()
        {

            #region Menu Access By on Role

            Int32 intResult = 0;
            // bool PPorNonPP = false;

            UserMaster objUserMaster = new UserMaster();
            objUserMaster = (UserMaster)Session["LoggedInUser"];

            var Result = objSP_PRICINGEntities.VALIDATE_MENU_BY_ROLE(objUserMaster.RoleID, "AMRM");
            intResult = Convert.ToInt32(Result.SingleOrDefault());

            if (intResult == 0)
                return RedirectToAction("UserNotAuthorize", "Login");

            #endregion

            return View();
        }

        [HttpPost]
        public ActionResult RoleMaster(RoleMaster objRoleMaster)
        {
            return View();
        }

        [HttpGet]
        public ActionResult RoleMenuMapping()
        {
            return View();
        }

        [HttpPost]
        public ActionResult RoleMenuMapping(UserMaster objUserMaster)
        {
            return View();
        }

        public JsonResult FetchRoleDetails()
        {
            List<RoleMaster> RoleMasterList = new List<RoleMaster>();

            ObjectResult<RoleResult> objRoleResult = objSP_PRICINGEntities.SP_FETCH_ROLE_MASTER_DETAILS();
            List<RoleResult> RoleResultList = objRoleResult.ToList();

            if (RoleResultList != null && RoleResultList.Count > 0)
            {
                foreach (RoleResult oRoleResult in RoleResultList)
                {
                    RoleMaster objRoleMaster = new RoleMaster();
                    General.ReflectSingleData(objRoleMaster, oRoleResult);
                    RoleMasterList.Add(objRoleMaster);
                }
            }

            var RoleMasterListData = RoleMasterList.ToList();
            return Json(RoleMasterListData, JsonRequestBehavior.AllowGet);
        }

        public JsonResult AddRole(FormCollection frmCollection)
        {
            int intRoleID = 0;
            UserMaster objUserMaster = new UserMaster();
            objUserMaster = (UserMaster)Session["LoggedInUser"];

            //int intUserID = 1;
            string strRoleName = frmCollection["RoleName"];
            string strRoleDescription = frmCollection["RoleDescription"];
            bool blnStatus = true;

            Int32 intResult = 0;
            var Count = objSP_PRICINGEntities.SP_MANAGE_ROLE_MASTER_DETAILS(intRoleID, strRoleName, strRoleDescription, blnStatus, objUserMaster.UserID);
            intResult = Count.SingleOrDefault().Value;

            string strResult = "";

            if (intResult == -1)
                strResult = "Role Already Exists";
            else if (intResult == 1)
                strResult = "Role Inserted Successfully";
            else if (intResult == 2)
                strResult = "Role Updated Successfully";

            return Json(strResult);
        }

        public JsonResult EditRole(FormCollection frmCollection)
        {
            int intRoleID = Convert.ToInt32(frmCollection["RoleID"]);
            UserMaster objUserMaster = new UserMaster();
            objUserMaster = (UserMaster)Session["LoggedInUser"];
            string strRoleName = frmCollection["RoleName"];
            string strRoleDescription = frmCollection["RoleDescription"];
            string strStatus = frmCollection["StatusText"];
            bool blnStatus = false;

            if (strStatus.ToUpper() == "ACTIVE")
                blnStatus = true;

            Int32 intResult = 0;
            var Count = objSP_PRICINGEntities.SP_MANAGE_ROLE_MASTER_DETAILS(intRoleID, strRoleName, strRoleDescription, blnStatus, objUserMaster.UserID);
            intResult = Count.SingleOrDefault().Value;

            string strResult = "";

            if (intResult == -1)
                strResult = "Role Already Exists";
            else if (intResult == 1)
                strResult = "Role Inserted Successfully";
            else if (intResult == 2)
                strResult = "Role Updated Successfully";

            return Json(strResult);
        }

        public JsonResult ManageRoleStatus(string RoleID, string Status)
        {
            int intRoleID = Convert.ToInt32(RoleID);
            bool blnStatus = Convert.ToBoolean(Status);

            Int32 intResult = 0;
            var Count = objSP_PRICINGEntities.SP_MANAGE_ROLE_STATUS(intRoleID, blnStatus);
            intResult = Count.SingleOrDefault().Value;

            string strResult = "";

            return Json(strResult);
        }

        public JsonResult ManageUserStatus(string UserID, string Status)
        {
            int intUserID = Convert.ToInt32(UserID);
            bool blnStatus = Convert.ToBoolean(Status);

            Int32 intResult = 0;
            var Count = objSP_PRICINGEntities.SP_MANAGE_USER_STATUS(intUserID, blnStatus);
            intResult = Count.SingleOrDefault().Value;

            string strResult = "";

            return Json(strResult);
        }

        public JsonResult FetchUserDetails()
        {
            List<UserMaster> UserMasterList = new List<Models.UserMaster>();

            ObjectResult<UserResult> objUserResult = objSP_PRICINGEntities.SP_FETCH_USER_MASTER_DETAILS();
            List<UserResult> UserResultList = objUserResult.ToList();

            if (UserResultList != null && UserResultList.Count > 0)
            {
                foreach (UserResult oUserResult in UserResultList)
                {
                    UserMaster objUserMaster = new UserMaster();
                    General.ReflectSingleData(objUserMaster, oUserResult);
                    UserMasterList.Add(objUserMaster);
                }
            }

            var UserMasterListData = UserMasterList.ToList();
            return Json(UserMasterListData, JsonRequestBehavior.AllowGet);
        }

        public JsonResult GetUsersAD(string SearchText)
        {
            SearchText = SearchText.Trim();
            //string filter = "(&(objectCategory=person)(objectClass=user)(|(givenName=" + SearchText + "*)(Description=" + SearchText + "*)(sn=" + SearchText + "*)))";
            string filter = "(&(objectCategory=person)(objectClass=user)(|(displayName=" + SearchText + "*)(Description=" + SearchText + "*)(sn=" + SearchText + "*)))";

            DirectorySearcher search = new DirectorySearcher(filter);
            List<UserMaster> UserMasterList = new List<UserMaster>();

            foreach (SearchResult result in search.FindAll())
            {
                DirectoryEntry entry = result.GetDirectoryEntry();

                string strFirstName = "";
                string strLastName = "";
                UserMaster objUserMaster = new UserMaster();

                objUserMaster.LoginName = entry.Properties["samaccountname"].Value != null ? entry.Properties["samaccountname"].Value.ToString() : "";

                strFirstName = entry.Properties["givenName"].Value != null ? entry.Properties["givenName"].Value.ToString() : "";
                strLastName = entry.Properties["sn"].Value != null ? entry.Properties["sn"].Value.ToString() : "";
                objUserMaster.UserName = strFirstName + " " + strLastName;

                objUserMaster.Department = entry.Properties["Department"].Value != null ? entry.Properties["Department"].Value.ToString() : "";
                objUserMaster.Email = entry.Properties["sn"] != null ? (entry.Properties["mail"].Value != null ? entry.Properties["mail"].Value.ToString() : "") : "";

                try
                {
                    objUserMaster.EmpID = Convert.ToInt32(entry.Properties["Description"].Value != null ? entry.Properties["Description"].Value.ToString() : "");
                }
                catch (Exception ex)
                {
                    objUserMaster.EmpID = 0;
                }

                UserMasterList.Add(objUserMaster);
            }

            var UserMasterListData = UserMasterList.ToList();
            return Json(UserMasterListData, JsonRequestBehavior.AllowGet);
        }

        public JsonResult AddUser(FormCollection frmCollection)
        {
            UserMaster objUserMaster = new UserMaster();
            objUserMaster = (UserMaster)Session["LoggedInUser"];

            int intLoggedUserID = 0;
            int intLoginID = Convert.ToInt32(frmCollection["EmpID"]);
            string strLoginName = frmCollection["LoginName"];
            string strUserName = frmCollection["UserName"];
            string strEmail = frmCollection["Email"];
            string strDepartment = frmCollection["Department"];
            string strStatus = frmCollection["StatusText"];
            bool blnStatus = false;
            int intRoleID = 0;

            

            if (strStatus.ToUpper() == "ACTIVE")
                blnStatus = true;

            Int32 intResult = 0;
            var Count = objSP_PRICINGEntities.SP_MANAGE_USER_MASTER_DETAILS(objUserMaster.UserID, intLoginID, strLoginName, strUserName, blnStatus, intRoleID, strDepartment, strEmail, intLoggedUserID);
            intResult = Count.SingleOrDefault().Value;

            string strResult = "";

            if (intResult == -1)
                strResult = "User Already Exists";
            else if (intResult == 1)
                strResult = "User Inserted Successfully";
            else if (intResult == 2)
                strResult = "User Updated Successfully";
            else if (intResult == 3)
                strResult = "User Deleted Successfully";

            return Json(strResult);
        }

        public JsonResult EditUser(FormCollection frmCollection)
        {
            return Json("");
        }

        public JsonResult DeleteUser(FormCollection frmCollection)
        {
            return Json("");
        }

        public JsonResult ManageUserDetails(string LoginID, string LoginName, string UserName, string Email, string Department, string Role)
        {
            List<UserMaster> UserMasterList = new List<UserMaster>();

            UserMaster objUserMaster = new UserMaster();
            objUserMaster = (UserMaster)Session["LoggedInUser"];

            Int32 intLoggedUserID = 1;
            bool blnStatus = true;
            //Role = "1";

            Int32 intResult = 0;
            var Count = objSP_PRICINGEntities.SP_MANAGE_USER_MASTER_DETAILS(0, Convert.ToInt32(LoginID), LoginName, UserName, blnStatus, Convert.ToInt32(Role), Department, Email, intLoggedUserID);
            intResult = Count.SingleOrDefault().Value;

            return Json(intResult);
        }
    }
}