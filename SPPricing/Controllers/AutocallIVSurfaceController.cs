using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using SPPricing.Models;
using System.Data;

namespace SPPricing.Controllers
{
    public class AutocallIVSurfaceController : Controller
    {
        //
        // GET: /AutocallIVSurface/

        public ActionResult Index()
        {
            return View();
        }

        [HttpGet]
        public ActionResult AutocallIVSurface()
        {
            return View();
        }

        [HttpPost]
        public ActionResult AutocallIVSurface(AutocallIVSurface objAutocallIVSurface)
        {
            return View();
        }

        public JsonResult FetchAutocallIVSurfaceList(string RedemptionDays, string NoOfSimulation, string IV, string RFR)
        {
            try
            {
                List<AutocallIVSurface> AutocallIVSurfaceList = new List<AutocallIVSurface>();

                if (RedemptionDays == "" || RedemptionDays == "0")
                    RedemptionDays = "-1";

                if (NoOfSimulation == "" || NoOfSimulation == "0")
                    NoOfSimulation = "-1";

                if (IV == "" || IV == "0")
                    IV = "-1";

                if (RFR == "" || RFR == "0")
                    RFR = "-1";

                DataSet dsResult = new DataSet();
                dsResult = General.ExecuteDataSet("SP_FETCH_AUTOCALL_IV_SURFACE_DETAILS", RedemptionDays, NoOfSimulation, IV, RFR);

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in dsResult.Tables[0].Rows)
                    {
                        AutocallIVSurface obj = new AutocallIVSurface();

                        obj.MaxRedemptionDays = Convert.ToInt32(dr["MaxRedemptionDays"]);
                        obj.MaxSimulation = Convert.ToInt32(dr["MaxSimulation"]);
                        obj.IV = Convert.ToDouble(dr["IV"]);
                        obj.RFR = Convert.ToDouble(dr["RFR"]);
                        
                        AutocallIVSurfaceList.Add(obj);
                    }
                }

                var AutocallIVSurfaceListData = AutocallIVSurfaceList.ToList();
                return Json(AutocallIVSurfaceListData, JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "FetchFixedCouponList", objUserMaster.UserID);
                return Json("");
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
    }
}