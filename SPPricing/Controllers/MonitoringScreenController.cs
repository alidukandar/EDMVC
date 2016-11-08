using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Data;
using SPPricing.Models;

namespace SPPricing.Controllers
{
    public class MonitoringScreenController : Controller
    {
        SP_PRICINGEntities objSP_PRICINGEntities = new SP_PRICINGEntities();
        //
        // GET: /MonitoringScreen/

        public ActionResult Index()
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

                    var Result = objSP_PRICINGEntities.VALIDATE_MENU_BY_ROLE(objUserMaster.RoleID, "MSI");
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
                LogError(ex.Message, ex.StackTrace, "MonitoringScreenController", "Index Get", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        public ActionResult Autocall()
        {
            return View();
        }

        public ActionResult Barrier()
        {
            return View();
        }

        #region "Bind Autocall List"

        public JsonResult FetchAutoCallReportList()
        {
            //try
            //{
            //    ObjectResult<WealthValuationResult> objWealthValuationResult = objSP_PRICINGEntities.FETCH_SP_WEALTH_VALUATION();
            //    List<WealthValuationResult> WealthValuationResultList = objWealthValuationResult.ToList();

            //    return Json(WealthValuationResultList, JsonRequestBehavior.AllowGet);
            //}
            //catch (Exception ex)
            //{
            //    LogError(ex.Message, ex.StackTrace, "PeriodicValuationController", "FetchWeeklyValuationList", objUserMaster.UserID);
            //    return Json("");
            //}

            DataSet dsReportAutoCall = General.ExecuteDataSet("SP_REP0RT_AUTOCALL");
            if (dsReportAutoCall != null && dsReportAutoCall.Tables.Count > 0)
            {
                if (dsReportAutoCall.Tables[0] != null && dsReportAutoCall.Tables[0].Rows.Count > 0)
                {
                    List<AutoCallReport> lstAutoCallReport = new List<AutoCallReport>();
                    for (int i = 0; i < dsReportAutoCall.Tables[0].Rows.Count; i++)
                    {
                        AutoCallReport obj = new AutoCallReport();
                        General.ReflectData(obj, dsReportAutoCall.Tables[0].Rows[i]);
                        lstAutoCallReport.Add(obj);
                    }
                    return Json(lstAutoCallReport, JsonRequestBehavior.AllowGet);
                }
            }
            return null;
        }

        #endregion

        #region "Bind Barrier List"
        public JsonResult FetchBarrierReportList()
        {
            DataSet dsReportBarrier = General.ExecuteDataSet("SP_REP0RT_BARRIER");
            if (dsReportBarrier != null && dsReportBarrier.Tables.Count > 0)
            {
                if (dsReportBarrier.Tables[0] != null && dsReportBarrier.Tables[0].Rows.Count > 0)
                {
                    List<BarrierReport> lstReportBarrier = new List<BarrierReport>();
                    for (int i = 0; i < dsReportBarrier.Tables[0].Rows.Count; i++)
                    {
                        BarrierReport obj = new BarrierReport();
                        General.ReflectData(obj, dsReportBarrier.Tables[0].Rows[i]);
                        lstReportBarrier.Add(obj);
                    }
                    return Json(lstReportBarrier, JsonRequestBehavior.AllowGet);
                }
            }
            return null;
        }
        #endregion

        #region "Export Options"
        public void ExportAutoCallToExcel()
        {
            DataSet dsReportAutoCall = General.ExecuteDataSet("SP_REP0RT_AUTOCALL");
            if (dsReportAutoCall != null && dsReportAutoCall.Tables.Count > 0)
            {
                if (dsReportAutoCall.Tables[0] != null && dsReportAutoCall.Tables[0].Rows.Count > 0)
                {
                    string strExportData = dsReportAutoCall.Tables[0].ToCSV();
                    Response.Clear();
                    Response.ClearContent();
                    Response.ClearHeaders();
                    Response.ContentType = "application/text";
                    Response.AddHeader("Content-Disposition", string.Format("attachment;filename=EXPORT_AUTOCALL.csv; size={0}", strExportData.Length));
                    Response.Output.Write(strExportData);
                    Response.Flush();
                    Response.End();
                }
            }
        }

        public void ExportBarrierToExcel()
        {
            DataSet dsReportBarrier = General.ExecuteDataSet("SP_REP0RT_BARRIER");
            if (dsReportBarrier != null && dsReportBarrier.Tables.Count > 0)
            {
                if (dsReportBarrier.Tables[0] != null && dsReportBarrier.Tables[0].Rows.Count > 0)
                {
                    string strExportData = dsReportBarrier.Tables[0].ToCSV();
                    Response.Clear();
                    Response.ClearContent();
                    Response.ClearHeaders();
                    Response.ContentType = "application/text";
                    Response.AddHeader("Content-Disposition", string.Format("attachment;filename=EXPORT_BARRIER.csv; size={0}", strExportData.Length));
                    Response.Output.Write(strExportData);
                    Response.Flush();
                    Response.End();
                }
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
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                objLoginController.LogError(ex.Message, ex.StackTrace, "MonitoringScreenController", "ValidateSession", -1);
                return false;
            }
        }
    }
}
