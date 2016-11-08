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
    public class ProductParametersController : Controller
    {
        //
        // GET: /ProductParameters/

        SP_PRICINGEntities objSP_PRICINGEntities = new SP_PRICINGEntities();

        public ActionResult Index()
        {
            return View();
        }

        [HttpGet]
        public ActionResult ProductParameters(int? underlyingID)
        {
            LoginController objLoginController = new LoginController();
            ProductParameter objProductParameter = new ProductParameter();

            try
            {
                if (ValidateSession())
                {
                    ObjectResult<ProductParameterResult> objProductParameterResult = objSP_PRICINGEntities.FETCH_PRODUCT_PARAMETERS_DETAILS(Convert.ToInt32(underlyingID), -1, -1, 0);
                    List<ProductParameterResult> ProductParameterResultList = objProductParameterResult.ToList();

                    if (ProductParameterResultList != null && ProductParameterResultList.Count == 1)
                    {
                        objProductParameter.ID = ProductParameterResultList[0].ID;
                        objProductParameter.UnderlyingID = ProductParameterResultList[0].UnderlyingID;
                        objProductParameter.IV = ProductParameterResultList[0].IV;
                        objProductParameter.RFR = ProductParameterResultList[0].RFR;
                        objProductParameter.BuiltInAdjustment = ProductParameterResultList[0].BuiltInAdjustment;
                    }

                    return View(objProductParameter);
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
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "ProductParameters Get", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        [HttpPost]
        public ActionResult ProductParameters(ProductParameter objProductParameter, string Command, FormCollection collection, HttpPostedFileBase file)
        {
            LoginController objLoginController = new LoginController();
            List<UploadFileMaster> UploadFileMasterList = (List<UploadFileMaster>)Session["UploadFileMasterList"];

            try
            {
                if (ValidateSession())
                {

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
                LogError(ex.Message, ex.StackTrace, "UnderlyingCreationController", "ProductParameters Post", objUserMaster.UserID);
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
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                objLoginController.LogError(ex.Message, ex.StackTrace, "MCPricersController", "ValidateSession", -1);
                return false;
            }
        }

        public void LogError(string strErrorDescription, string strStackTrace, string strClassName, string strMethodName, Int32 intUserId)
        {
            SP_PRICINGEntities objSP_PRICINGEntities = new SP_PRICINGEntities();
            var Count = objSP_PRICINGEntities.SP_ERROR_LOG(strErrorDescription, strStackTrace, strClassName, strMethodName, intUserId);
        }

        public JsonResult FetchProductParameterList(string UnderlyingID, string IV, string RFR, string BuiltInAdjustment)
        {
            try
            {
                if (ValidateSession())
                {
                    List<ProductParameter> ProductParameterList = new List<ProductParameter>();

                    if (UnderlyingID == "" || UnderlyingID == "--Select--")
                        UnderlyingID = "-1";

                    if (IV == "" || IV == "0")
                        IV = "-1";

                    if (RFR == "" || RFR == "0")
                        RFR = "-1";

                    if (BuiltInAdjustment == "")
                        BuiltInAdjustment = "0";

                    ObjectResult<ProductParameterResult> objProductParameterResult = objSP_PRICINGEntities.FETCH_PRODUCT_PARAMETERS_DETAILS(Convert.ToInt32(UnderlyingID), Convert.ToDouble(IV), Convert.ToDouble(RFR), Convert.ToDouble(BuiltInAdjustment));
                    List<ProductParameterResult> ProductParameterResultList = objProductParameterResult.ToList();

                    if (ProductParameterResultList != null && ProductParameterResultList.Count > 0)
                    {
                        foreach (ProductParameterResult oProductParameterResult in ProductParameterResultList)
                        {
                            ProductParameter objProductParameter = new ProductParameter();

                            objProductParameter.ID = oProductParameterResult.ID;
                            objProductParameter.UnderlyingID = oProductParameterResult.UnderlyingID;
                            objProductParameter.UnderlyingName = oProductParameterResult.UnderlyingName;
                            objProductParameter.IV = oProductParameterResult.IV;
                            objProductParameter.RFR = oProductParameterResult.RFR;
                            objProductParameter.BuiltInAdjustment = oProductParameterResult.BuiltInAdjustment;

                            ProductParameterList.Add(objProductParameter);
                        }
                    }

                    var ProductParameterListData = ProductParameterList.ToList();
                    return Json(ProductParameterListData, JsonRequestBehavior.AllowGet);
                }
                else
                    return Json("");
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "FetchFixedCouponMLDList", objUserMaster.UserID);
                return Json("");//("Index", "ErrorDetails");
            }
        }

        public Int32 ManageProductParameter(string ID, string UnderlyingID, string IV, string RFR, string BuiltInAdjustment)
        {
            UserMaster objUserMaster = new UserMaster();
            objUserMaster = (UserMaster)Session["LoggedInUser"];
            Int32 intResult = 0;

            try
            {
                if (ValidateSession())
                {
                    ProductParameter objProductParameter = new ProductParameter();

                    var Result = objSP_PRICINGEntities.SP_MANAGE_PRODUCT_PARAMETERS(Convert.ToInt32(ID), Convert.ToInt32(UnderlyingID), Convert.ToDouble(IV), Convert.ToDouble(RFR), Convert.ToDouble(BuiltInAdjustment), objUserMaster.UserID);
                    intResult = Convert.ToInt32(Result.SingleOrDefault());
                }

                return intResult;
            }
            catch (Exception ex)
            {
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "FetchFixedCouponMLDList", objUserMaster.UserID);
                return 0;
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
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "FetchUnderlyingList", objUserMaster.UserID);
                return Json("");
            }

            //return Json(UnderlyingListData);
        }

        public JsonResult FetchUnderlyingListStringData()
        {
            try
            {
                string strData = "";
                DataSet dsResult = General.ExecuteDataSet("SP_FETCH_UNDERLYING_DETAILS");

                strData += "-1:--Select--;";

                if (dsResult != null && dsResult.Tables.Count > 0 && dsResult.Tables[0].Rows.Count > 0)
                {
                    foreach (DataRow dr in dsResult.Tables[0].Rows)
                    {
                        strData += dr["UnderlyingID"] + ":" + dr["UnderlyingShortName"] + ";";
                    }
                }

                return Json(strData);
            }
            catch (Exception ex)
            {
                UserMaster objUserMaster = new UserMaster();
                objUserMaster = (UserMaster)Session["LoggedInUser"];
                LogError(ex.Message, ex.StackTrace, "BlackscholesPricersController", "FetchUnderlyingList", objUserMaster.UserID);
                return Json("");
            }

            //return Json(UnderlyingListData);
        }

        public JsonResult AddProductParameters(FormCollection frmCollection)
        {
            string strProductParamID = "0";
            string strUnderlyingID = frmCollection["UnderlyingID"];
            string strUnderlyingName = frmCollection["UnderlyingName"];
            string strIV = frmCollection["IV"];
            string strRFR = frmCollection["RFR"];
            string strBuiltInAdjustment = frmCollection["BuiltInAdjustment"];

            Int32 intResult = ManageProductParameter(strProductParamID, strUnderlyingName, strIV, strRFR, strBuiltInAdjustment);
            string strResult = "";

            if (intResult == -1)
                strResult = "Already Exists";
            else if (intResult == 1)
                strResult = "Details Saved Successfully";
            else if (intResult == 2)
                strResult = "Details Updated Successfully";

            return Json(strResult);
        }

        public JsonResult EditProductParameters(FormCollection frmCollection)
        {
            string strProductParamID = frmCollection["ID"];
            strProductParamID = strProductParamID.Split(',')[0];

            string strUnderlyingID = frmCollection["UnderlyingID"];
            string strUnderlyingName = frmCollection["UnderlyingName"];
            string strIV = frmCollection["IV"];
            string strRFR = frmCollection["RFR"];
            string strBuiltInAdjustment = frmCollection["BuiltInAdjustment"];

            Int32 intResult = ManageProductParameter(strProductParamID, strUnderlyingID, strIV, strRFR, strBuiltInAdjustment);
            string strResult = "";

            if (intResult == -1)
                strResult = "Already Exists";
            else if (intResult == 1)
                strResult = "Details Saved Successfully";
            else if (intResult == 2)
                strResult = "Details Updated Successfully";

            return Json(strResult);
        }
    }
}