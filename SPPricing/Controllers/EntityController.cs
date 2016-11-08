using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using SPPricing.Models;
using System.Data.Objects;

namespace SPPricing.Controllers
{
    public class EntityController : Controller
    {
        //
        // GET: /Entity/

        SP_PRICINGEntities objSP_PRICINGEntities = new SP_PRICINGEntities();

        public ActionResult Index()
        {
            return View();
        }

        [HttpGet]
        public ActionResult EntityMaster(int? EntityID)
        {
            LoginController objLoginController = new LoginController();
            Entity objEntity = new Entity();

            try
            {
                if (ValidateSession())
                {
                    ObjectResult<EntityMasterResult> objEntityMasterResult = objSP_PRICINGEntities.SP_FETCH_ENTITY_MASTER_DETAILS(EntityID, "");
                    List<EntityMasterResult> EntityMasterResultList = objEntityMasterResult.ToList();

                    if (EntityMasterResultList != null && EntityMasterResultList.Count == 1)
                    {
                        objEntity.EntityID = EntityMasterResultList[0].EntityID;
                        objEntity.EntityCode = EntityMasterResultList[0].EntityCode;
                        objEntity.EntityName = EntityMasterResultList[0].EntityName;
                        objEntity.StatusText = EntityMasterResultList[0].StatusText;

                        if (EntityMasterResultList[0].Status == true)
                            objEntity.IsActive = true;
                        else
                            objEntity.IsActive = false;
                    }
                    else
                        objEntity.IsActive = true;

                    return View(objEntity);
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
                LogError(ex.Message, ex.StackTrace, "EntityController", "EntityMaster Get", objUserMaster.UserID);
                return RedirectToAction("ErrorPage", "Login");
            }
        }

        [HttpPost]
        public ActionResult EntityMaster(Entity objEntity)
        {
            return View();
        }

        public JsonResult FetchEntityMasterList(string EntityName)
        {
            try
            {
                if (ValidateSession())
                {
                    List<Entity> EntityList = new List<Entity>();

                    ObjectResult<EntityMasterResult> objEntityMasterResult = objSP_PRICINGEntities.SP_FETCH_ENTITY_MASTER_DETAILS(0, EntityName);
                    List<EntityMasterResult> EntityMasterResultList = objEntityMasterResult.ToList();

                    if (EntityMasterResultList != null && EntityMasterResultList.Count > 0)
                    {
                        foreach (EntityMasterResult oEntityMasterResult in EntityMasterResultList)
                        {
                            Entity objEntity = new Entity();

                            objEntity.EntityID = oEntityMasterResult.EntityID;
                            objEntity.EntityCode = oEntityMasterResult.EntityCode;
                            objEntity.EntityName = oEntityMasterResult.EntityName;
                            objEntity.StatusText = oEntityMasterResult.StatusText;
                            objEntity.CreatedBy = oEntityMasterResult.CreatedBy;
                            objEntity.CreatedOn = oEntityMasterResult.CreatedOn;

                            EntityList.Add(objEntity);
                        }
                    }

                    var EntityListData = EntityList.ToList();
                    return Json(EntityListData, JsonRequestBehavior.AllowGet);
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

        public Int32 ManageEntity(string EntityID, string EntityName, string IsActive)
        {
            UserMaster objUserMaster = new UserMaster();
            objUserMaster = (UserMaster)Session["LoggedInUser"];
            Int32 intResult = 0;

            try
            {
                if (ValidateSession())
                {
                    ProductParameter objProductParameter = new ProductParameter();

                    var Result = objSP_PRICINGEntities.SP_MANAGE_ENTITY_MASTER_LIST(Convert.ToInt32(EntityID), EntityName, Convert.ToBoolean(IsActive), objUserMaster.UserID);
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
    }
}