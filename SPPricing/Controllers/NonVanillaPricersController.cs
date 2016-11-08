using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace SPPricing.Controllers
{
    public class NonVanillaPricersController : Controller
    {
        //
        // GET: /NonVanillaPricers/

        public ActionResult Index()
        {
            return View();
        }

        public ActionResult Barrier()
        {
            return View();
        }

        public ActionResult DRA()
        {
            return View();
        }

        public ActionResult Generic()
        {
            return View();
        }
    }
}
