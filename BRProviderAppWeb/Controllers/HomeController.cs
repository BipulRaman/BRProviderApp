using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace BRProviderAppWeb.Controllers
{
    public class HomeController : Controller
    {
      
        [SharePointContextFilter]
        public ActionResult Index()
        {
            User spUser = null;

            Session["appweburl"] = Request.QueryString["SPAppWebUrl"];
            ViewBag.url = Session["appweburl"];
            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);

            using (var clientContext = spContext.CreateUserClientContextForSPHost())
            {
                if (clientContext != null)
                {
                    spUser = clientContext.Web.CurrentUser;
                    clientContext.Load(spUser, user => user.Title);
                    clientContext.ExecuteQuery();
                    ViewBag.UserName = spUser.Title;
                                       
                }
            }

            return View();
        }

        public ActionResult About()
        { 
            ViewBag.Message = "About U-Expense App";
            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "You can esaily reach us via :";
            return View();
        }

        public ActionResult App()
        {
            ViewBag.Message = "This is just another App";
            return View();
        }

              
    }
}
