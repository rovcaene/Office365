﻿using System.Web.Mvc;
using System.Configuration;
using System;
using U2U.Provisioning.SiteCollectionCreation.Utils;

namespace U2U.Provisioning.SiteCollectionCreation.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            if (ConfigurationManager.AppSettings["ida:ClientID"] == null)
            {
                ViewBag.DidNotAddConnectedServices = true;
            }
            else
            {
                ViewBag.DidNotAddConnectedServices = false;
            }
            
            string tenantID = ConfigurationManager.AppSettings["ida:TenantID"] ?? "" ;
            Guid resultGuid;
            
            try
            {
                resultGuid = Guid.ParseExact(tenantID, "D");
                ViewBag.TenantIDIsNull = false;
            }
            catch (ArgumentNullException) 
            { 
               ViewBag.TenantIDIsNull = true;
            }   
            catch (FormatException) 
            {
                ViewBag.TenantIDIsNull = true;
            }

            return View();
        }
    }
}