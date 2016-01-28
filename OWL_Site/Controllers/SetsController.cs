﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using Kendo.Mvc.Extensions;
using Kendo.Mvc.UI;
using OWL_Site.Models;

namespace OWL_Site
{
    public class SetsController : Controller
    {
        private SettingsService settingsService;

        public SetsController()
        {
            settingsService =new SettingsService(new aspnetdbEntities());
        }
        protected override void Dispose(bool disposing)
        {
            settingsService.Dispose();

            base.Dispose(disposing);
        }

        public ActionResult Index()
        {
            return View();
        }

        public ActionResult Sett_Read([DataSourceRequest] DataSourceRequest request)
        {
            using (var northwind = new aspnetdbEntities())
            {
                IQueryable<Setting> products = northwind.Settings;
                DataSourceResult result = products.ToDataSourceResult(request);
                return Json(result);
            }
        }
        [AcceptVerbs(HttpVerbs.Post)]
        public ActionResult Sett_Create([DataSourceRequest] DataSourceRequest request, SettingsViewModel settings)
        {
            if (settings != null && ModelState.IsValid)
            {
                settingsService.Create(settings);
            }

            return Json(new[] { settings }.ToDataSourceResult(request, ModelState));
        }
        [AcceptVerbs(HttpVerbs.Post)]
        public ActionResult Sett_Update([DataSourceRequest] DataSourceRequest request, SettingsViewModel settings)
        {
            if (settings != null && ModelState.IsValid)
            {
                settingsService.Update(settings);
            }

            return Json(new[] { settings }.ToDataSourceResult(request, ModelState));
        }


    }
}
