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
        private aspnetdbEntities db = new aspnetdbEntities();

        public ActionResult Index()
        {
            return View();
        }

        public ActionResult Settings_Read([DataSourceRequest]DataSourceRequest request)
        {
            IQueryable<Setting> settings = db.Settings;
            DataSourceResult result = settings.ToDataSourceResult(request, setting => new {
                Id = setting.Id,
                AuthDnAddress = setting.AuthDnAddress,
                OU = setting.OU,
                UserGroup = setting.UserGroup,
                AdminGroup = setting.AdminGroup,
                DnAdminUn = setting.DnAdminUn,
                DnAdminPass = setting.DnAdminPass,
                CobaMngAddress = setting.CobaMngAddress,
                CobaCfgAddress = setting.CobaCfgAddress,
                CobaRecordsAddress = setting.CobaRecordsAddress,
                CobaRecLogin = setting.CobaRecLogin,
                CobaRecPass = setting.CobaRecPass,
                CobaRecBdName = setting.CobaRecBdName,
                CobaRecBdTable = setting.CobaRecBdTable,
                SmtpServer = setting.SmtpServer,
                SmtpPort = setting.SmtpPort,
                SmtpSSL = setting.SmtpSSL,
                SmtpLogin = setting.SmtpLogin,
                SmtpPassword = setting.SmtpPassword,
                MailFrom_email = setting.MailFrom_email,
                MailFrom_name = setting.MailFrom_name,
                CobaMngLogin = setting.CobaMngLogin,
                CobaMngPass = setting.CobaMngPass
            });

            return Json(result);
        }

        protected override void Dispose(bool disposing)
        {
            db.Dispose();
            base.Dispose(disposing);
        }
    }
}
