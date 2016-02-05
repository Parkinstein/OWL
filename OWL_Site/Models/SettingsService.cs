using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Web;

namespace OWL_Site.Models
{
    public class SettingsService : IDisposable
    {
        private aspnetdbEntities entities;
        public SettingsService(aspnetdbEntities entities)
        {
            this.entities = entities;
        }
        public IEnumerable<SettingsViewModel> Read()
        {
            return entities.Settings.Select(setting => new SettingsViewModel
            {
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
            }).AsEnumerable();
        }
        public void Create(SettingsViewModel setting)
        {
            var entity = new Setting
            {
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
            };

            entities.Settings.Add(entity);
            entities.SaveChanges();
            setting.Id = entity.Id;
        }
        public void Update(SettingsViewModel setting)
        {
            var entity = new Setting
            {
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
            };
            entities.Settings.Attach(entity);
            entities.Entry(entity).State = EntityState.Modified;
            entities.SaveChanges();
            setting.Id = entity.Id;

        }
        public void Dispose()
        {
            entities.Dispose();
        }
    }
}