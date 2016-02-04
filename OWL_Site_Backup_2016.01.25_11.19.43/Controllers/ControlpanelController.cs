using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Text;
using System.Web.Mvc;
using Newtonsoft.Json;
using OWL_Site.Models;

namespace OWL_Site.Controllers
{
    public class ControlpanelController : Controller
    {
        public List<ActiveConfsModel.AConfs> AllConfs;
        public ActiveConfsModel.ResponseParent AllConfs_wm;
        public List<ActivePartsModel.AParts> AllParts;
        public ActivePartsModel.ResponseParent AllParts_wm;

        private string Win1251ToUTF8(string source)
        {
            Encoding utf8 = Encoding.GetEncoding("windows-1251");
            Encoding win1251 = Encoding.GetEncoding("utf-8");

            byte[] utf8Bytes = win1251.GetBytes(source);
            byte[] win1251Bytes = Encoding.Convert(win1251, utf8, utf8Bytes);
            source = win1251.GetString(win1251Bytes);
            return source;
        }
        // GET: Controlpanel
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult ActiveParts_Ajax(ActiveConference.DTResult param, string confname)
        {
            IEnumerable<ActivePartsModel.AParts> filteredresult;

            if (!string.IsNullOrEmpty(param.Search.Value))
            {
                filteredresult = GetActiveParts(confname).Where(c => c.display_name.Contains(param.Search.Value));
            }
            else
            {
                filteredresult = GetActiveParts(confname);
            }

            return Json(new
            {
                recordsTotal = GetActiveParts(confname).Count(),
                recordsFiltered = filteredresult.Count(),
                data = filteredresult,
            }, JsonRequestBehavior.AllowGet);
        }


        List<ActivePartsModel.AParts> GetActiveParts(string confname)
        {
            try
            {
                AllParts = new List<ActivePartsModel.AParts>();
                Uri statusapi =
                    new Uri("https://" + MvcApplication.set.CobaMngAddress +
                            "/api/admin/status/v1/participant/?conference=" + confname);
                WebClient client = new WebClient();
                client.Credentials = new NetworkCredential(MvcApplication.set.CobaMngLogin, MvcApplication.set.CobaMngPass);
                client.Headers.Add("auth", MvcApplication.set.CobaMngLogin+","+ MvcApplication.set.CobaMngPass);
                client.Headers.Add("veryfy", "False");
                ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
                string reply = client.DownloadString(statusapi);
                string reply1 = Win1251ToUTF8(reply);
                if (!String.IsNullOrEmpty(reply))
                {
                    Debug.WriteLine(reply);
                    AllParts_wm = JsonConvert.DeserializeObject<ActivePartsModel.ResponseParent>(reply1);
                    AllParts = AllParts_wm.obj;
                    AllParts.FindAll(p => p.has_media);
                    foreach (var part in AllParts)
                    {
                        Debug.WriteLine(part.display_name);
                    }
                    
                }
                return AllParts;
            }
            catch (Exception errException)
            {
                Debug.WriteLine(errException.Message);
            }
            return AllParts;
        }

        #region GetActiveConfs
        public ActionResult ActiveConf_Ajax(ActiveConference.DTResult param)
        {
            IEnumerable<ActiveConfsModel.AConfs> filteredresult;

            if (!string.IsNullOrEmpty(param.Search.Value))
            {
                filteredresult = GetData().Where(c => c.name.Contains(param.Search.Value));
            }
            else
            {
                filteredresult = GetData();
            }

            return Json(new
            {
                recordsTotal = GetData().Count(),
                recordsFiltered = filteredresult.Count(),
                data = filteredresult,
            }, JsonRequestBehavior.AllowGet);
        }

        private IEnumerable<ActiveConfsModel.AConfs> GetData()
        {
            var data = GetActiveConfs();

            return data;
        }

        public List<ActiveConfsModel.AConfs> GetActiveConfs()
        {
            try
            {

                AllConfs = new List<ActiveConfsModel.AConfs>();
                Uri statusapi = new Uri("https://" + MvcApplication.set.CobaMngAddress + "/api/admin/status/v1/conference/");

                WebClient client = new WebClient();
                client.Credentials = new NetworkCredential(MvcApplication.set.CobaMngLogin , MvcApplication.set.CobaMngPass);
                client.Headers.Add("auth", MvcApplication.set.CobaMngLogin + "," + MvcApplication.set.CobaMngPass);
                client.Headers.Add("veryfy", "False");
                ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
                string reply = client.DownloadString(statusapi);
                string reply1 = Win1251ToUTF8(reply);
                if (!String.IsNullOrEmpty(reply))
                {
                    AllConfs_wm = JsonConvert.DeserializeObject<ActiveConfsModel.ResponseParent>(reply1);
                    AllConfs = AllConfs_wm.obj;
                    foreach (var conf in AllConfs)
                    {
                        DateTime dt = DateTime.Parse(conf.start_time);
                        DateTime dt2 = dt + TimeSpan.FromHours(3);
                        string result = dt2.ToString("dd-MMM-yyyy  HH:mm:ss");
                        conf.start_time2 = result;
                        if (conf.is_locked)
                        {
                            conf.lock_path = "<img src=\"../images/lock.png\")\" style=\"max-width: 28px; max-height: 28px;\" />";
                        }
                        if (!conf.is_locked)
                        {
                            conf.lock_path = "<img src=\"../images/unlock.png\")\" style=\"max-width: 28px; max-height: 28px;\" />";
                        }

                    }
                }
                return AllConfs;
            }
            catch (Exception errException)
            {
                Debug.WriteLine(errException.Message);
            }
            return AllConfs;
        }
        #endregion
        public ActionResult Control(string confname, string dispname)
        {
            aspnetdbEntities db = new aspnetdbEntities();
            var curvmr = db.AllVmrs.FirstOrDefault(v => v.name == confname);
            var curalias = db.VmrAliases.FirstOrDefault(v => v.vmid == curvmr.Id);

            if (curvmr != null)
            {
                ViewData["dispnm"] = dispname;
                ViewData["pinc"] = curvmr.pin;
                ViewData["confnm"] = curalias.alias;
            }
            return View();
            
        }
    }
}