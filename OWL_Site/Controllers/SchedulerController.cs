using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;
using Kendo.Mvc.Extensions;
using Kendo.Mvc.UI;
using Microsoft.Ajax.Utilities;
using OWL_Site.Models;
using TaskScheduler;
using Task = TaskScheduler.Task;

//using Telerik.Charting;

namespace OWL_Site.Controllers
{
    public sealed class SchedulerController : Controller
    {
        private SchedulerMeetingService meetingService;

        public string spisok;

        public static string oplink;
        public static IEnumerable<MeetingViewModel> meetings_all, mettingsFiltered;

        public SchedulerController()
        {
            this.meetingService = new SchedulerMeetingService();
        }
        public JsonResult GetUsers([DataSourceRequest] DataSourceRequest request)
        {
            aspnetdbEntities db = new aspnetdbEntities();
            var data = db.AspNetUsers;
            return Json(data.ToDataSourceResult(request));
        }

        public JsonResult GetVmRs([DataSourceRequest] DataSourceRequest request)
        {
            aspnetdbEntities db = new aspnetdbEntities();
            var data = db.AllVmrs;
            return Json(data.ToDataSourceResult(request));
        }

        [Authorize]
        public ActionResult Index()
        {
            aspnetdbEntities db = new aspnetdbEntities();
            var data = db.AspNetUsers;
            IEnumerable<AspNetUser> filteredresult = data;
            ViewBag.Users = filteredresult;
            return View();
        }

        public JsonResult Meetings_Create([DataSourceRequest] DataSourceRequest request, MeetingViewModel meeting)
        {
            RegexUtilities util = new RegexUtilities();
            aspnetdbEntities db = new aspnetdbEntities();
            var rooms = db.AllVmrs;
            var init = GetAllPB().FirstOrDefault(m => m.Sammaccount == User.Identity.Name);
            var roomalias = db.VmrAliases.FirstOrDefault(m => m.vmid == meeting.RoomID);
            //meeting.RoomID = AccountController.UID;
            //int rid = meeting.RoomID;
            //var init = GetAllPB().FirstOrDefault(m => m.Id == rid);
            //meeting.OpLink = string.Concat("https://", "10.129.15.129", "/webapp/?conference=", init.Sammaccount, "&name=Operator&bw=512&join=1");
            //List<object> attend = meeting.Attendees.Select(att => GetAllPB().FirstOrDefault(m => m.Id == att)).Cast<object>().ToList();
            //meeting.InitName = init.Sammaccount;
            //meeting.InitFullname = init.DispName;
            if (meeting.Start < DateTime.Today + TimeSpan.FromHours(3))
            { Debug.WriteLine("@@@"); }
            List<AspNetUser> emaillist = new List<AspNetUser>();
            //emaillist.Add(init);
            StringBuilder strB = new StringBuilder();
            foreach (var att in meeting.Attendees)
            {
                //AspNetUser attemail = (GetAllPB().FirstOrDefault(m => m.Id == att));
                //emaillist.Add(attemail);
            }
            List<string> AddAtt = new List<string>();
            if (meeting.AddAttend != null) { AddAtt = (meeting.AddAttend.Split((",").ToCharArray())).ToList(); }
            foreach (var aa in AddAtt)
            {
                strB.Append(aa + ";" + aa + Environment.NewLine);
                if (util.IsValidEmail(aa))
                {
                    AspNetUser ar = new AspNetUser();
                    ar.Email = aa;
                    ar.DispName = aa;
                    emaillist.Add(ar);
                }
                else { }
            }

            foreach (var mail in emaillist)
            {

                if (mail.DispName.Contains(" "))
                {
                    mail.DispName = mail.DispName.Replace(" ", "%20");
                }
                else { }
                //string link = "https://" + "10.129.15.129" + "/webapp/?conference=" + init.DispName + "&name=" +
                //              Uri.EscapeDataString(mail.DispName) + "&bw=512&join=1";
                //string body = "Уважамый(ая), " + mail.DispName + "!" + Environment.NewLine + meeting.Start +
                //              TimeSpan.FromHours(3) + " состоится конференция на тему \"" + meeting.Title + "\"." +
                //              Environment.NewLine + "Инициатор конференции: " + init.DispName + Environment.NewLine +
                //              "В указанное время, для участия в конференции, просьба перейти по ссылке: " +
                //              Environment.NewLine + link;
                Debug.WriteLine(mail.Email);
                try
                {
                    //Sendmail(mail.email, meeting.Title, body);
                }
                catch (Exception e)
                {

                    Debug.WriteLine(e.Message);
                    Debug.WriteLine(e.HResult);
                }
           }

            //if (ModelState.IsValid)
            //{
            //    var owner = AccountController.currentuser;
            //    var filename = "meeting-" + owner.GivenName + "-" +
            //                       (meeting.Start + TimeSpan.FromHours(3)).ToString("dd-MM-yyyy_hh-mm") + ".csv";
            //    string path = Path.Combine(Server.MapPath("~/Content/OpFiles/CSV"), filename);
            //    Debug.WriteLine(path);
            //    meeting.FileLink = "Content/OpFiles/CSV/" + filename;

            //    Debug.WriteLine("Valid");
            //    using (FileStream fileStream = new FileStream(path, FileMode.OpenOrCreate))
            //    {
            //        using (StreamWriter streamWriter = new StreamWriter(fileStream, Encoding.UTF8))
            //        {
            //            streamWriter.Write(strB.ToString());
            //        }
            //    }
            //    if (meeting.Record)
            //    {
            //        Debug.WriteLine("Задача на запись создана");
            //        var tasktitle = String.Concat("rec-", owner.GivenName, "-",
            //            (meeting.Start + TimeSpan.FromHours(3)).ToString("dd-MM-yyyy_hh-mm"));
            //        var taskapp = Path.Combine(Server.MapPath("~/Content/OpFiles"), "flvstreamer.exe");
            //        var file_name = "rec-" + owner.GivenName + "-" +
            //                        (meeting.Start + TimeSpan.FromHours(3)).ToString("dd-MM-yyyy_hh-mm") + ".flv";
            //        var pathflv = Path.Combine(Server.MapPath("~/Content/OpFiles/FLV"), file_name);
            //        var stream_link = "rtmp://www.planeta-online.tv:1936/live/soyuz";
            //        meeting.Recfile = "Content/OpFiles/FLV/" + file_name;
            //        var comment = "Запись конференции " + owner.GivenName + "-" +
            //                      (meeting.Start + TimeSpan.FromHours(3)).ToString("dd-MM-yyyy_hh-mm");
            //        var acc_un = "boris_000";
            //        var acc_pass = "1Q2w3e4r!";
            //        var task_start = meeting.Start;
            //        var task_end = meeting.End;
            //        RecordTask(tasktitle, taskapp, pathflv, stream_link, comment, acc_un, acc_pass, task_start, task_end);

            //    }
            //    
            //}
            meetingService.Insert(meeting, ModelState);
            return Json(new[] { meeting }.ToDataSourceResult(request, ModelState));
        }
        public JsonResult Meetings_Destroy([DataSourceRequest] DataSourceRequest request, MeetingViewModel meeting)
        {
            if (ModelState.IsValid)
            {
                meetingService.Delete(meeting, ModelState);
            }
            return Json(new[] { meeting }.ToDataSourceResult(request, ModelState));
        }
        public JsonResult Meetings_Read([DataSourceRequest] DataSourceRequest request)
        {
            meetings_all = meetingService.GetAll();
            mettingsFiltered = meetings_all.AsEnumerable().Where(m => m.InitName == User.Identity.Name);
            foreach (var all in meetings_all)
            {
                if (!all.Recfile.IsNullOrWhiteSpace())
                {
                    ViewBag.Rec = true;
                }
            }
            if (User.IsInRole("Admins"))
            {
              return Json(meetingService.GetAll().ToDataSourceResult(request));
            }
            if (User.IsInRole("User"))
            {
                return Json(mettingsFiltered.ToDataSourceResult(request));
            }
            return null;
        }
        public JsonResult Meetings_Update([DataSourceRequest] DataSourceRequest request, MeetingViewModel meeting)
        {
            if (ModelState.IsValid)
            {
                meetingService.Update(meeting, ModelState);
                RegexUtilities util = new RegexUtilities();
                var idf = GetAllPB().FirstOrDefault(m => m.Sammaccount == User.Identity.Name);
                //meeting.RoomID = AccountController.UID;
                int rid = meeting.RoomID;
                //var init = GetAllPB().FirstOrDefault(m => m.Id == rid);
                //meeting.OpLink = string.Concat("https://", "10.129.15.129", "/webapp/?conference=", init.Sammaccount, "&name=Operator&bw=512&join=1");
                //var attend = meeting.Attendees.Select(att => GetAllPB().FirstOrDefault(m => m.Id == att)).Cast<object>().ToList();
                StringBuilder strB = new StringBuilder();
                List<string> AddAtt = new List<string>();
                if (meeting.AddAttend != null) { AddAtt = (meeting.AddAttend.Split((",").ToCharArray())).ToList(); }
                //foreach (int at in attend)
                //{
                //    string em = at.
                //var init = AccountController.currentuser.DisplayName;
                //var name = AccountController.currentuser;
                ////spisok = string.Concat(name.PhoneNumber, ";", name.UserName);
                //strB.Append(name.VoiceTelephoneNumber + ";" + name.DisplayName + Environment.NewLine);
                //if (name.EmailAddress != null)
                //{
                //    string link = "https://" + "10.129.15.129" + "/webapp/?conference=" + init.UserConfID + "&name=" +
                //                  Uri.EscapeDataString(name.UserName) + "&bw=512&join=1";
                //    string body = "Уважамый(ая), " + name.UserName + "!" + Environment.NewLine +
                //                  " конференция на тему \"" + meeting.Title + "\" переносится на " + meeting.Start +
                //                  TimeSpan.FromHours(3) + Environment.NewLine + "Инициатор конференции: " +
                //                  init.UserName + Environment.NewLine +
                //                  "В указанное время, для участия в конференции, просьба перейти по ссылке: " +
                //                  Environment.NewLine + link;
                //    //Sendmail(name.Email, meeting.Title, body);
                //}
                //}
                //foreach (var aa in AddAtt)
                //{
                //    strB.Append(aa + ";" + aa + Environment.NewLine);
                //}
                ////var owner = applicationDbContext.Users.FirstOrDefault(p => p.UserConfID == roomID);
                //var filename = "meeting-" + AccountController.Uname + "-" +
                //                   (meeting.Start + TimeSpan.FromHours(3)).ToString("dd-MM-yyyy_hh-mm") + ".csv";
                //string path = Path.Combine(Server.MapPath("~/Content/OpFiles/CSV"), filename);
                //Debug.WriteLine(path);
                //meeting.FileLink = "Content/OpFiles/CSV/" + filename;
                //meetingService.Update(meeting, ModelState);

                //using (FileStream fileStream = new FileStream(path, FileMode.OpenOrCreate))
                //{
                //    using (StreamWriter streamWriter = new StreamWriter(fileStream, Encoding.UTF8))
                //    {
                //        streamWriter.Write(strB.ToString());
                //    }
                //}
            }
            return Json(new[] { meeting }.ToDataSourceResult(request, ModelState));
        }

        public Task<ActionResult> Sendmail(string to, string subj, string body)
        {
            SmtpClient smtpClient = new SmtpClient(MvcApplication.set.AuthDnAddress, 25)
            {
                UseDefaultCredentials = false,
                EnableSsl = true,

                Credentials = new NetworkCredential("cobaservice", "Ciscocisco123"),

                DeliveryMethod = SmtpDeliveryMethod.Network,
                Timeout = 20000



            };
            MailMessage mailMessage = new MailMessage()
            {
                Priority = MailPriority.High,
                From = new MailAddress("cobaservice@nkc.ru", "Планировщик системы видео-конференц-связи 'Сова'")
            };
            mailMessage.To.Add(new MailAddress(to));
            mailMessage.Subject = subj;
            mailMessage.Body = body;

            smtpClient.Send(mailMessage);
            return null;
        }

        public Task<ActionResult> RecordTask(string tasktitle, string taskapp, string stream_link, string pathflv, string comment, string acc_un, string accpass, DateTime task_start, DateTime task_end)
        {
            ScheduledTasks st = new ScheduledTasks();
            Task t;
            t = st.CreateTask(tasktitle);
            t.ApplicationName = taskapp;
            t.Parameters = " -r " + "-q" + pathflv + " -o " + "\"" + stream_link + "\"";
            t.Comment = comment;
            t.SetAccountInformation(acc_un, accpass);
            t.IdleWaitMinutes = 10;
            TimeSpan worktime = task_end - task_start;
            Debug.WriteLine(worktime);
            t.MaxRunTime = new TimeSpan(worktime.Ticks);

            t.Priority = System.Diagnostics.ProcessPriorityClass.Idle;
            DateTime starttask = new DateTime();
            starttask = task_start + TimeSpan.FromHours(3);
            t.Triggers.Add(new RunOnceTrigger(starttask));
            t.Save();
            t.Close();
            st.Dispose();
            return null;
        }
        private IEnumerable<AspNetUser> GetPB()
        {
            var data = GetPhBOw(User.Identity.Name);
            return data;
        }
        private IEnumerable<AspNetUser> GetAllPB()
        {
            var data = GetPB();
            return data;
        }
        public List<AspNetUser> GetPhBOw(string Owname) //get phonebook
        {
            aspnetdbEntities db = new aspnetdbEntities();
            List<AspNetUser> selrec = new List<AspNetUser>();
            var selectets = db.PrivatePhBs.Where(m => m.OwSAN == Owname);
            foreach (var sel in selectets)
            {
                AspNetUser temp = new AspNetUser();
                var srec = db.AspNetUsers.FirstOrDefault(m => m.Id == sel.IdREC);
                temp = srec;
                if (!String.IsNullOrEmpty(sel.Group))
                {
                    temp.Group = sel.Group;
                }
                if (String.IsNullOrEmpty(sel.Group))
                {
                    temp.Group = "Группа не назначена";
                }
                selrec.Add(temp);

            }

            return selrec;
        }
    }
}