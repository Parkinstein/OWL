using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;
using System.Threading.Tasks;
using System.Web.Mvc;
using Kendo.Mvc.Extensions;
using Kendo.Mvc.UI;
using MeetingRequest;
using OWL_Site.Models;
//using TaskScheduler;
//using Task = TaskScheduler.Task;

//using Telerik.Charting;

namespace OWL_Site.Controllers
{
    public sealed class SchedulerController : Controller
    {
        private SchedulerMeetingService meetingService;
        private string gp,gpmail;
        public static IEnumerable<MeetingViewModel> meetings_all, mettingsFiltered;
        public SchedulerController()
        {
            this.meetingService = new SchedulerMeetingService();
        }
        public JsonResult GetUsers([DataSourceRequest] DataSourceRequest request)
        {
            aspnetdbEntities db = new aspnetdbEntities();
            var data = db.AspNetUsers.AsEnumerable();
            return Json(data.ToDataSourceResult(request, o => new {id = o.Id,name = o.DispName}));
        }
        public JsonResult GetVmRs([DataSourceRequest] DataSourceRequest request)
        {
            aspnetdbEntities db = new aspnetdbEntities();
            var data = db.AllVmrs.Where(m => m.service_type == "conference");
            return Json(data.ToDataSourceResult(request));
        }
        [Authorize]
        public ActionResult Index()
        {
           return View();
        }
        public JsonResult Meetings_Create([DataSourceRequest] DataSourceRequest request, MeetingViewModel meeting)
        {
            RegexUtilities util = new RegexUtilities();
            aspnetdbEntities db = new aspnetdbEntities();
            var init = GetAllPB().FirstOrDefault(m => m.Sammaccount == User.Identity.Name);
            meeting.InitName = init.Id;
            meeting.FName = init.DispName;
            if (meeting.RoomID == 0)
            {
                meeting.RoomID = 1;
            }
            var roomalias = db.VmrAliases.FirstOrDefault(m => m.vmid == meeting.RoomID);
            var currentroom = db.AllVmrs.FirstOrDefault(m => m.Id == meeting.RoomID);
            meeting.OpLink = string.Concat("https://", MvcApplication.set.CobaCfgAddress, "/webapp/?conference=", roomalias.alias, "&name=Operator&bw=512&join=1");
            List<AspNetUser> emaillist = new List<AspNetUser>();
            //emaillist.Add(init);
            StringBuilder strB = new StringBuilder();
            if (meeting.Attendees != null)
            {
                foreach (var att in meeting.Attendees)
                {
                    AspNetUser attemail = (GetAllPB().FirstOrDefault(m => m.Id == att));
                    emaillist.Add(attemail);
                }
            }
            
            List<string> AddAtt = new List<string>();
            if (meeting.AddAttend != null) { AddAtt = (meeting.AddAttend.Split((",").ToCharArray())).ToList(); }
            if (!String.IsNullOrEmpty(currentroom.guest_pin))
            {
                gpmail = "PIN для входа: " + currentroom.guest_pin;
                gp = "&pin="+ currentroom.guest_pin;
            }
            if (String.IsNullOrEmpty(currentroom.guest_pin))
            {
                gpmail = "";
                gp = "&role=guest";
            }
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
                string link = "<a href =\"" + "https://" + MvcApplication.set.CobaCfgAddress + "/webapp/?conference=" +
                              roomalias.alias + "&name=" + Uri.EscapeDataString(mail.DispName) + "&bw=512" + gp +
                              "&join=1" +
                              "\">ссылке</a>";
                string body = "Уважамый(ая), " + mail.DispName + " !<br>" + (meeting.Start +
                                                                             TimeSpan.FromHours(3)).ToString("U") +
                              " состоится конференция на тему \"" +
                              meeting.Title + "\"." + "<br>" + "Инициатор конференции: " + init.DispName + "<br>" +
                              "В указанное время, для участия в конференции, просьба перейти по " +
                              link + "<br><br>" + "<b><i>Данные для самостоятельного входа:<i><b><br> Адрес сервера: " +
                              "https://" + MvcApplication.set.CobaCfgAddress + "/" + "<br>" +
                              "Имя конференции: " + roomalias.alias + "<br>" + gpmail + "<br>" + "SIP-адрес: " +
                              roomalias.alias + "@" + MvcApplication.set.CobaCfgAddress;
                Event new_event = new Event()
                {
                    UID = meeting.MeetingID.ToString(),
                    Location = "Moscow",
                    Status = EventStatus.CONFIRMED,
                    Organizer = new Organizer() {PublicName = init.DispName, Email = init.Email },
                    StartTime = meeting.Start,
                    EndTime = meeting.End,
                    Description = "Voici une conf",
                    Title = meeting.Title,
                    Attendees = { }
                };
            try
                {
                    Sendmail(mail.Email, "NEW: " + meeting.Title, body, meeting, new_event);
                }
                catch (Exception e)
                {

                    Debug.WriteLine(e.Message);
                    Debug.WriteLine(e.HResult);
                }
           }
            #region old

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

            #endregion
            meetingService.Insert(meeting, ModelState);
            return Json(new[] { meeting }.ToDataSourceResult(request, ModelState));
        }
        public JsonResult Meetings_Destroy([DataSourceRequest] DataSourceRequest request, MeetingViewModel meeting)
        {
            if (ModelState.IsValid)
            {
                meetingService.Delete(meeting, ModelState);
                RegexUtilities util = new RegexUtilities();
                aspnetdbEntities db = new aspnetdbEntities();
                var init = GetAllPB().FirstOrDefault(m => m.Sammaccount == User.Identity.Name);
                var currentroom = db.AllVmrs.FirstOrDefault(m => m.Id == meeting.RoomID);
                List<AspNetUser> emaillist = new List<AspNetUser>();
                StringBuilder strB = new StringBuilder();
                var roomalias = db.VmrAliases.FirstOrDefault(m => m.vmid == meeting.RoomID);
                foreach (var att in meeting.Attendees)
                {
                    AspNetUser attemail = (GetAllPB().FirstOrDefault(m => m.Id == att));
                    emaillist.Add(attemail);
                }
                List<string> AddAtt = new List<string>();
                if (meeting.AddAttend != null) { AddAtt = (meeting.AddAttend.Split((",").ToCharArray())).ToList(); }
                if (meeting.AddAttend != null)
                {
                    AddAtt = (meeting.AddAttend.Split((",").ToCharArray())).ToList();
                }
                if (!String.IsNullOrEmpty(currentroom.guest_pin))
                {
                    gpmail = "PIN для входа: " + currentroom.guest_pin;
                    gp = "&pin=" + currentroom.guest_pin;
                }
                if (String.IsNullOrEmpty(currentroom.guest_pin))
                {
                    gpmail = "";
                    gp = "&role=guest";
                }
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
                    string body = "Уважамый(ая), " + mail.DispName + " !<br>" + "Конференция \"" + meeting.Title +
                                  "\" была отменена пользователем " + init.DispName + ". <br>";
                    try
                    {
                        Sendmail(mail.Email, "DEL: " + meeting.Title, body,meeting, null);
                    }
                    catch (Exception e)
                    {

                        Debug.WriteLine(e.Message);
                        Debug.WriteLine(e.HResult);
                    }
                }
            }
            return Json(new[] { meeting }.ToDataSourceResult(request, ModelState));
        }
        public JsonResult Meetings_Read([DataSourceRequest] DataSourceRequest request)
        {
            meetings_all = meetingService.GetAll();
            var init = GetAllPB().FirstOrDefault(m => m.Sammaccount == User.Identity.Name);
            
            mettingsFiltered = meetings_all.AsEnumerable().Where(m => m.InitName == init.Id);
            Debug.WriteLine(mettingsFiltered.GetEnumerator().Current);
            foreach (var all in meetings_all)
            {
                if (!String.IsNullOrEmpty(all.Recfile))
                {
                    ViewBag.Rec = true;
                }
            }
            if (User.IsInRole("Admin"))
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
                
                RegexUtilities util = new RegexUtilities();
                aspnetdbEntities db = new aspnetdbEntities();
                var init = GetAllPB().FirstOrDefault(m => m.Sammaccount == User.Identity.Name);
                
                var currentroom = db.AllVmrs.FirstOrDefault(m => m.Id == meeting.RoomID);
                List<AspNetUser> emaillist = new List<AspNetUser>();
                StringBuilder strB = new StringBuilder();
                var roomalias = db.VmrAliases.FirstOrDefault(m => m.vmid == meeting.RoomID);
                meeting.OpLink = string.Concat("https://", MvcApplication.set.CobaCfgAddress, "/webapp/?conference=", roomalias.alias, "&name=Operator&bw=512&join=1");
                meetingService.Update(meeting, ModelState);
                foreach (var att in meeting.Attendees)
                {
                    AspNetUser attemail = (GetAllPB().FirstOrDefault(m => m.Id == att));
                    emaillist.Add(attemail);
                }
                List<string> AddAtt = new List<string>();
                if (meeting.AddAttend != null) { AddAtt = (meeting.AddAttend.Split((",").ToCharArray())).ToList(); }
                if (meeting.AddAttend != null)
                {
                    AddAtt = (meeting.AddAttend.Split((",").ToCharArray())).ToList();
                }
                if (!String.IsNullOrEmpty(currentroom.guest_pin))
                {
                    gpmail = "PIN для входа: " + currentroom.guest_pin;
                    gp = "&pin=" + currentroom.guest_pin;
                }
                if (String.IsNullOrEmpty(currentroom.guest_pin))
                {
                    gpmail = "";
                    gp = "&role=guest";
                }
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
                    string link = "<a href =\"" + "https://" + MvcApplication.set.CobaCfgAddress + "/webapp/?conference=" +
                    roomalias.alias + "&name=" + Uri.EscapeDataString(mail.DispName) + "&bw=512" + gp + "&join=1" +
                    "\">ссылке</a>";
                    string body = "Уважамый(ая), " + mail.DispName + " !<br>" + "Параметры конференции \"" + meeting.Title + "\" были изменены." + "<br>" + "Время начала конференции: " + (meeting.Start +
                                  TimeSpan.FromHours(3)).ToString("U") +"<br>" + "Инициатор конференции: " + init.DispName + "<br>" + "В указанное время, для участия в конференции, просьба перейти по " + link + "<br><br>" + "<b><i>Данные для самостоятельного входа:<i><b><br> Адрес сервера: " + "https://" + MvcApplication.set.CobaCfgAddress + "/" + "<br>" + "Имя конференции: " + roomalias.alias + "<br>" + gpmail + "<br>" + "SIP-адрес: " + roomalias.alias + "@" + MvcApplication.set.CobaCfgAddress;
                    try
                    {
                       Sendmail(mail.Email, "UPD: " + meeting.Title, body, meeting, null);
                    }
                    catch (Exception e)
                    {

                        Debug.WriteLine(e.Message);
                        Debug.WriteLine(e.HResult);
                    }
                }
            }
            return Json(new[] { meeting }.ToDataSourceResult(request, ModelState));
        }
        public Task<ActionResult> Sendmail(string to, string subj, string body, MeetingViewModel meet, Event nEvent)
        {
            SmtpClient smtpClient = new SmtpClient(MvcApplication.set.AuthDnAddress, 25)
            {
                UseDefaultCredentials = false,
                EnableSsl = true,
                Credentials = new NetworkCredential(MvcApplication.set.SmtpLogin, MvcApplication.set.SmtpPassword),

                DeliveryMethod = SmtpDeliveryMethod.Network,
                Timeout = 20000
            };
            MailMessage mailMessage = new MailMessage()
            {
                Priority = MailPriority.High,
                From = new MailAddress("putin@kremlin.ru", "Планировщик системы видео-конференц-связи 'Сова'")
            };
            AlternateView alternateHtml = AlternateView.CreateAlternateViewFromString(body,
                                                                            new ContentType("text/html"));
            mailMessage.AlternateViews.Add(alternateHtml);
            mailMessage.To.Add(new MailAddress(to));
            mailMessage.Subject = subj;
            mailMessage.Body = body;

            //Event myEvent = new Event();
            
            mailMessage.AddEvent(nEvent);


            #region Method 1

            //StringBuilder str = new StringBuilder();
            //str.AppendLine("BEGIN:VCALENDAR");

            ////PRODID: identifier for the product that created the Calendar object
            //str.AppendLine("PRODID:-//ABC Company//Outlook MIMEDIR//EN");
            //str.AppendLine("VERSION:2.0");
            //str.AppendLine("METHOD:REQUEST");

            //str.AppendLine("BEGIN:VEVENT");

            //str.AppendLine(string.Format("DTSTART:{0:yyyyMMddTHHmmssZ}", meet.Start));
            //    //TimeZoneInfo.ConvertTimeToUtc("BeginTime").ToString("yyyyMMddTHHmmssZ")));
            //str.AppendLine(string.Format("DTSTAMP:{0:yyyyMMddTHHmmssZ}", DateTime.UtcNow));
            //str.AppendLine(string.Format("DTEND:{0:yyyyMMddTHHmmssZ}", meet.End));
            //    //TimeZoneInfo.ConvertTimeToUtc("EndTime").ToString("yyyyMMddTHHmmssZ")));
            //str.AppendLine(string.Format("LOCATION: {0}", "Система ВКС 'Сова'"));

            //// UID should be unique.
            //str.AppendLine(string.Format("UID:{0}", Guid.NewGuid()));
            //str.AppendLine(string.Format("DESCRIPTION:{0}", mailMessage.Body));
            //str.AppendLine(string.Format("X-ALT-DESC;FMTTYPE=text/html:{0}", mailMessage.Body));
            //str.AppendLine(string.Format("SUMMARY:{0}", mailMessage.Subject));

            //str.AppendLine("STATUS:CONFIRMED");
            //str.AppendLine("BEGIN:VALARM");
            //str.AppendLine("TRIGGER:-PT15M");
            //str.AppendLine("ACTION:Accept");
            //str.AppendLine("DESCRIPTION:Reminder");
            //str.AppendLine("X-MICROSOFT-CDO-BUSYSTATUS:BUSY");
            //str.AppendLine("END:VALARM");
            //str.AppendLine("END:VEVENT");

            //str.AppendLine(string.Format("ORGANIZER:MAILTO:{0}", mailMessage.From.Address));
            //str.AppendLine(string.Format("ATTENDEE;CN=\"{0}\";RSVP=TRUE:mailto:{1}", mailMessage.To[0].DisplayName,
            //    mailMessage.To[0].Address));

            //str.AppendLine("END:VCALENDAR");
            //ContentType ct = new ContentType("text/calendar");
            //ct.Parameters.Add("method", "REQUEST");
            //ct.Parameters.Add("name", "meeting.ics");
            //AlternateView avCal = AlternateView.CreateAlternateViewFromString(str.ToString(), ct);
            //mailMessage.AlternateViews.Add(avCal);

            #endregion


            smtpClient.Send(mailMessage);
            return null;
        }
        //public Task<ActionResult> RecordTask(string tasktitle, string taskapp, string stream_link, string pathflv, string comment, string acc_un, string accpass, DateTime task_start, DateTime task_end)
        //{
        //    ScheduledTasks st = new ScheduledTasks();
        //    Task t;
        //    t = st.CreateTask(tasktitle);
        //    t.ApplicationName = taskapp;
        //    t.Parameters = " -r " + "-q" + pathflv + " -o " + "\"" + stream_link + "\"";
        //    t.Comment = comment;
        //    t.SetAccountInformation(acc_un, accpass);
        //    t.IdleWaitMinutes = 10;
        //    TimeSpan worktime = task_end - task_start;
        //    Debug.WriteLine(worktime);
        //    t.MaxRunTime = new TimeSpan(worktime.Ticks);

        //    t.Priority = System.Diagnostics.ProcessPriorityClass.Idle;
        //    DateTime starttask = new DateTime();
        //    starttask = task_start + TimeSpan.FromHours(3);
        //    t.Triggers.Add(new RunOnceTrigger(starttask));
        //    t.Save();
        //    t.Close();
        //    st.Dispose();
        //    return null;
        //}
        private IEnumerable<AspNetUser> GetPB()
        {
            var data = GetPhBOw(User.Identity.Name);
            return data;
        }
        private IEnumerable<AspNetUser> GetAllPB()
        {
            aspnetdbEntities db = new aspnetdbEntities();
            var data = db.AspNetUsers.AsEnumerable();
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
                if (!String.IsNullOrEmpty(sel.UsersGroup))
                {
                    temp.Group = sel.UsersGroup;
                }
                if (String.IsNullOrEmpty(sel.UsersGroup))
                {
                    temp.Group = "Группа не назначена";
                }
                selrec.Add(temp);

            }

            return selrec;
        }
        public JsonResult GetAttendies([DataSourceRequest] DataSourceRequest request)
        {
            aspnetdbEntities db = new aspnetdbEntities();
            
            var data = db.AspNetUsers.AsEnumerable();
            return Json(data.ToDataSourceResult(request, o => new { id = o.Id, name = o.DispName }));
        }
    }
}