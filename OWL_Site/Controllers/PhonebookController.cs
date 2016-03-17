using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.Entity.Migrations;
using System.Data.Entity.SqlServer;
using System.Diagnostics;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using OWL_Site.Context;
using OWL_Site.Models;
using DataTables;

namespace OWL_Site.Controllers
{
    public class PhonebookController : Controller
    {
        // GET: Phonebook view
        public ActionResult Phonebook()
        {
            return View();
        }

        //Get Personal Phonebook
        public ActionResult Phonebook_Ajax()
        {
            var request = HttpContext.Request.Form;
            using (var db = new Database("sqlserver", ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString))
            {
                var personalPhonebook = new Editor(db, "PrivatePhBs", "Id")
                    .Model<Phonebook>()
                    .Field(new Field("PrivatePhBs.UsersGroup").Validator(Validation.MaxLen(50)).Xss(false))
                    .Where("OwSAN", User.Identity.Name)
                    .LeftJoin("AspNetUsers", "AspNetUsers.Id", "=", "PrivatePhBs.IdREC")
                    .Process(request)
                    .Data();
                return Json(personalPhonebook);
            }
        }

        public string Decoder(string input)
        {
            Debug.WriteLine(input);
            Debug.WriteLine(HttpUtility.HtmlDecode(input));
            return HttpUtility.HtmlDecode(input);
        }
        //Get Full Phonebook
        public ActionResult PhonebookAll_Ajax()
        {
            ApplicationDbContext cont = new ApplicationDbContext();
            List<ApplicationUser> users = cont.Users.ToList();
            return Json(new
            {
                data = users
            }, JsonRequestBehavior.AllowGet);
        }


        // Add Phonebook records method
        public void Phonebook_Add(object[] pbrArray)
        {
            var allpriv = GetPhBOw(User.Identity.Name).AsQueryable();
            if (pbrArray != null)
            {
                foreach (string pbr in pbrArray)
                {
                    if (!allpriv.Any(m => m.Id == pbr))
                        AddToPrivat(pbr);
                }
            }
            else
            {
                ViewBag.Message = "Запись уже существует";
                Debug.WriteLine("Запись уже существует");
            }
        }
        
        // Check already created personal records
        public List<AspNetUser> GetPhBOw(string Owname)
        {
            aspnetdbEntities db = new aspnetdbEntities();
            List<AspNetUser> selrec = new List<AspNetUser>();
            var selectets = db.PrivatePhBs.Where(m => m.OwSAN == Owname);
            foreach (var sel in selectets)
            {
                var srec = db.AspNetUsers.FirstOrDefault(m => m.Id == sel.IdREC);
                var temp = srec;
                if (!String.IsNullOrEmpty(sel.UsersGroup))
                {
                    if (temp != null) temp.Group = sel.UsersGroup;
                }
                if (String.IsNullOrEmpty(sel.UsersGroup))
                {
                    if (temp != null) temp.Group = "Группа не назначена";
                }
                selrec.Add(temp);

            }
            return selrec;
        }

        // Add records to private phonebook
        public void AddToPrivat(string ids)
        {
            bool result = addUserToPrivat(User.Identity.Name, ids, null);
        }

        public bool addUserToPrivat(string Owner, string IdRec, string UsersGroup)
        {
            aspnetdbEntities db = new aspnetdbEntities();

            PrivatePhB newrecpriv = new PrivatePhB();
            newrecpriv.OwSAN = Owner;
            newrecpriv.IdREC = IdRec;
            newrecpriv.UsersGroup = UsersGroup;
            db.PrivatePhBs.AddOrUpdate(newrecpriv);

            try
            {
                db.SaveChanges();
                return true;
            }
            catch(Exception ex)
            {
                Debug.WriteLine(ex.InnerException);
                return false;
            }

        }
        
        // Delete Phonebook records method
        public void Phonebook_Delete(object[] pbrArray)
        {
            if (pbrArray != null)
            {
                foreach (string pbr in pbrArray)
                {
                    DeleteFromPrivat(pbr);
                }
            }
            
        }

        // Delete Phonebook records method
        public void DeleteFromPrivat(string ids)
        {
            DeleteRecFromDb(ids, User.Identity.Name);
        }

        public bool DeleteRecFromDb(string id, string ownm)
        {
            aspnetdbEntities db = new aspnetdbEntities();
            var deleteDetails =
            from prop in db.PrivatePhBs
            where prop.IdREC == id && prop.OwSAN == ownm
            select prop;
            foreach (var detal in deleteDetails)
            {
                db.PrivatePhBs.Remove(detal);
            }

            try
            {
                db.SaveChanges();
                return true;
            }
            catch (Exception e)
            {
                Debug.WriteLine(e);
                return false;
            }
        }
    }
}