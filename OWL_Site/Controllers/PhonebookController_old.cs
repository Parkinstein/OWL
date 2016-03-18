using System;
using System.Collections.Generic;
using System.Data.Entity.Migrations;
using System.Diagnostics;
using System.Linq;
using System.Web.Mvc;
using OWL_Site.Context;
using OWL_Site.Models;

namespace OWL_Site.Controllers
{
    public partial class PhonebookController_old : Controller
    {
        // GET: Phonebook view
        public ActionResult Phonebook()
        {
            return View();
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


        
        //Get Personal Phonebook
        public ActionResult Phonebook_Ajax()
        {
            aspnetdbEntities dbs = new aspnetdbEntities();
            ApplicationDbContext db = new ApplicationDbContext();
            List<ApplicationUser> selrec = new List<ApplicationUser>();
            IEnumerable<PrivatePhB> selectets = (dbs.PrivatePhBs.Where(m => m.OwSAN == User.Identity.Name));
            foreach (var sel in selectets)
            {
              Debug.WriteLine((db.Users.FirstOrDefault(m => m.Id == sel.IdREC).DispName));
              selrec.Add(db.Users.FirstOrDefault(m => m.Id == sel.IdREC));
            }
            return Json(new
            {
                data = selrec,
            }, JsonRequestBehavior.AllowGet);
        }
        // Add Phonebook records method
        public void Phonebook_Add(object[] pbrArray)
        {
            var allpriv = GetPhBOw(User.Identity.Name).AsQueryable();
            foreach (string pbr in pbrArray)
            {
                if (!allpriv.Any(m => m.Id == pbr))
                    AddToPrivat(pbr);
                else { ViewBag.Message = "Запись уже существует"; Debug.WriteLine("Запись уже существует"); }
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

        public List<AspNetUser> GetPhBOw(string Owname) //get phonebook
        {
            aspnetdbEntities db = new aspnetdbEntities();
            List<AspNetUser> selrec = new List<AspNetUser>();
            var selectets = db.PrivatePhBs.Where(m => m.OwSAN == Owname);
            foreach (var sel in selectets)
            {
                var srec = db.AspNetUsers.FirstOrDefault(m => m.Id == sel.IdREC);
                var temp = srec;
                if (!String.IsNullOrEmpty(sel.Group))
                {
                    if (temp != null) temp.Group = sel.Group;
                }
                if (String.IsNullOrEmpty(sel.Group))
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

        // Delete Phonebook records method
        public void DeleteFromPrivat(string ids)
        {
            
            DeleteRecFromDb(ids, User.Identity.Name);
            //ViewBag.DeletedRecs = String.Format("Удалено {0} записей", i);
        }
        public bool addUserToPrivat(string Owner, string IdRec, string Group)
        {
            aspnetdbEntities db = new aspnetdbEntities();

            PrivatePhB newrecpriv = new PrivatePhB();
            newrecpriv.OwSAN = Owner;
            newrecpriv.IdREC = IdRec;
            newrecpriv.Group = Group;
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