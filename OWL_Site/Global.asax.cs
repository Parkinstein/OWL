using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Diagnostics;
using System.DirectoryServices;
using System.Linq;
using System.Text;
using System.Web.Mvc;
using System.Web.Optimization;
using System.Web.Routing;
using OWL_Site.Models;

namespace OWL_Site
{
    public class MvcApplication : System.Web.HttpApplication
    {
        protected void Application_Start()
        {
            AreaRegistration.RegisterAllAreas();
            FilterConfig.RegisterGlobalFilters(GlobalFilters.Filters);
            RouteConfig.RegisterRoutes(RouteTable.Routes);
            BundleConfig.RegisterBundles(BundleTable.Bundles);
            //SqlDependency.Start(ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString);
        

            Settings_Read();

            // GetPhonebookUsers();
            // Database.SetInitializer<ApplicationDbContext>(new ApplicationDbContext.DropCreateAlwaysInitializer());
        }
        protected void Application_End()
        {
            //SqlDependency.Stop(ConfigurationManager.ConnectionStrings["DefaultConnection"].ConnectionString);
        }

        public static Setting set ;
        public static string GetProperty(SearchResult searchResult, string PropertyName)
        {
            if (searchResult.Properties.Contains(PropertyName))
            {
                return searchResult.Properties[PropertyName][0].ToString();
            }
            else
            {
                return string.Empty;
            }
        }
        private string Win1251ToUTF8(string source)
        {
            Encoding utf8 = Encoding.GetEncoding("windows-1251");
            Encoding win1251 = Encoding.GetEncoding("utf-8");
            byte[] utf8Bytes = win1251.GetBytes(source);
            byte[] win1251Bytes = Encoding.Convert(win1251, utf8, utf8Bytes);
            source = win1251.GetString(win1251Bytes);
            return source;
        }
        public Setting Settings_Read()
        {
           var db = new aspnetdbEntities();
            set = db.Settings.FirstOrDefault();
            return set;
        }
        #region Get_local_Users_&_Compare
        public List<ApplicationUser> GetPhonebookUsers()
        {
            List<ApplicationUser> allreco = new List<ApplicationUser>();
            try
            {

                string domainPath = String.Concat(MvcApplication.set.AuthDnAddress, "/OU=", MvcApplication.set.OU, ",DC=rad,DC=lan,DC=local");//"dc0.rad.lan.local/OU=Pepux,DC=rad,DC=lan,DC=local";
                DirectoryEntry directoryEntry = new DirectoryEntry("LDAP://" + domainPath, MvcApplication.set.DnAdminUn, MvcApplication.set.DnAdminPass);
                DirectorySearcher dirSearcher = new DirectorySearcher(directoryEntry);
                dirSearcher.SearchScope = SearchScope.Subtree;
                dirSearcher.Filter = "(objectClass=user)";
                dirSearcher.PropertiesToLoad.Add("givenName");
                dirSearcher.PropertiesToLoad.Add("sn");
                dirSearcher.PropertiesToLoad.Add("title");
                dirSearcher.PropertiesToLoad.Add("telephoneNumber");
                dirSearcher.PropertiesToLoad.Add("sAMAccountName");
                dirSearcher.PropertiesToLoad.Add("displayName");
                dirSearcher.PropertiesToLoad.Add("email");
                dirSearcher.PropertiesToLoad.Add("ipPhone");
                SearchResultCollection resultCol = dirSearcher.FindAll();
                foreach (SearchResult resul in resultCol)
                {
                    ApplicationUser objSurveyUsers = new ApplicationUser();
                    objSurveyUsers.Name = GetProperty(resul, "givenName");//(String)resul.Properties["givenName"][0];
                    objSurveyUsers.Surname = GetProperty(resul, "sn"); //(String)resul.Properties["sn"][0];
                    objSurveyUsers.Tel_int = GetProperty(resul, "telephoneNumber"); //(String)resul.Properties["telephoneNumber"][0];
                    objSurveyUsers.Position = GetProperty(resul, "title"); //(String)resul.Properties["title"][0];
                    objSurveyUsers.Email = GetProperty(resul, "email"); //(String)resul.Properties["email"][0];
                    objSurveyUsers.Sammaccount = GetProperty(resul, "sAMAccountName"); //(String)resul.Properties["sAMAccountName"][0];
                    objSurveyUsers.DispName = GetProperty(resul, "displayName"); //(String)resul.Properties["displayName"][0];
                    objSurveyUsers.DispName = GetProperty(resul, "ipPhone");
                    allreco.Add(objSurveyUsers);
                }
                foreach (var user in allreco)
                {
                    Debug.WriteLine("Найден юзверь");
                    Debug.WriteLine(user.DispName);
                }
                //CompareUsers(allreco);
            }
            catch (Exception er)
            {
                Debug.WriteLine(er.HResult);
                Debug.WriteLine(er.Message);
            }
            return allreco;
        }
        

        public void CompareUsers(List<ApplicationUser> adusList)
        {
            var db = new aspnetdbEntities();
            var temp_list = new List<string>();
            var id_list = new List<string>();
            List<ApplicationUser> allr = new List<ApplicationUser>();
            var NameQuery =
                    from samaccountname in db.AspNetUsers
                    select samaccountname;
            if (NameQuery != null)
            {
                foreach (var customer in NameQuery)
                {
                    if ((!adusList.Exists(x => x.Sammaccount == customer.Sammaccount) /*&& !customer.location*/))
                    {
                        temp_list.Add(customer.Sammaccount);
                        id_list.Add(customer.Id);
                    }
                }
                foreach (var stroke in temp_list)
                {
                    var deleteUsers =
                        from samaccountname in db.AspNetUsers
                        where samaccountname.Sammaccount == stroke
                        select samaccountname;
                    
                    


                }
                //foreach (var ids in id_list)
                //{
                //    var deleteRecs =
                //        from Id in db.PrivatePhB
                //        where Id.IdREC == ids
                //        select Id;
                //    db.PrivatePhB.Remove(deleteRecs.First());
                //}
                //db.SaveChanges();

            }

            foreach (var adus in adusList)
            {
                if (!NameQuery.AsEnumerable().ToList().Exists(x => x.Sammaccount == adus.Sammaccount))
                {
                    
                }
            }


            foreach (var temp in temp_list)
            {
                Debug.WriteLine(temp);
            }
        }
        
        #endregion
    }
}
