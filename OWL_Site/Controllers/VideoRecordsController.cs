using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Web.Mvc;
using MeetingRequest;
using MySql.Data.MySqlClient;
using OWL_Site.Models;
using Renci.SshNet;

namespace OWL_Site.Controllers
{
    public class VideoRecordsController : Controller
    {
        // GET: Records
        public ActionResult VideoRecords()
        {
            return View();
        }

        public ActionResult VideoRecords_Ajax()
        {
            IEnumerable<VideoRecords.Object> result = GetRec();
            return Json(new
            {
               data = result,
            }, JsonRequestBehavior.AllowGet);
        }


        public IEnumerable<VideoRecords.Object> GetRec()
        {
            //if (!User.IsInRole("Admin"))
            //{
            //    return Videorecs(User.Identity.Name);
            //}
            //else
            //{
            //    return Videorecs("");
            //}
            return Videorecs("");
        }
        // Delete Records method
        public void VideoRecords_Delete(object[] pbrArray)
        {
            foreach (int pbr in pbrArray)
            {
                DeleteFromRecords(pbr);
            }
        }

        // Delete Phonebook records method
        public void DeleteFromRecords(int ids)
        {
            DeleteRecordsFromDb(ids);
        }

        public List<VideoRecords.Object> Videorecs(string val)
        {
            //MySql.Data.MySqlClient.MySqlConnection conn;
            MySqlDataAdapter daVrec;
            DataSet dsVrec;
            string myConnectionString;
            List<VideoRecords.Object> all_recs = new List<VideoRecords.Object>();

            myConnectionString = "server=" + MvcApplication.set.CobaRecordsAddress + ";uid=" + MvcApplication.set.CobaRecLogin + ";" +
                "pwd=" + MvcApplication.set.CobaRecPass + ";database=" + MvcApplication.set.CobaRecBdName + ";Convert Zero Datetime=True";
            if (val != "")
            {
                try
                {
                    string sql = "SELECT * FROM records WHERE Conf " + " = \"" + val + "\""; //
                    Debug.WriteLine(sql);
                    daVrec = new MySqlDataAdapter(sql, myConnectionString);
                    MySqlCommandBuilder cb = new MySqlCommandBuilder(daVrec);
                    dsVrec = new DataSet();
                    daVrec.Fill(dsVrec, "records");
                    foreach (DataRow dr in dsVrec.Tables["records"].Rows)
                    {
                        all_recs.Add(new VideoRecords.Object { ID = Int32.Parse(dr["ID"].ToString()), Conf = Convert.ToString(dr["Conf"]), PName = Convert.ToString(dr["PName"]), Tstart = Convert.ToString(dr["Tstart"]), Tfinish = Convert.ToString(dr["Tfinish"]), Link = Convert.ToString(dr["Link"]) });
                    }
                    return all_recs;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex.Message.ToString());
                }
            }
            try
            {

                string sql = "SELECT * FROM records";
                daVrec = new MySqlDataAdapter(sql, myConnectionString);
                MySqlCommandBuilder cb = new MySqlCommandBuilder(daVrec);
                dsVrec = new DataSet();
                daVrec.Fill(dsVrec, "records");
                foreach (DataRow dr in dsVrec.Tables["records"].Rows)
                {
                    all_recs.Add(new VideoRecords.Object { ID = Int32.Parse(dr["ID"].ToString()), Conf = Convert.ToString(dr["Conf"]), PName = Convert.ToString(dr["PName"]), Tstart = Convert.ToString(dr["Tstart"]), Tfinish = Convert.ToString(dr["Tfinish"]), Link = Convert.ToString(dr["Link"]) });
                }
                return all_recs;

            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message.ToString());
            }


            return all_recs;
        }

        public bool DeleteRecordsFromDb(int id)
        {
            var myConnectionString = "server=" + MvcApplication.set.CobaRecordsAddress + ";uid=" + MvcApplication.set.CobaRecLogin + ";" +
                "pwd=" + MvcApplication.set.CobaRecPass + ";database=" + MvcApplication.set.CobaRecBdName + ";Convert Zero Datetime=True";
            MySqlConnection conn = new MySqlConnection(myConnectionString);
            try
            {
                MySqlCommand cmd = new MySqlCommand("SELECT Link FROM records WHERE ID = " + id + "", conn);
                conn.Open();
                Debug.WriteLine(cmd.CommandText);
                var mysqlAdp = new MySqlDataAdapter(cmd);
                var mysqlDS = new DataSet();
                mysqlAdp.Fill(mysqlDS, "Links");
                foreach (DataRow dr in mysqlDS.Tables["Links"].Rows)
                {
                    var link = Convert.ToString(dr["Link"]);
                    Debug.WriteLine(link);
                    int found = link.IndexOf("/records");
                    Debug.WriteLine(link.Substring(found));
                    cmd = new MySqlCommand("DELETE FROM records WHERE ID = " + id + "", conn);
                    Debug.WriteLine(cmd.CommandText);
                    cmd.ExecuteNonQuery();
                    // SSH procedure to delete records files
                    var PasswordConnection = new PasswordAuthenticationMethod(MvcApplication.set.CobaRecLogin, MvcApplication.set.CobaRecPass);
                    var KeyboardInteractive = new KeyboardInteractiveAuthenticationMethod(MvcApplication.set.CobaRecLogin);

                    var connectionInfo = new ConnectionInfo(MvcApplication.set.CobaRecordsAddress, 22, MvcApplication.set.CobaRecLogin, PasswordConnection, KeyboardInteractive);
                    using (SshClient ssh = new SshClient(connectionInfo))
                    {
                        ssh.Connect();
                        var command = ssh.RunCommand("rm -f /home/rerih/public_html" + link.Substring(found));
                        Debug.WriteLine(command.CommandText);
                        ssh.Disconnect();
                    }

                }
                conn.Close();
                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                return false;
            }
        }
    }
}