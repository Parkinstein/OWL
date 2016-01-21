using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Net;
using System.Text;
using System.Web.Mvc;
using Newtonsoft.Json;
using OpenAccessRuntime;
using OWL_Site;
using OWL_Site.Models;

namespace PepuxFront.Controllers
{
    public class StatisticsController : Controller
    {
        public VmrStatsResponse HistoryVmrFull;
        public List<VmrStats> HistoryVmrData;
        public ParticipantStatsResponse HistoryParticipantsFull;
        public List<ParticipantStats> HistoryParticipantsData;

        public ActionResult Statistics()
        {
            return View();
        }

        public ActionResult GetHistoryVMR()
        {
            IEnumerable<VmrStats> result = GetHistoryVmrData();
            return Json(new
            {
                data = result
            }, JsonRequestBehavior.AllowGet);
        }

        public ActionResult GetHistoryParticipants(string id)
        {
            IEnumerable<ParticipantStats> result = GetHistoryParticipantsData(id);
            return Json(new
            {
                data = result
            }, JsonRequestBehavior.AllowGet);
        }
        


        private List<VmrStats> GetHistoryVmrData()
        {
            try
            {
                HistoryVmrData = new List<VmrStats>();
                Uri historyapi = new Uri("https://" + MvcApplication.set.CobaMngAddress + "/api/admin/history/v1/conference/?limit=100");
                WebClient client = new WebClient();
                client.Credentials = new NetworkCredential(MvcApplication.set.CobaMngLogin, MvcApplication.set.CobaMngPass);
                client.Headers.Add("auth", MvcApplication.set.CobaMngLogin+","+MvcApplication.set.CobaMngPass);
                client.Headers.Add("veryfy", "False");
                ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
                HistoryVmrFull = JsonConvert.DeserializeObject<VmrStatsResponse>(Win1251ToUTF8(client.DownloadString(historyapi)));
                HistoryVmrData = HistoryVmrFull.objects;
                foreach (var historyRecords in HistoryVmrData)
                {
                    historyRecords.start_time2 =
                        (DateTime.Parse(historyRecords.start_time) + TimeSpan.FromHours(3)).ToString("dd-MMM-yyyy  HH:mm:ss");
                    historyRecords.end_time2 =
                        (DateTime.Parse(historyRecords.end_time) + TimeSpan.FromHours(3)).ToString("dd-MMM-yyyy  HH:mm:ss");
                }
                return HistoryVmrData;
            }
            catch (Exception errException)
            {
                Debug.WriteLine(errException.Message);
            }
            return HistoryVmrData;
        }

        private List<ParticipantStats> GetHistoryParticipantsData(string id)
        {
            try
            {
                HistoryParticipantsData = new List<ParticipantStats>();
                Uri participantsapi =
                    new Uri("https://" + MvcApplication.set.CobaMngAddress + "/api/admin/history/v1/participant/?conference=" + id);
                WebClient client = new WebClient();
                client.Credentials = new NetworkCredential(MvcApplication.set.CobaMngLogin , MvcApplication.set.CobaMngPass);
                client.Headers.Add("auth", MvcApplication.set.CobaMngLogin + "," + MvcApplication.set.CobaMngPass);
                client.Headers.Add("veryfy", "False");
                ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
                HistoryParticipantsFull =
                    JsonConvert.DeserializeObject<ParticipantStatsResponse>(
                        Win1251ToUTF8(client.DownloadString(participantsapi)));
                HistoryParticipantsData = HistoryParticipantsFull.objects;
                foreach (var participantsRecords in HistoryParticipantsData)
                {
                    participantsRecords.start_time2 =
                        (DateTime.Parse(participantsRecords.start_time) + TimeSpan.FromHours(3)).ToString(
                            "dd-MMM-yyyy  HH:mm:ss");
                    participantsRecords.end_time2 =
                        (DateTime.Parse(participantsRecords.end_time) + TimeSpan.FromHours(3)).ToString(
                            "dd-MMM-yyyy  HH:mm:ss");
                    foreach (var mediastreamsRecords in participantsRecords.media_streams)
                    {
                        if (mediastreamsRecords.stream_type == "video")
                        {
                            participantsRecords.vid_stream_type = "video";
                            participantsRecords.vid_tx_codec = mediastreamsRecords.tx_codec;
                            participantsRecords.vid_tx_bitrate = mediastreamsRecords.tx_bitrate;
                            participantsRecords.vid_tx_resolution = mediastreamsRecords.tx_resolution;
                            participantsRecords.vid_rx_codec = mediastreamsRecords.rx_codec;
                            participantsRecords.vid_rx_bitrate = mediastreamsRecords.rx_bitrate;
                            participantsRecords.vid_rx_resolution = mediastreamsRecords.rx_resolution;

                        }
                        else if (mediastreamsRecords.stream_type == "audio")
                        {
                            participantsRecords.aud_stream_type = "audio";
                            participantsRecords.aud_tx_codec = mediastreamsRecords.tx_codec;
                            participantsRecords.aud_tx_bitrate = mediastreamsRecords.tx_bitrate;
                            participantsRecords.aud_tx_resolution = mediastreamsRecords.tx_resolution;
                            participantsRecords.aud_rx_codec = mediastreamsRecords.rx_codec;
                            participantsRecords.aud_rx_bitrate = mediastreamsRecords.rx_bitrate;
                            participantsRecords.aud_rx_resolution = mediastreamsRecords.rx_resolution;

                        }
                        else if (mediastreamsRecords.stream_type == "presentation")
                        {
                            participantsRecords.pre_stream_type = "presentation";
                            participantsRecords.pre_tx_codec = mediastreamsRecords.tx_codec;
                            participantsRecords.pre_tx_bitrate = mediastreamsRecords.tx_bitrate;
                            participantsRecords.pre_tx_resolution = mediastreamsRecords.tx_resolution;
                            participantsRecords.pre_rx_codec = mediastreamsRecords.rx_codec;
                            participantsRecords.pre_rx_bitrate = mediastreamsRecords.rx_bitrate;
                            participantsRecords.pre_rx_resolution = mediastreamsRecords.rx_resolution;
                        }
                    }
                }
                return HistoryParticipantsData;
            }
            catch (Exception errException)
            {
                Debug.WriteLine(errException.Message);
            }
            return HistoryParticipantsData;
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
    }
}