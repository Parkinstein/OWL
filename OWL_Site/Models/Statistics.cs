using System.Collections.Generic;

namespace OWL_Site.Models
{
    public class VmrStatsResponse
    {
        public Meta meta { get; set; }
        public List<VmrStats> objects { get; set; }
    }

    public class Meta
    {
        public int limit { get; set; }
        public object next { get; set; }
        public int offset { get; set; }
        public object previous { get; set; }
        public int total_count { get; set; }
    }

    public class VmrStats
    {
        public string id { get; set; }
        public string name { get; set; }
        public string start_time { get; set; }
        public string start_time2 { get; set; }
        public string end_time { get; set; }
        public string end_time2 { get; set; }
        public string duration { get; set; }
        public int participant_count { get; set; }
        public string[] participants { get; set; }
    }

    public class ParticipantStatsResponse
    {
        public Meta meta { get; set; }
        public List<ParticipantStats> objects { get; set; }
    }

    public class ParticipantStats
    {
        public string conference_name { get; set; }
        public string role { get; set; }
        public string local_alias { get; set; }
        public string remote_alias { get; set; }
        public string start_time { get; set; }
        public string start_time2 { get; set; }
        public string end_time { get; set; }
        public string end_time2 { get; set; }
        public string duration { get; set; }
        public string display_name { get; set; }
        public string protocol { get; set; }
        public string vendor { get; set; }
        public string remote_address { get; set; }
        public string remote_port { get; set; }
        public string disconnect_reason { get; set; }
        public List<Mediastreams> media_streams { get; set; }

        public string vid_stream_type { get; set; }
        public string vid_tx_bitrate { get; set; }
        public string vid_tx_codec { get; set; }
        public string vid_tx_resolution { get; set; }
        public string vid_rx_bitrate { get; set; }
        public string vid_rx_codec { get; set; }
        public string vid_rx_resolution { get; set; }

        public string aud_stream_type { get; set; }
        public string aud_tx_bitrate { get; set; }
        public string aud_tx_codec { get; set; }
        public string aud_tx_resolution { get; set; }
        public string aud_rx_bitrate { get; set; }
        public string aud_rx_codec { get; set; }
        public string aud_rx_resolution { get; set; }

        public string pre_stream_type { get; set; }
        public string pre_tx_bitrate { get; set; }
        public string pre_tx_codec { get; set; }
        public string pre_tx_resolution { get; set; }
        public string pre_rx_bitrate { get; set; }
        public string pre_rx_codec { get; set; }
        public string pre_rx_resolution { get; set; }
    }

    public class Mediastreams
    {
        public string stream_type { get; set; }

        public string tx_bitrate { get; set; }
        public string tx_codec { get; set; }
        public string tx_resolution { get; set; }

        public string rx_bitrate { get; set; }
        public string rx_codec { get; set; }
        public string rx_resolution { get; set; }
    }
}