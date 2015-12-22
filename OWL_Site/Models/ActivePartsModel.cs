using System.Collections.Generic;
using System.Runtime.Serialization;
using Newtonsoft.Json;

namespace OWL_Site.Models
{
    public class ActivePartsModel
    {
        [DataContract(Name = "")]
        [JsonObject(MemberSerialization.OptIn)]
        public class ResponseParent
        {
            [DataMember(Order = 1), Newtonsoft.Json.JsonProperty]
            public Meta metas { get; set; }

            [DataMember(Order = 2), Newtonsoft.Json.JsonProperty("objects")]

            public List<AParts> obj { get; set; }
        }

        [DataContract(Name = "meta")]
        [JsonObject(MemberSerialization.OptOut)]
        public class Meta
        {
            [DataMember(Name = "limit")]
            public int limit { get; set; }

            [DataMember(Name = "next")]
            public object next { get; set; }

            [DataMember(Name = "offset")]
            public int offset { get; set; }

            [DataMember(Name = "previous")]
            public object previous { get; set; }

            [DataMember(Name = "total_count")]
            public int total_count { get; set; }
        }

        [DataContract]
        [Newtonsoft.Json.JsonObject(MemberSerialization = Newtonsoft.Json.MemberSerialization.OptIn)]
        public class AParts
        {
            [DataMember(Name = "bandwidth"), Newtonsoft.Json.JsonProperty]
            public int bandwidth { get; set; }

            [DataMember(Name = "call_direction"), Newtonsoft.Json.JsonProperty]
            public string call_direction { get; set; }

            [DataMember(Name = "call_uuid"), Newtonsoft.Json.JsonProperty]
            public string call_uuid { get; set; }

            [DataMember(Name = "conference"), Newtonsoft.Json.JsonProperty]
            public string conference { get; set; }

            [DataMember(Name = "connect_time"), Newtonsoft.Json.JsonProperty]
            public string connect_time { get; set; }

            [DataMember(Name = "destination_alias"), Newtonsoft.Json.JsonProperty]
            public string destination_alias { get; set; }

            [DataMember(Name = "display_name"), Newtonsoft.Json.JsonProperty]
            public string display_name { get; set; }

            [DataMember(Name = "encryption"), Newtonsoft.Json.JsonProperty]
            public string encryption { get; set; }

            [DataMember(Name = "has_media"), Newtonsoft.Json.JsonProperty]
            public bool has_media { get; set; }

            [DataMember(Name = "id"), Newtonsoft.Json.JsonProperty]
            public string id { get; set; }

            [DataMember(Name = "is_muted"), Newtonsoft.Json.JsonProperty]
            public bool is_muted { get; set; }

            [DataMember(Name = "is_on_hold"), Newtonsoft.Json.JsonProperty]
            public bool is_on_hold { get; set; }

            [DataMember(Name = "is_presentation_supported"), Newtonsoft.Json.JsonProperty]
            public bool is_presentation_supported { get; set; }

            [DataMember(Name = "is_presenting"), Newtonsoft.Json.JsonProperty]
            public bool is_presenting { get; set; }

            [DataMember(Name = "is_streaming"), Newtonsoft.Json.JsonProperty]
            public bool is_streaming { get; set; }

            [DataMember(Name = "license_count"), Newtonsoft.Json.JsonProperty]
            public int license_count { get; set; }

            [DataMember(Name = "media_node"), Newtonsoft.Json.JsonProperty]
            public string media_node { get; set; }

            [DataMember(Name = "parent_id"), Newtonsoft.Json.JsonProperty]
            public string parent_id { get; set; }

            [DataMember(Name = "participant_alias"), Newtonsoft.Json.JsonProperty]
            public string participant_alias { get; set; }

            [DataMember(Name = "protocol"), Newtonsoft.Json.JsonProperty]
            public string protocol { get; set; }

            [DataMember(Name = "remote_address"), Newtonsoft.Json.JsonProperty]
            public string remote_address { get; set; }

            [DataMember(Name = "remote_port"), Newtonsoft.Json.JsonProperty]
            public int remote_port { get; set; }

            [DataMember(Name = "resource_uri"), Newtonsoft.Json.JsonProperty]
            public string resource_uri { get; set; }

            [DataMember(Name = "role"), Newtonsoft.Json.JsonProperty]
            public string role { get; set; }

            [DataMember(Name = "service_tag"), Newtonsoft.Json.JsonProperty]
            public string service_tag { get; set; }

            [DataMember(Name = "service_type"), Newtonsoft.Json.JsonProperty]
            public string service_type { get; set; }

            [DataMember(Name = "signalling_node"), Newtonsoft.Json.JsonProperty]
            public string signalling_node { get; set; }

            [DataMember(Name = "source_alias"), Newtonsoft.Json.JsonProperty]
            public string source_alias { get; set; }

            [DataMember(Name = "system_location"), Newtonsoft.Json.JsonProperty]
            public string system_location { get; set; }

            [DataMember(Name = "vendor"), Newtonsoft.Json.JsonProperty]
            public string vendor { get; set; }
        }
    }
}