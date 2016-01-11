using System.Collections.Generic;
using System.Runtime.Serialization;
using Newtonsoft.Json;

namespace OWL_Service
{
    public class AllVMRS
    {
        [DataContract(Name = "allvmr")]

        [JsonObject(MemberSerialization.OptIn)]

        public class VmrParent
        {
            [DataMember(Order = 1), Newtonsoft.Json.JsonProperty]
            public Meta metas { get; set; }

            [DataMember(Order = 2), Newtonsoft.Json.JsonProperty("objects")]

            public List<AllVmrs> obj { get; set; }
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
        public partial class AllVmrs
        {
            [DataMember(Order = 1), Newtonsoft.Json.JsonProperty]
            public List<Aliasess> aliases { get; set; }

            [DataMember(Order = 2), Newtonsoft.Json.JsonProperty]
            public bool allow_guests { get; set; }

            [DataMember(Order = 3), Newtonsoft.Json.JsonProperty]
            public List<AutoPartis> automatic_participants { get; set; }

            [DataMember(Order = 4), Newtonsoft.Json.JsonProperty]
            public string description { get; set; }

            [DataMember(Order = 5), Newtonsoft.Json.JsonProperty]
            public bool force_presenter_into_main { get; set; }

            [DataMember(Order = 6), Newtonsoft.Json.JsonProperty]
            public string guest_pin { get; set; }
            [DataMember(Order = 7), Newtonsoft.Json.JsonProperty]
            public string guest_view { get; set; }

            [DataMember(Order = 8), Newtonsoft.Json.JsonProperty]
            public string host_view { get; set; }
            [DataMember(Order = 9), Newtonsoft.Json.JsonProperty]
            public int id { get; set; }
            [DataMember(Order = 10), Newtonsoft.Json.JsonProperty]
            public ivr_theme ivr_theme { get; set; }
            [DataMember(Order = 11), Newtonsoft.Json.JsonProperty]
            public string max_callrate_in { get; set; }
            [DataMember(Order = 12), Newtonsoft.Json.JsonProperty]
            public string max_callrate_out { get; set; }
            [DataMember(Order = 13), Newtonsoft.Json.JsonProperty]
            public string name { get; set; }
            [DataMember(Order = 14), Newtonsoft.Json.JsonProperty]
            public string participant_limit { get; set; }
            [DataMember(Order = 15), Newtonsoft.Json.JsonProperty]
            public string pin { get; set; }
            [DataMember(Order = 16), Newtonsoft.Json.JsonProperty]
            public string resource_uri { get; set; }
            [DataMember(Order = 17), Newtonsoft.Json.JsonProperty]
            public string service_type { get; set; }
            [DataMember(Order = 18), Newtonsoft.Json.JsonProperty]
            public string sync_tag { get; set; }
            [DataMember(Order = 19), Newtonsoft.Json.JsonProperty]
            public string tag { get; set; }
        }
        [DataContract]
        [Newtonsoft.Json.JsonObject(MemberSerialization = Newtonsoft.Json.MemberSerialization.OptIn)]
        public partial class Aliasess
        {
            [DataMember(Order = 1), Newtonsoft.Json.JsonProperty]
            public string alias { get; set; }
            [DataMember(Order = 2), Newtonsoft.Json.JsonProperty]
            public string conference { get; set; }
            [DataMember(Order = 3), Newtonsoft.Json.JsonProperty]
            public string description { get; set; }
            [DataMember(Order = 4), Newtonsoft.Json.JsonProperty]
            public int id { get; set; }
        }
        [DataContract]
        [Newtonsoft.Json.JsonObject(MemberSerialization = Newtonsoft.Json.MemberSerialization.Fields)]
        public  class ivr_theme
        {
            [DataMember(Order = 1), Newtonsoft.Json.JsonProperty]
            public int id { get; set; }
            [DataMember(Order = 2), Newtonsoft.Json.JsonProperty]
            public string name { get; set; }
            [DataMember(Order = 3), Newtonsoft.Json.JsonProperty]
            public string uuid { get; set; }
        }
        [DataContract]
        [Newtonsoft.Json.JsonObject(MemberSerialization = Newtonsoft.Json.MemberSerialization.OptIn)]
        public partial class AutoPartis
        {
            [DataMember(Order = 1), Newtonsoft.Json.JsonProperty]
            public string alias { get; set; }
            [DataMember(Order = 1), Newtonsoft.Json.JsonProperty]
            public string description { get; set; }
            [DataMember(Order = 1), Newtonsoft.Json.JsonProperty]
            public string dtmf_sequence { get; set; }
            [DataMember(Order = 1), Newtonsoft.Json.JsonProperty]
            public int id { get; set; }
            [DataMember(Order = 1), Newtonsoft.Json.JsonProperty]
            public string protocol { get; set; }
            [DataMember(Order = 1), Newtonsoft.Json.JsonProperty]
            public string role { get; set; }
            [DataMember(Order = 1), Newtonsoft.Json.JsonProperty]
            public bool streaming { get; set; }
        }
    }
}
