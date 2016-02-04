using System.Collections.Generic;
using System.Runtime.Serialization;
using Newtonsoft.Json;

namespace OWL_Site.Models
{
    public class ActiveConfsModel
    {
        [DataContract(Name = "")]
        [JsonObject(MemberSerialization.OptIn)]
        public class ResponseParent
        {
            [DataMember(Order = 1), Newtonsoft.Json.JsonProperty]
            public Meta metas { get; set; }

            [DataMember(Order = 2), Newtonsoft.Json.JsonProperty("objects")]

            public List<AConfs> obj { get; set; }
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
        public class AConfs
        {
            [DataMember(Order = 1), Newtonsoft.Json.JsonProperty]
            public string id { get; set; }

            [DataMember(Order = 2), Newtonsoft.Json.JsonProperty]
            public bool is_locked { get; set; }

            [DataMember(Order = 3), Newtonsoft.Json.JsonProperty]
            public string name { get; set; }

            [DataMember(Order = 4), Newtonsoft.Json.JsonProperty]
            public string resource_uri { get; set; }

            [DataMember(Order = 5), Newtonsoft.Json.JsonProperty]
            public string service_type { get; set; }

            [DataMember(Order = 6), Newtonsoft.Json.JsonProperty]
            public string start_time { get; set; }
            [DataMember(Order = 7), Newtonsoft.Json.JsonProperty]
            public string start_time2 { get; set; }

            [DataMember(Order = 8), Newtonsoft.Json.JsonProperty]
            public string tag { get; set; }
            [DataMember(Order = 9), Newtonsoft.Json.JsonProperty]
            public string lock_path { get; set; }

        }
    }

    class ActiveConfsModelImpl : ActiveConfsModel
    {
    }
}