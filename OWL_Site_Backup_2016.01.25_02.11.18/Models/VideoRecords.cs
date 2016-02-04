using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace OWL_Site.Models
{
    public class VideoRecords
    {
        public class Object
        {
            public int ID { get; set; }
            public string Conf { get; set; }
            public string PName { get; set; }
            public DateTime Tstart { get; set; }
            public DateTime Tfinish { get; set; }
            public string Link { get; set; }
        }
    }
}