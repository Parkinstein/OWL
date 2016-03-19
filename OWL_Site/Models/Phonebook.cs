using System.Collections.Generic;

namespace OWL_Site.Models
{
    public class Phonebook
    {
        public class PrivatePhBs
        {
            public int Id { get; set; }
            public string OwSAN { get; set; }
            public string IdREC { get; set; }
            public string UsersGroup { get; set; }
        }

        public class AspNetUsers
        {
            public string Id { get; set; }        
            public string Name { get; set; }
            public string Surname { get; set; }
            public string Position { get; set; }
            public string Tel_int { get; set; }
            public string Tel_ext { get; set; }
            public string Tel_mob { get; set; }
            public string Email { get; set; }
            public string Timezone { get; set; }
            public string Sip_addr { get; set; }
            public string H323_addr { get; set; }
        }
    }
}