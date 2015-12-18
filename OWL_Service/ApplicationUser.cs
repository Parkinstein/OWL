using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Security.Claims;
using System.Threading.Tasks;
using Microsoft.AspNet.Identity;
using Microsoft.AspNet.Identity.EntityFramework;

namespace OWL_Service
{
    // You can add profile data for the user by adding more properties to your ApplicationUser class, please visit http://go.microsoft.com/fwlink/?LinkID=317594 to learn more.
    public class ApplicationUser : IdentityUser
    {
        public async Task<ClaimsIdentity> GenerateUserIdentityAsync(UserManager<ApplicationUser> manager)
        {
            // Note the authenticationType must match the one defined in CookieAuthenticationOptions.AuthenticationType
            var userIdentity = await manager.CreateIdentityAsync(this, DefaultAuthenticationTypes.ApplicationCookie);
            // Add custom user claims here
            return userIdentity;
        }
        [Display(Name = "Id")]
        
        public string Name { get; set; }
        [Display(Name = "Surname")]
        public string Surname { get; set; }
        [Display(Name = "DispName")]
        public string DispName { get; set; }
        [Display(Name = "Sammaccount")]
        public string Sammaccount { get; set; }
        [Display(Name = "Position")]
        public string Position { get; set; }
        [Display(Name = "Tel_int")]
        public string Tel_int { get; set; }
        [Display(Name = "Tel_ext")]
        public string Tel_ext { get; set; }
        [Display(Name = "Tel_mob")]
        public string Tel_mob { get; set; }
        [Display(Name = "Timezone")]
        public string Timezone { get; set; }
        [Display(Name = "Sip_addr")]
        public string Sip_addr { get; set; }
        [Display(Name = "H323_addr")]
        public string H323_addr { get; set; }
        [Display(Name = "Group")]
        public string Group { get; set; }

        public ICollection<ApplicationUserRole> UserRoles { get; set; }
    }
}
