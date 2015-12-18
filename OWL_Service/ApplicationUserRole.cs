using Microsoft.AspNet.Identity.EntityFramework;

namespace OWL_Service
{
    public class ApplicationUserRole : IdentityUserRole
    {
        public ApplicationUserRole()
            : base()
        { }

        public ApplicationRole Role { get; set; }
    }
}