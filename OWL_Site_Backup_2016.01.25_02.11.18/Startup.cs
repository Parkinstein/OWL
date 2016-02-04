using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(OWL_Site.Startup))]
namespace OWL_Site
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
