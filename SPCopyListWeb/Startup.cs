using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(SPCopyListWeb.Startup))]
namespace SPCopyListWeb
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
