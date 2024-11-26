using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net.Mail;
using System.Web;
using System.Web.Security;
using System.Web.SessionState;
using System.Text;
using Microsoft.IdentityModel.Tokens;
using Microsoft.IdentityModel.Web;
using Microsoft.IdentityModel.Web.Configuration;

namespace SSO
{
    public class Global : System.Web.HttpApplication
    {

        protected void Application_Start(object sender, EventArgs e)
        {
            FederatedAuthentication.ServiceConfigurationCreated += OnServiceConfigurationCreated;
        }
        private void OnServiceConfigurationCreated(object sender, ServiceConfigurationCreatedEventArgs e)
        {
            List<CookieTransform> sessionTransforms = new List<CookieTransform>
            (
                new CookieTransform[]
                {
                    new DeflateCookieTransform(),
                    new CustomCookieTransform(),
                    new CustomCookieTransform()
                }
            );
            SessionSecurityTokenHandler sessionHandler = new SessionSecurityTokenHandler(sessionTransforms.AsReadOnly());
            e.ServiceConfiguration.SecurityTokenHandlers.AddOrReplace(sessionHandler);
        }



        protected void Session_Start(object sender, EventArgs e)
        {
            
        }

        protected void Application_BeginRequest(object sender, EventArgs e)
        {

        }

        protected void Application_AuthenticateRequest(object sender, EventArgs e)
        {

        }

        protected void Application_Error(object sender, EventArgs e)
        {

        }

        protected void Session_End(object sender, EventArgs e)
        {

        }

        protected void Application_End(object sender, EventArgs e)
        {

        }
    }
    public class CustomCookieTransform : CookieTransform
    {
        public override byte[] Decode(byte[] encoded)
        {
            string encodedStr = Encoding.UTF8.GetString(encoded, 0, encoded.Length);
            byte[] decoded = System.Web.Security.MachineKey.Decode(encodedStr, System.Web.Security.MachineKeyProtection.All);
            return decoded;
        }

        public override byte[] Encode(byte[] value)
        {
            string encodedStr = System.Web.Security.MachineKey.Encode(value, System.Web.Security.MachineKeyProtection.All);
            byte[] encoded = Encoding.UTF8.GetBytes(encodedStr);
            return encoded;
        }
    }

}