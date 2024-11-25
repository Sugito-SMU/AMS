using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Threading;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Microsoft.Identity.Client;
using Microsoft.IdentityModel.Claims;
using System.IdentityModel.Tokens.Jwt;
using System.Net;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;
using System.Text.RegularExpressions;

namespace SSO
{
	public partial class Login : System.Web.UI.Page
	{
		private static Dictionary<string, string> _routeConfig = null;

		private string tenantId = ConfigurationManager.AppSettings["AzureAD:TenantId"];
		private string clientId = ConfigurationManager.AppSettings["AzureAD:ClientId"];
		private string clientSecret = ConfigurationManager.AppSettings["AzureAD:ClientSecret"];
		private string redirectUri = ConfigurationManager.AppSettings["AzureAD:RedirectUri"];

		protected void Page_Load(object sender, EventArgs e)
		{
			ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
			LoadRouteConfig();
			string userid = null;
			string token = string.Empty;

			if (!IsPostBack)
			{
				if (Request.QueryString["code"] != null)
				{
					string authCode = Request.QueryString["code"];
					userid = ProcessAuthorizationCode(authCode);
					token = SSOUtilities.AESToken.GenerateToken(ConfigurationManager.AppSettings["appCode"], userid, ConfigurationManager.AppSettings["symmKey"], Convert.ToInt32(ConfigurationManager.AppSettings["expiryInMins"]));
					var responseCookie = new HttpCookie("SSO_TOKEN")
					{
						HttpOnly = true,
						Secure = true,
						Value = token,
						Expires = DateTime.Now.AddMinutes(Convert.ToInt32(ConfigurationManager.AppSettings["expiryInMins"]))
					};
					Response.Cookies.Set(responseCookie);

					string returnUrl = Request.QueryString["state"];
					if (!string.IsNullOrEmpty(returnUrl))
					{
						Response.Redirect(_routeConfig[returnUrl]);
					}
				}
				else
				{
					string route = Request.QueryString["route"];
					RedirectToAzureLogin(route);
				}
			}
		}

		private void RedirectToAzureLogin(string route)
		{
			string authority = $"https://login.microsoftonline.com/{tenantId}/oauth2";
			string authorizationUrl = $"{authority}/authorize?client_id={clientId}&response_type=code&redirect_uri={redirectUri}&response_mode=query&scope=openid&state={HttpUtility.UrlEncode(route)}";
			Response.Redirect(authorizationUrl);
		}

		private string ProcessAuthorizationCode(string code)
		{
			string authority = $"https://login.microsoftonline.com/{tenantId}/oauth2";
			var app = ConfidentialClientApplicationBuilder.Create(clientId)
						.WithClientSecret(clientSecret)
						.WithRedirectUri(redirectUri)
						.WithAuthority(new Uri(authority))
						.Build();

			var scopes = new[] { "openid", "profile", "email" };

			// Exchange the authorization code for tokens
			var result = app.AcquireTokenByAuthorizationCode(scopes, code)
							.ExecuteAsync().Result;

			// Extract user claims from the ID token
			var idToken = result.AccessToken;
			var handler = new JwtSecurityTokenHandler();
			var jwtToken = handler.ReadJwtToken(idToken);

			// Get claims
			var uniqueName = jwtToken.Claims.FirstOrDefault(c => c.Type == ConfigurationManager.AppSettings["AzureAD:Claim"])?.Value;

			if (uniqueName != null)
			{
				return ConvertADFSLoginID(uniqueName);
			}
			else
			{
				return "";
			}
		}

		private void LoadRouteConfig()
		{
			if (_routeConfig == null)
			{
				_routeConfig = new Dictionary<string, string>();
				string cfgAll = File.ReadAllText(Server.MapPath("~/route.cfg"));
				string[] cfgLines = cfgAll.Split('\n');
				foreach (string cfgLine in cfgLines)
				{
					string[] cfgFields = cfgLine.Split('|');
					if (cfgFields.Length == 2)
					{
						string key = Microsoft.Security.Application.Sanitizer.GetSafeHtmlFragment(cfgFields[0]).Trim().ToUpper();
						string value = Microsoft.Security.Application.Sanitizer.GetSafeHtmlFragment(cfgFields[1]).Trim();
						_routeConfig.Add(key, value);
					}
				}
			}
		}
		protected string ConvertADFSLoginID(string loginID)
		{
			if (loginID != null)
			{
				if (!loginID.Contains("\\"))
				{
					Regex reStaff = new Regex(@"@smu\.edu\.sg$");
					Regex reStudent = new Regex(@"@[a-zA-Z]+\.(smu\.edu\.sg)$");

					if (reStaff.IsMatch(loginID))
					{
						loginID = "smustf\\" + loginID;
					}
					else if (reStudent.IsMatch(loginID))
					{
						loginID = "smustu\\" + loginID;
					}
				}
				if (loginID.Contains("@"))
				{
					Regex reEmailDomain = new Regex(@"@.*smu.edu.sg");
					loginID = reEmailDomain.Replace(loginID, string.Empty);
				}
			}
			return loginID;
		}
	}
}