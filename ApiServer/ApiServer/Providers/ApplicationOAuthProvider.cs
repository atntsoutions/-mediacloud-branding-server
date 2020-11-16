using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using Microsoft.AspNet.Identity;
using Microsoft.AspNet.Identity.EntityFramework;
using Microsoft.AspNet.Identity.Owin;
using Microsoft.Owin.Security;
using Microsoft.Owin.Security.Cookies;
using Microsoft.Owin.Security.OAuth;
using ApiServer.Models;

using DataBase;
using BLAdmin;

namespace ApiServer.Providers
{
    public class ApplicationOAuthProvider : OAuthAuthorizationServerProvider
    {
        private readonly string _publicClientId;

        public ApplicationOAuthProvider(string publicClientId)
        {
            if (publicClientId == null)
            {
                throw new ArgumentNullException("publicClientId");
            }

            _publicClientId = publicClientId;
        }

        public override async Task GrantResourceOwnerCredentials(OAuthGrantResourceOwnerCredentialsContext context)
        {

            //var userManager = context.OwinContext.GetUserManager<ApplicationUserManager>();
            //ApplicationUser user = await userManager.FindAsync(context.UserName, context.Password);

            await Task.Run(() => {

              var userName = context.UserName;
              var password = context.Password;
              var company_code = context.Request.Headers["company-code"].ToString();
              var userService = new UserService(); // our created one
              string ipaddress = "";

              try
              {
                  ipaddress = context.Request.RemoteIpAddress.ToString();
              }
              catch (Exception)
              {
                  ipaddress = "BLANK IP";
              }

              try
              {
                  var user = userService.ValidateUser(userName, password, company_code, ipaddress);
                  if (user != null)
                  {
                      var claims = new List<Claim>()
                    {
                        new Claim(ClaimTypes.Sid, Convert.ToString(user.user_pkid)),
                        new Claim(ClaimTypes.Name, user.user_name),
                        new Claim(ClaimTypes.Email, user.user_email)
                    };
                      ClaimsIdentity oAuthIdentity = new ClaimsIdentity(claims, Startup.OAuthOptions.AuthenticationType);
                      var properties = CreateProperties(user);
                      var ticket = new AuthenticationTicket(oAuthIdentity, properties);
                      context.Validated(ticket);
                  }
                  else
                  {
                      context.SetError("invalid_grant", "The user name or password is incorrect");
                  }
              }
              catch (Exception Ex)
              {
                  context.SetError("invalid_grant", Ex.Message.ToString());
              }

          });

            /*

            if (user == null)
            {
                context.SetError("invalid_grant", "The user name or password is incorrect.");
                return;
            }

            ClaimsIdentity oAuthIdentity = await user.GenerateUserIdentityAsync(userManager,
               OAuthDefaults.AuthenticationType);
            ClaimsIdentity cookiesIdentity = await user.GenerateUserIdentityAsync(userManager,
                CookieAuthenticationDefaults.AuthenticationType);

            AuthenticationProperties properties = CreateProperties(user.UserName);
            AuthenticationTicket ticket = new AuthenticationTicket(oAuthIdentity, properties);
            context.Validated(ticket);
            context.Request.Context.Authentication.SignIn(cookiesIdentity);
            */
        }

        public override Task TokenEndpoint(OAuthTokenEndpointContext context)
        {
            foreach (KeyValuePair<string, string> property in context.Properties.Dictionary)
            {
                context.AdditionalResponseParameters.Add(property.Key, property.Value);
            }

            return Task.FromResult<object>(null);
        }

        public override Task ValidateClientAuthentication(OAuthValidateClientAuthenticationContext context)
        {
            // Resource owner password credentials does not provide a client ID.
            if (context.ClientId == null)
            {
                context.Validated();
            }

            return Task.FromResult<object>(null);
        }

        public override Task ValidateClientRedirectUri(OAuthValidateClientRedirectUriContext context)
        {
            if (context.ClientId == _publicClientId)
            {
                Uri expectedRootUri = new Uri(context.Request.Uri, "/");

                if (expectedRootUri.AbsoluteUri == context.RedirectUri)
                {
                    context.Validated();
                }
            }

            return Task.FromResult<object>(null);
        }


        public static AuthenticationProperties CreateProperties(User user)
        {
            IDictionary<string, string> data = new Dictionary<string, string>
            {
                { "userName", user.user_name },
                { "userpkid", user.user_pkid },
                { "usercode", user.user_code },
                { "usercompanyid", user.user_company_id},
                { "usercompanycode", user.user_company_code},
                { "userbranchid", user.user_branch_id},
                { "useremail", user.user_email},
                { "usersmanid", user.user_sman_id},
                { "usersmanname", user.user_sman_name},
                { "userlocalserver", user.user_local_server},
                { "useripaddress", user.user_ipaddress},
                { "usertokenid", user.user_token_id},
                { "user_branch_user", (user.user_branch_user) ? "Y" : "N"},
                { "user_region_id", user.user_region_id } ,
                { "user_vendor_id", user.user_vendor_id } ,
                { "user_role_id", user.user_role_id } ,
                { "user_role_name", user.user_role_name } ,
                { "user_role_rights_id", user.user_role_rights_id }, 
            };
            return new AuthenticationProperties(data);
        }


        public static AuthenticationProperties CreateProperties(string userName)
        {
            IDictionary<string, string> data = new Dictionary<string, string>
            {
                { "userName", userName }
            };
            return new AuthenticationProperties(data);
        }
    }
}