#r "System.Runtime"
#r "System.Threading.Tasks"

using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.SharePoint.Client;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;
using System.Linq;
using System.Net;
using System.Runtime;
using System;
using System.IO;

private static string ClientId = "068f1af4-ee37-45a8-885d-deec513c52d8";
private static string Cert = "O365S2S.pfx";
private static string CertPassword = "O365S2S";
private static string Authority = "https://login.windows.net/oleksiionsoftware.onmicrosoft.com/";
private static string Resource = "https://oleksiionsoftware.sharepoint.com/";

public static async Task<HttpResponseMessage> Run(HttpRequestMessage req, TraceWriter log)
{
    log.Info("Request for Destination List provisioning has been submitted.");
    try 
    {
        string siteUrl = req.GetQueryNameValuePairs()
            .FirstOrDefault(q => string.Compare(q.Key, "siteUrl", true) == 0)
            .Value;

        string listTitle = req.GetQueryNameValuePairs()
            .FirstOrDefault(q => string.Compare(q.Key, "listTitle", true) == 0)
            .Value;       

        await Provision(siteUrl, listTitle);

        return req.CreateResponse(HttpStatusCode.OK);
    } 
    catch (Exception ex)
    {
        return req.CreateResponse(HttpStatusCode.BadRequest, ex);
    }
}

private static async Task Provision(string siteUrl, string listTitle)
{
    using (var ctx = await GetClientContext(siteUrl))
    {
        ctx.Load(ctx.Web.Lists);
        ctx.ExecuteQuery();

        var list = ctx.Web.Lists.FirstOrDefault(_ => _.Title == listTitle);
        if (list == null)
        {
            var creationInfo = new ListCreationInformation
            {
                Title = listTitle,
                TemplateType = (int)ListTemplateType.GenericList
            };

            ctx.Web.Lists.Add(creationInfo);
            ctx.ExecuteQuery();

            list = ctx.Web.Lists.GetByTitle(listTitle);
            ctx.Load(list);
            ctx.ExecuteQuery();

            list.Fields.AddFieldAsXml(@"<Field ID='{1511BF28-A787-4061-B2E1-71F64CC93FD0}' Name='Count' DisplayName='Count' Type='Number' Required='FALSE' Group='Custom'></Field>", true, AddFieldOptions.AddFieldToDefaultView | AddFieldOptions.AddFieldInternalNameHint);
            list.Fields.AddFieldAsXml(@"<Field ID='{1511BF28-A787-4061-B2E1-71F64CC93FD1}' Name='Month' DisplayName='Month' Type='Text' Required='FALSE' Group='Custom'></Field>", true, AddFieldOptions.AddFieldToDefaultView | AddFieldOptions.AddFieldInternalNameHint);
            ctx.ExecuteQuery();
        }
    }
}

private async static Task<ClientContext> GetClientContext(string siteUrl)
{
    var authenticationContext = new AuthenticationContext(Authority, false);

    var certPath = Path.Combine(Environment.GetEnvironmentVariable("HOME"), "site\\wwwroot\\ProvisionSourceList\\", Cert);
    var cert = new X509Certificate2(System.IO.File.ReadAllBytes(certPath),
        CertPassword,
        X509KeyStorageFlags.Exportable |
        X509KeyStorageFlags.MachineKeySet |
        X509KeyStorageFlags.PersistKeySet);

    var authenticationResult = await authenticationContext.AcquireTokenAsync(Resource, new ClientAssertionCertificate(ClientId, cert));
    var token = authenticationResult.AccessToken;

    var ctx = new ClientContext(siteUrl);
    ctx.ExecutingWebRequest += (s, e) =>
    {
        e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + authenticationResult.AccessToken;
    };

    return ctx;
}