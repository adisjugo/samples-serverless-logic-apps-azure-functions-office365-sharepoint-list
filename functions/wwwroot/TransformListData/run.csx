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
    log.Info("Request for data transformation has been submitted. ");

    try 
    {
        string siteUrl = req.GetQueryNameValuePairs()
            .FirstOrDefault(q => string.Compare(q.Key, "siteUrl", true) == 0)
            .Value;

        string sourceListTitle = req.GetQueryNameValuePairs()
            .FirstOrDefault(q => string.Compare(q.Key, "sourceListTitle", true) == 0)
            .Value; 

        string destListTitle = req.GetQueryNameValuePairs()
            .FirstOrDefault(q => string.Compare(q.Key, "destListTitle", true) == 0)
            .Value;       

        await Transform(siteUrl, sourceListTitle, destListTitle);

        return req.CreateResponse(HttpStatusCode.OK);
    } 
    catch (Exception ex)
    {
        return req.CreateResponse(HttpStatusCode.BadRequest, ex);
    }
}

private static async Task Transform(string siteUrl, string sourceListTitle, string destListTitle)
{
    using (var ctx = await GetClientContext(siteUrl))
    {
        var sourceList = ctx.Web.Lists.GetByTitle(sourceListTitle);
        var items = sourceList.GetItems(CamlQuery.CreateAllItemsQuery());
        ctx.Load(items, _ => _.Include(
                i => i.DisplayName,
                i => i["Date"]));
        ctx.ExecuteQuery();

        var dict = new Dictionary<int, Dictionary<string, int>>();
        foreach (var item in items)
        {
            var monthNumber = ((DateTime)item["Date"]).Month;
            if (!dict.ContainsKey(monthNumber))
            {
                dict.Add(monthNumber, new Dictionary<string, int>());
            }

            if (!dict[monthNumber].ContainsKey(item.DisplayName))
            {
                dict[monthNumber].Add(item.DisplayName, 0);
            }

            dict[monthNumber][item.DisplayName]++;
        }

        var destList = ctx.Web.Lists.GetByTitle(destListTitle);
        var destListItems = destList.GetItems(CamlQuery.CreateAllItemsQuery());
        ctx.Load(destListItems);
        ctx.ExecuteQuery();

        while (destListItems.Count != 0)
        {
            destListItems[0].DeleteObject();
        }

        ctx.ExecuteQuery();

        foreach (var month in dict.Keys)
        {
            foreach (var rec in dict[month])
            {
                var item = destList.AddItem(new ListItemCreationInformation());
                item["Month"] = month;
                item["Title"] = rec.Key;
                item["Count"] = rec.Value;
                item.SystemUpdate();
            }
        }

        ctx.ExecuteQuery();
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