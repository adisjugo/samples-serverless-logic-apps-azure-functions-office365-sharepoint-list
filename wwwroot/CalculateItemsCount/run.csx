#r "Newtonsoft.Json"

using System.Net;
using System.Security;
using System.Configuration;
using Microsoft.SharePoint.Client;

public static async Task<HttpResponseMessage> Run(HttpRequestMessage req, TraceWriter log)
{
    string origin = req.GetQueryNameValuePairs()
        .FirstOrDefault(q => string.Compare(q.Key, "origin", true) == 0)
        .Value;
    
    string sourceListTitle = req.GetQueryNameValuePairs()
        .FirstOrDefault(q => string.Compare(q.Key, "sourceListTitle", true) == 0)
        .Value;

    string destListTitle = req.GetQueryNameValuePairs()
        .FirstOrDefault(q => string.Compare(q.Key, "destListTitle", true) == 0)
        .Value;

    log.Info($"Request started. Origin: {origin}, Target: {ConfigurationManager.AppSettings["SPO_SITE"]}, Source List: {sourceListTitle}, Destination List: {destListTitle}");

    using (var ctx = GetClientContext())
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

    return req.CreateResponse(HttpStatusCode.OK, "OK");
}

static ClientContext GetClientContext()
{
    var ctx = new ClientContext(ConfigurationManager.AppSettings["SPO_SITE"]);
    var ssp = new SecureString();
    foreach (var c in ConfigurationManager.AppSettings["SPO_PASSWORD"])
    {
        ssp.AppendChar(c);
    }

    ctx.Credentials = new SharePointOnlineCredentials(ConfigurationManager.AppSettings["SPO_LOGIN"], ssp);
    return ctx;
}