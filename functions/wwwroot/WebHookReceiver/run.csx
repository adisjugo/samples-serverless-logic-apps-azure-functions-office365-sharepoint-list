#r "Newtonsoft.Json"

using System.Net;
using System.Configuration;
using System.Net.Http;
using System.Net.Http.Formatting;
using Newtonsoft.Json;

public static async Task<HttpResponseMessage> Run(HttpRequestMessage req, TraceWriter log)
{
    log.Info("Request started.");

    string validationToken = req.GetQueryNameValuePairs()
        .FirstOrDefault(q => string.Compare(q.Key, "validationToken", true) == 0)
        .Value;

    if(!string.IsNullOrEmpty(validationToken))
    {
        log.Info($"Validation request from SharePoint: {validationToken}");
        var resp = new HttpResponseMessage(HttpStatusCode.OK);
        resp.Content = new StringContent(validationToken, System.Text.Encoding.UTF8, "text/plain");
        return resp;
    }    

    var content = await req.Content.ReadAsStringAsync();
    if(!string.IsNullOrEmpty(content)) 
    {
        log.Info($"Notification request from SharePoint List: {content}");
        var notifications = JsonConvert.DeserializeObject<ResponseModel<NotificationModel>>(content);    
        var objectContent = new ObjectContent<ResponseModel<NotificationModel>>(notifications, new JsonMediaTypeFormatter());       
        var httpClient = new HttpClient();
        await httpClient.PostAsync(ConfigurationManager.AppSettings["RT_URL"], objectContent);
        return new HttpResponseMessage(HttpStatusCode.OK);
    }   

    return new HttpResponseMessage(HttpStatusCode.OK);
}

public class ResponseModel<T>
{
    [JsonProperty(PropertyName = "value")]
    public List<T> Value { get; set; }
}

public class NotificationModel
{
    [JsonProperty(PropertyName = "subscriptionId")]
    public string SubscriptionId { get; set; }

    [JsonProperty(PropertyName = "clientState")]
    public string ClientState { get; set; }

    [JsonProperty(PropertyName = "expirationDateTime")]
    public DateTime ExpirationDateTime { get; set; }

    [JsonProperty(PropertyName = "resource")]
    public string Resource { get; set; }

    [JsonProperty(PropertyName = "tenantId")]
    public string TenantId { get; set; }

    [JsonProperty(PropertyName = "siteUrl")]
    public string SiteUrl { get; set; }

    [JsonProperty(PropertyName = "webId")]
    public string WebId { get; set; }
}