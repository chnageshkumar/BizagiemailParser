using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;

namespace takeda.bizagi.connector
{
    public class BizAgiCacheManagement
    {
        public BizAgiCacheWebservice.Cache connObject = null;
        public string SOASuffix = "webservices/Cache.asmx";

        public BizAgiCacheManagement(string url)
        {
            connObject = new BizAgiCacheWebservice.Cache();
            if (url.EndsWith("/"))
            {
                connObject.Url = url + SOASuffix;
            }
            else
            {
                connObject.Url = url + "/" + SOASuffix;
            }
            connObject.UseDefaultCredentials = true;
            connObject.PreAuthenticate = true;
            connObject.Credentials = CredentialCache.DefaultNetworkCredentials;
        }

        public void RunCacheClearingRoutine()
        {
            connObject.CleanRenderCache();
            connObject.CleanTracing();
            connObject.CleanUpCache("*", "*");
            connObject.FreeLocalizationResources();
            connObject.UpdatePortal();
            connObject.cleanParameters();
            connObject.cleanUpRuleCache();
        }
    }
}
