using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using takeda.bizagi.connector.EntityManagerSOA;
using takeda.bizagi.common;
using System.Net;
using System.Xml;
using System.Configuration;

namespace takeda.bizagi.connector
{
    public class Connection
    {
        public EntityManagerSOA.EntityManagerSOA connObject = null;
        public string entitymangerSOASuffix = "webservices/entitymanagersoa.asmx";
        
        public Connection(string url)
        {
            connObject = new EntityManagerSOA.EntityManagerSOA();
            if (url.EndsWith("/"))
            {
                connObject.Url = url + entitymangerSOASuffix;
            }
            else
            {
                connObject.Url = url + "/" + entitymangerSOASuffix;
            }
            connObject.Timeout = 1200000;
            connObject.UseDefaultCredentials = true;
            connObject.PreAuthenticate = true;
            connObject.Credentials = CredentialCache.DefaultNetworkCredentials;
        }

        public XmlNode GetEntity(string entityName, string filterCondition)
        {
            XmlDocument requestNode = new XmlDocument();
            string xmlString = "<BizAgiWSParam>";
            xmlString += "<EntityData>";
            xmlString += "<EntityName>" + entityName + "</EntityName>";
            if (!string.IsNullOrEmpty(filterCondition))
            {
                xmlString += "<Filters>";
                xmlString += " <![CDATA[ " + filterCondition + " ]]>";
                xmlString += "</Filters>";
            }
            xmlString += "</EntityData>";
            xmlString += "</BizAgiWSParam>";
            requestNode.LoadXml(xmlString);
            return connObject.getEntities(requestNode);
        }

        public XmlNode GetEntitybySchema(string entityName, string filterCondition, string XSDLocation)
        {
            XmlDocument requestNode = new XmlDocument();
            XmlDocument XSDSchema = new XmlDocument();
            XSDSchema.Load(XSDLocation);
            
            string xmlString = "<BizAgiWSParam>";
            xmlString += "<EntityData>";
            xmlString += "<EntityName>" + entityName + "</EntityName>";
            if (!string.IsNullOrEmpty(filterCondition))
            {
                xmlString += "<Filters>";
                xmlString += " <![CDATA[ " + filterCondition + " ]]>";
                xmlString += "</Filters>";
            }
            xmlString += "</EntityData>";
            xmlString += "</BizAgiWSParam>";
            requestNode.LoadXml(xmlString);
            return connObject.getEntitiesUsingSchema(requestNode, XSDSchema);
        }

        public XmlNode SaveNewEntity(string entityName, XmlDocument request)
        {
            XmlDocument requestNode = new XmlDocument();
            string xmlString = "<BizAgiWSParam>";
            xmlString += "<Entities>";
            xmlString += "<" + entityName + ">";
            xmlString += request.InnerXml;
            xmlString += "</" + entityName + ">";
            xmlString += "</Entities>";
            xmlString += "</BizAgiWSParam>";
            Utility.Log("SAVE: " + xmlString);
            requestNode.LoadXml(xmlString);
            return connObject.saveEntity(requestNode);

        }

        public XmlNode SaveNewEntity(string entityName, List<KeyValuePair<string,string>> request)
        {
            string start = DateTime.Now.ToLongTimeString();
            XmlDocument requestNode = new XmlDocument();
            string xmlString = "<BizAgiWSParam>";
            xmlString += "<Entities>";
            xmlString += "<" + entityName + ">";
            foreach (KeyValuePair<string, string> pair in request)
            {
                xmlString += Utility.FormTag(pair.Key, pair.Value);
            }
            if (entityName == "WFUSER")
            {
                xmlString += Utility.FormTag("Organizations", "<Code>1</Code>");
            }
            xmlString += "</" + entityName + ">";
            xmlString += "</Entities>";
            xmlString += "</BizAgiWSParam>";
            Utility.Log("SAVE: (" + start + ")" + xmlString);
            requestNode.LoadXml(xmlString);
            return connObject.saveEntity(requestNode);

        }

        public XmlNode SaveExistingEntity(string entityName, string pkey, List<KeyValuePair<string, string>> request)
        {
            string start = DateTime.Now.ToLongTimeString();
            XmlDocument requestNode = new XmlDocument();
            string xmlString = "<BizAgiWSParam>";
            xmlString += "<Entities>";
            xmlString += "<" + entityName + " key=\"" + pkey + "\">";
            foreach (KeyValuePair<string, string> pair in request)
            {
                xmlString += Utility.FormTag(pair.Key, pair.Value);
            }
            xmlString += "</" + entityName + ">";
            xmlString += "</Entities>";
            xmlString += "</BizAgiWSParam>";
            Utility.Log("SAVE E (" + start + "): " + xmlString);
            requestNode.LoadXml(xmlString);
            return connObject.saveEntity(requestNode);

        }

        public XmlNode SaveExistingEntity(string entityName, string pkey, string updateKey, string updateValue)
        {
            string start = DateTime.Now.ToLongTimeString();
            XmlDocument requestNode = new XmlDocument();
            string xmlString = "<BizAgiWSParam>";
            xmlString += "<Entities>";
            xmlString += "<" + entityName + " key=\"" + pkey + "\">";
            xmlString += Utility.FormTag(updateKey, updateValue);
            xmlString += "</" + entityName + ">";
            xmlString += "</Entities>";
            xmlString += "</BizAgiWSParam>";
            Utility.Log("SAVE E (" + start + "): " + xmlString);
            requestNode.LoadXml(xmlString);
            return connObject.saveEntity(requestNode);

        }

        public XmlNode SaveExistingWFUSER_Special(string pkey, string idBossUserValue, string enabledValue, bool ClearUserAccountControl)
        {
            bool changeExists = false;
            XmlDocument requestNode = new XmlDocument();
            XmlNode responseNode = null;
            string entityName = "WFUSER";
            string xmlString = "<BizAgiWSParam>";
            xmlString += "<Entities>";
            xmlString += "<" + entityName + " key=\"" + pkey + "\">";
            if (!string.IsNullOrEmpty(idBossUserValue))
            {
                xmlString += Utility.FormTag("idBossUser", idBossUserValue);
                changeExists = true;
            }
            if (!string.IsNullOrEmpty(enabledValue))
            {
                xmlString += Utility.FormTag("enabled", enabledValue);
                changeExists = true;
            }
            if (ClearUserAccountControl)
            {
                xmlString += Utility.FormTag("UserAccountControl", "");
                changeExists = true;
            }
            xmlString += "</" + entityName + ">";
            xmlString += "</Entities>";
            xmlString += "</BizAgiWSParam>";
            Utility.Log("SAVE E : " + xmlString);
            requestNode.LoadXml(xmlString);
            if (changeExists)
            {
                responseNode = connObject.saveEntity(requestNode);
            }
            return responseNode;

        }

        public XmlNode SaveExistingWFUSER_MigrateDomain(string pkey, string newDomain)
        {
            bool changeExists = false;
            XmlDocument requestNode = new XmlDocument();
            XmlNode responseNode = null;
            string entityName = "WFUSER";
            string xmlString = "<BizAgiWSParam>";
            xmlString += "<Entities>";
            xmlString += "<" + entityName + " key=\"" + pkey + "\">";
            if (!string.IsNullOrEmpty(newDomain))
            {
                xmlString += Utility.FormTag("domain", newDomain);
                changeExists = true;
            }
            xmlString += "</" + entityName + ">";
            xmlString += "</Entities>";
            xmlString += "</BizAgiWSParam>";
            Utility.Log("SAVE E : " + xmlString);
            requestNode.LoadXml(xmlString);
            if (changeExists)
            {
                responseNode = connObject.saveEntity(requestNode);
            }
            return responseNode;

        }

        public XmlNode SaveExistingWFUSER_EnableDisable(string pkey, bool enabled)
        {
            XmlDocument requestNode = new XmlDocument();
            XmlNode responseNode = null;
            string entityName = "WFUSER";
            string xmlString = "<BizAgiWSParam>";
            xmlString += "<Entities>";
            xmlString += "<" + entityName + " key=\"" + pkey + "\">";
            xmlString += Utility.FormTag("enabled", enabled.ToString());
            xmlString += "</" + entityName + ">";
            xmlString += "</Entities>";
            xmlString += "</BizAgiWSParam>";
            Utility.Log("SAVE E : " + xmlString);
            requestNode.LoadXml(xmlString);
            responseNode = connObject.saveEntity(requestNode);
            return responseNode;

        }

        public XmlNode SaveExisting_BULKUpdate(string FormedBulkQuery)
        {
            string start = DateTime.Now.ToLongTimeString();
            XmlDocument requestNode = new XmlDocument();
            XmlNode responseNode = null;
            string xmlString = "<BizAgiWSParam>";
            xmlString += "<Entities>";
            xmlString += FormedBulkQuery;
            xmlString += "</Entities>";
            xmlString += "</BizAgiWSParam>";
            Utility.Log("BULK saveEntity Started - (" + start + ")");
            Utility.Log(xmlString);
            requestNode.LoadXml(xmlString);
            responseNode = connObject.saveEntity(requestNode);
            string end = DateTime.Now.ToLongTimeString();
            Utility.Log("BULK saveEntity Completed - (" + end + ")");
            return responseNode;
        }

        public bool QueryHasValues(string entityName, string filterCondition)
        {
            bool retVal = false;
            try
            {
                XmlNode response = GetEntity(entityName, filterCondition);
                return (response.Name != "BizAgiWSError" && response.FirstChild.HasChildNodes);
            }
            catch
            {
                retVal = false;
            }
            return retVal;
        }

        public string FindQueryKey(string entityName, string filterCondition)
        {
            string retVal = "0";
            try
            {
                XmlNode response = GetEntity(entityName, filterCondition);
                if (response.Name != "BizAgiWSError" && response.FirstChild.HasChildNodes)
                {
                    if (response.FirstChild.ChildNodes.Count >= 1)
                    {
                        retVal = response.FirstChild.ChildNodes[0].Attributes["key"].Value;
                    }
                    else
                    {
                        retVal = "Too Many records!";
                    }
                }
                else
                {
                    if (response.InnerXml == "<Entities></Entities>")
                    {
                        retVal = "BizAgiWSError : No User Found.";
                    }
                    else
                    {
                        retVal = "BizAgiWSError : " + response.Name + response.InnerXml;
                    }
                }
            }
            catch(Exception ex)
            {
                retVal = ex.Message;
            }
            return retVal;
        }

        public bool SaveUpdatePerRecordInEntities(string entityName, bool overrideExisting, string key, List<KeyValuePair<string, string>> ValuePairs)
        {
            XmlNode ResponseNode = null;
            string PrimaryValue = Utility.GetKeyValueFromListKeyValuePairs(key, ValuePairs);
            string s = FindQueryKey(entityName, key + " = '" + PrimaryValue + "'");
            int i = 0;
            string mode = string.Empty;
            if (overrideExisting && int.TryParse(s, out i))
            {
                if (i != 0)
                {
                    ResponseNode = SaveExistingEntity(entityName, s, ValuePairs);
                    mode = "Updated";
                }
            }
            else
            {
                ResponseNode = SaveNewEntity(entityName, ValuePairs);
                mode = "Created";
            }
            //Check if there was no error
            if (ResponseNode != null)
            {
                if(ResponseNode.HasChildNodes)
                if (ResponseNode.ChildNodes[0].Name.ToUpper() == entityName.ToUpper())
                {
                    Utility.Log(entityName + " entity Item " + mode + " : " + ResponseNode.InnerText);
                    return true;
                }
            }
            Utility.Log(entityName + " entity Item " + mode + " ... Failed");
            return false;
        }

        public XmlNode saveEntity(XmlDocument requestNode)
        {
            return connObject.saveEntity(requestNode);
        }

        #region UserManagement

        public bool CheckUserExists(string userName, string domain)
        {
            return QueryHasValues("WFUSER", "userName = '" + userName + "' AND domain = '" + domain + "'");
        }

        public string GetUserID(string userName, string domain)
        {
            return FindQueryKey("WFUSER", "userName = '" + userName + "' AND domain = '" + domain + "'");
        }

        public bool EnableDisableUser(string ID, bool Enabled)
        {
            bool retVal = false;
            XmlNode response = SaveExistingWFUSER_EnableDisable(ID, Enabled);
            if (response.Name != "BizAgiWSError" && response.FirstChild.HasChildNodes)
            {
                if (response.FirstChild.ChildNodes.Count > 0)
                {
                    retVal = true;
                }
                else
                {
                    retVal = false;
                }
            }
            else
            {
                retVal = false;
                Utility.Log("BizAgiWSError EnableDisableUser: " + response.Name + response.InnerXml);
            }

            return retVal;
        }

        public bool ChangeUserDomain(string ID, string newDomain)
        {
            bool retVal = false;
            XmlNode response = SaveExistingWFUSER_MigrateDomain(ID, newDomain);
            if (response.Name != "BizAgiWSError" && response.FirstChild.HasChildNodes)
            {
                if (response.FirstChild.ChildNodes.Count > 0)
                {
                    retVal = true;
                }
                else
                {
                    retVal = false;
                }
            }
            else
            {
                retVal = false;
                Utility.Log("BizAgiWSError ChangeUserDomain: " + response.Name + response.InnerXml);
            }

            return retVal;
        }

        #endregion
    }
}
