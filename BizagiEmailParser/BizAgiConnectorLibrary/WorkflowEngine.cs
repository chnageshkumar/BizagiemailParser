using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using takeda.bizagi.connector.WorkflowEngineSOA;
using System.Net;
using System.IO;
using takeda.bizagi.common;

namespace takeda.bizagi.connector
{
    public class WorkflowEngine
    {
        public WorkflowEngineSOA.WorkflowEngineSOA connObject = null;
        public string WorkflowEngineSOASuffix = "webservices/workflowenginesoa.asmx";

        public WorkflowEngine(string url)
        {
            connObject = new WorkflowEngineSOA.WorkflowEngineSOA();
            if (url.EndsWith("/"))
            {
                connObject.Url = url + WorkflowEngineSOASuffix;
            }
            else
            {
                connObject.Url = url + "/" + WorkflowEngineSOASuffix;
            }
            connObject.UseDefaultCredentials = true;
            connObject.PreAuthenticate = true;
            connObject.Credentials = CredentialCache.DefaultNetworkCredentials;
        }

        public XmlNode GetEventsForACase(string caseNumber)
        {
            XmlDocument requestNode = new XmlDocument();
            string xmlString = "<BizAgiWSParam>";
            xmlString += "<radNumber>" + caseNumber + "</radNumber>";
            xmlString += "</BizAgiWSParam>";
            requestNode.LoadXml(xmlString);
            return connObject.getEvents(requestNode);
        }

        public string GetEmailEventName(string caseNumber, string matchName)
        {
            string eventName = string.Empty;
            XmlNode response = GetEventsForACase(caseNumber);
            if (response != null)
            {
                if (response.HasChildNodes)
                {
                   foreach(XmlNode node in response.ChildNodes)
                   {
                       if (node.SelectSingleNode("task/taskName") != null)
                       {
                           XmlNode currentTask = node.SelectSingleNode("task/taskDisplayName");
                           if (currentTask.InnerText == matchName)
                           {
                               return node.SelectSingleNode("task/taskName").InnerText;
                           }
                       }
                   }
                }
            }
            return eventName;
        }

        public XmlNode PerformEmailLogActivity(string caseNumber, string outerEntityName, string eventName, string from, string to, string subject, string timesent, string body,string file)
        {
            string openEntity = string.Empty;
            string closeEntity = string.Empty;
            if (outerEntityName != "")
            {
                openEntity = "<" + outerEntityName + ">";
                closeEntity = "</" + outerEntityName + ">";
            }
            else
            {
                return null;
            }
            XmlDocument requestNode = new XmlDocument();
            string xmlString = "<BizAgiWSParam>";
            xmlString += "<domain>domain</domain>";
            xmlString += "<userName>admon</userName>";
            xmlString += "<ActivityData>";
            xmlString += "<radNumber>" + caseNumber + "</radNumber>";
            xmlString += "<taskName>" + eventName + "</taskName>";
            xmlString += "</ActivityData>";
            xmlString += "<Entities>";
            xmlString += openEntity;
            if (outerEntityName == "MassChangeRequest" || outerEntityName == "DataManagement" || outerEntityName == "UserManagementProcess")
            {
                xmlString += "<EmailOutgoing>false</EmailOutgoing>";
            }
            xmlString += "<EmailLog>";
            xmlString += "<RequestID>" + caseNumber + "</RequestID>";
            xmlString += "<Outgoing>false</Outgoing>";
            xmlString += "<TimeSent>" + timesent + "</TimeSent>";
            xmlString += "<EmailFrom>" + from + "</EmailFrom>";
            xmlString += "<EmailTo> " + to + " </EmailTo>";
            xmlString += "<Subject>" + subject + "</Subject>";
            xmlString += "<Message><![CDATA[" + body + "]]></Message>";
            xmlString += "<Attachment><File fileName=\"EmailResponse.eml\">" + ConvertToBase64(file) + "</File></Attachment>";
            xmlString += "</EmailLog>";
            xmlString += closeEntity;
            xmlString += "</Entities>";
            xmlString += "</BizAgiWSParam>";
            requestNode.LoadXml(xmlString);
            return connObject.performActivity(requestNode);
        }

        public XmlNode PerformEmailLogNewCase(string CreateUserName, string CreateUserDomain, string processName, string startingEntity, string fromEmail, string timesent, string body, string file, string subject)
        {
            XmlDocument requestNode = new XmlDocument();
            string xmlString = "<BizAgiWSParam>";
            xmlString += "<domain>" + CreateUserDomain + "</domain>";
            xmlString += "<userName>" + CreateUserName + "</userName>";
            xmlString += "<Cases>";
            xmlString += "<Case>";
            xmlString += "<Process>" + processName + "</Process>";
            xmlString += "<Entities>";
            xmlString += "<" + startingEntity + ">";
            xmlString += "<RequestorEmail>" + fromEmail + "</RequestorEmail>";
            if (processName == "TndEHelpdesk" || processName == "ServiceDesk")
            {
                if (processName == "TndEHelpdesk")
                {
                    xmlString += "<RequestDetails><![CDATA[" + body + "]]></RequestDetails>";
                    xmlString += "<EmailAttachment><File fileName=\"EmailAttachment.eml\">" + ConvertToBase64(file) + "</File></EmailAttachment>";
                }
                else
                {
                    xmlString += "<RequestDescription><![CDATA[" + body + "]]></RequestDescription>";
                    xmlString += "<EmailAttachments><File fileName=\"EmailAttachment.eml\">" + ConvertToBase64(file) + "</File></EmailAttachments>";
                    xmlString += "<RequestSubject>" + subject + "</RequestSubject>";
                }
                xmlString += "<CreationType><CreationCode>EM</CreationCode></CreationType>";
            }
            else
            {
                xmlString += "<Description><![CDATA[" + body + "]]></Description>";
                xmlString += "<Attachment><File fileName=\"EmailResponse.eml\">" + ConvertToBase64(file) + "</File></Attachment>";
            }
            if (processName != "ServiceDesk")
            {
                xmlString += "<CreationDate>" + timesent + "</CreationDate>";
            }
            xmlString += "</" + startingEntity + ">";
            xmlString += "</Entities>";
            xmlString += "</Case>";
            xmlString += "</Cases>";
            xmlString += "</BizAgiWSParam>";
            requestNode.LoadXml(xmlString);
            return connObject.createCases(requestNode);
        }

        //Additional Functionality Added By Abhiram Venkata (abhiram.dv@infosys.com) 4-jan-2018
        public XmlNode CreateCase(string CreateUserName, string CreateUserDomain, string processName, string entityName, List<KeyValuePair<string, string>> request, string entityNameWithFileType, string fileName, string bytes)
        {
            XmlDocument requestNode = new XmlDocument();
            string xmlString = "<BizAgiWSParam>";
            xmlString += "<domain>" + CreateUserDomain + "</domain>";
            xmlString += "<userName>" + CreateUserName + "</userName>";
            xmlString += "<Cases>";
            xmlString += "<Case>";
            xmlString += "<Process>" + processName + "</Process>";
            xmlString += "<Entities>";
            xmlString += "<" + entityName + ">";
            foreach (KeyValuePair<string, string> pair in request)
            {
                xmlString += Utility.FormTag(pair.Key, pair.Value);
            }
            xmlString += "<" + entityNameWithFileType+">";
            xmlString += string.Format("<File fileName=\"{0}\">{1}</File>", fileName, bytes);
            xmlString += "</" + entityNameWithFileType + ">";
            xmlString += "</" + entityName + ">";
            xmlString += "</Entities>";
            xmlString += "</Case>";
            xmlString += "</Cases>";
            xmlString += "</BizAgiWSParam>";
            requestNode.LoadXml(xmlString);
            return connObject.createCases(requestNode);
        }
        public XmlNode UpdateCase(string CreateUserName, string CreateUserDomain, string processName,string caseNumber, string activityName, string entityName, List<KeyValuePair<string, string>> request, string entityNameWithFileType, string fileName, string bytes)
        {
            XmlDocument requestNode = new XmlDocument();
            string xmlString = "<BizAgiWSParam>";
            xmlString += "<domain>" + CreateUserDomain + "</domain>";
            xmlString += "<userName>" + CreateUserName + "</userName>";
            xmlString += "<ActivityData>";
            xmlString += Utility.FormTag("radNumber", caseNumber);
            xmlString += Utility.FormTag("taskName", activityName);
            xmlString += "</ActivityData>";
            //xmlString += "<Cases>";
            //xmlString += "<Case>";
            //xmlString += "<Process>" + processName + "</Process>";
            xmlString += "<Entities>";
            xmlString += "<" + entityName + ">";
            foreach (KeyValuePair<string, string> pair in request)
            {
                xmlString += Utility.FormTag(pair.Key, pair.Value);
            }
            xmlString += "<" + entityNameWithFileType + ">";
            xmlString += string.Format("<File fileName=\"{0}\">{1}</File>", fileName, bytes);
            xmlString += "</" + entityNameWithFileType + ">";
            xmlString += "</" + entityName + ">";
            xmlString += "</Entities>";
            xmlString += "</BizAgiWSParam>";
            requestNode.LoadXml(xmlString);
            return connObject.saveActivity(requestNode);
        }

        //public XmlNode UpdateCase(string CreateUserName, string CreateUserDomain, string processName, string entityName, List<KeyValuePair<string, string>> request, string entityNameWithFileType, string fileName, string path)
        //{
        //    XmlDocument requestNode = new XmlDocument();
        //    string filePath = string.Format("{0}\\{1}", fileName, path);
        //    string xmlString = "<BizAgiWSParam>";
        //    xmlString += "<domain>" + CreateUserDomain + "</domain>";
        //    xmlString += "<userName>" + CreateUserName + "</userName>";
        //    xmlString += "<Cases>";
        //    xmlString += "<Case>";
        //    xmlString += "<Process>" + processName + "</Process>";
        //    xmlString += "<Entities>";
        //    xmlString += "<" + entityName + ">";
        //    foreach (KeyValuePair<string, string> pair in request)
        //    {
        //        xmlString += Utility.FormTag(pair.Key, pair.Value);
        //    }
        //    xmlString += "<" + entityNameWithFileType + ">";
        //    xmlString += string.Format("<File fileName=\"{0}\">{1}</File>", fileName, ConvertToBase64(filePath));
        //    xmlString += "</" + entityNameWithFileType + ">";
        //    xmlString += "</" + entityName + ">";
        //    xmlString += "</Entities>";
        //    xmlString += "</Case>";
        //    xmlString += "</Cases>";
        //    xmlString += "</BizAgiWSParam>";
        //    requestNode.LoadXml(xmlString);
        //    return connObject.createCases(requestNode);
        //}

        public string ConvertToBase64(string file)
        {
            FileStream stream = File.Open(file, FileMode.Open, FileAccess.Read);
            //The fileStream is loaded into a BinaryReader
            BinaryReader binaryReader = new BinaryReader(stream);

            //Read the bytes and save them in an array
            byte[] oBytes = binaryReader.ReadBytes(Convert.ToInt32(stream.Length));

            //Load an empty StringBuilder into an XmlWriter
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            TextWriter tw = new StringWriter(sb);
            XmlTextWriter m_XmlWriter = new XmlTextWriter(tw);

            //Transform the bytes in the StringBuilder into Base64
            m_XmlWriter.WriteBase64(oBytes, 0, oBytes.Length);
            stream.Close();
            return sb.ToString();

        }
    }
}
