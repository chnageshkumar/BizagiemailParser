using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Configuration;
using System.IO;
using System.Xml;
using System.Net.Mail;
using System.Data.OleDb;
using System.Data;

namespace takeda.bizagi.common
{
    public static class Utility
    {
        public static string FormTag(string fieldName, string value)
        {
            // Following reformatting is needed to ensure the SOAP request is able to transmit special characters
            //&amp; (Translates to &)
            //&quot; (Translates to ") 
            //&apos; (Translates to ')

            string formattedValue = value.Replace("&", "&amp;").Replace("'", "&apos;").Replace("\"", "&quot;");
            return "<" + fieldName + ">" + formattedValue + "</" + fieldName + ">";
        }

        public static string GetKeyValueFromListKeyValuePairs(string key, List<KeyValuePair<string, string>> Valuepairs)
        {
            string retVal = string.Empty;

            foreach (KeyValuePair<string, string> pair in Valuepairs)
            {
                if (key == pair.Key)
                {
                    return pair.Value;
                }
            }
            return retVal;
        }

        public static void Log(string log)
        {
            try
            {
                string loglocation = string.Empty;
                string logCategory = string.Empty;
                if (ConfigurationManager.AppSettings["LogLocation"] != null)
                {
                    
                    loglocation = ConfigurationManager.AppSettings["LogLocation"].ToString();

                    if (ConfigurationManager.AppSettings["LogCategory"] != null)
                    {
                        logCategory = ConfigurationManager.AppSettings["LogCategory"].ToString();
                        Log(loglocation, logCategory, log);
                    }
                    else
                    {
                        Log(loglocation, "_", log);
                    }

                    
                }
            }
            catch
            {
            }
        }

        public static void Log(string customLogLocation, string logCategory, string log)
        {
            try
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine(DateTime.Now.ToLongTimeString() + " : " + log);
                string filename = "BizAgiLogging_" + logCategory + "_" + DateTime.Now.ToShortDateString().Replace('/', '_');
                File.AppendAllText(customLogLocation + filename + ".txt", sb.ToString());
            }
            catch
            {
            }
        }

        public static void LogWithoutTimeStamp(string customLogLocation, string logCategory, string log)
        {
            try
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine(log);
                string filename = "BizAgiLogging_" + logCategory + "_" + DateTime.Now.ToShortDateString().Replace('/', '_');
                File.AppendAllText(customLogLocation + filename + ".txt", sb.ToString());
            }
            catch
            {
            }
        }

        public static DataTable GetExcelDataAsTable(string excelPath, string sheetName)
        {
            System.Data.OleDb.OleDbConnection MyConnection;
            System.Data.OleDb.OleDbCommand MyCommand;
            System.Data.OleDb.OleDbDataAdapter da;
            MyConnection = new System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + excelPath + ";Extended Properties='Excel 12.0 Xml;HDR=Yes'");
            MyConnection.Open();
            MyCommand = new OleDbCommand("SELECT * From [" + sheetName + "$]", MyConnection);
            DataSet ds = new DataSet();
            da = new OleDbDataAdapter(MyCommand.CommandText, MyConnection);
            da.Fill(ds);
            MyConnection.Close();
            return ds.Tables[0];
        }

        public static string GetAppConfigValues(string key)
        {
            string retVal = string.Empty;
            try
            {
                if (ConfigurationManager.AppSettings[key] != null)
                    retVal = ConfigurationManager.AppSettings[key].ToString();
            }
            catch (Exception ex)
            {
                Utility.Log("*********Exception fetching the App config value for : " + key);
                throw;
            }

            return retVal;
        }

        public static List<List<KeyValuePair<string, string>>> ConvertSPUserGroupServiceToBizAgiUsers(XmlNode userDataXML)
        {
            Utility.Log("Method ConvertSPUsersToBizAgiUsers");
            // All Users
            List<List<KeyValuePair<string, string>>> retValuePairs = new List<List<KeyValuePair<string, string>>>();

            if (userDataXML != null)
            {
                // Check if it has Users Node
                if (userDataXML.HasChildNodes)
                {
                    // Check if it has Per User Node
                    if (userDataXML.ChildNodes[0].HasChildNodes)
                    {
                        XmlNodeList UserNodes = userDataXML.ChildNodes[0].ChildNodes;
                        foreach (XmlNode child in UserNodes)
                        {
                            // Per User
                            List<KeyValuePair<string, string>> perUser = new List<KeyValuePair<string, string>>();
                            bool excludeUser = false;
                            //User Name
                            if (child.Attributes["LoginName"] != null)
                            {
                                if (!string.IsNullOrEmpty(child.Attributes["LoginName"].Value))
                                {
                                    string[] split = child.Attributes["LoginName"].Value.Split(new char[] { '\\' });
                                    if (split.Length == 2)
                                    {
                                        perUser.Add(new KeyValuePair<string, string>("domain", split[0]));
                                        perUser.Add(new KeyValuePair<string, string>("userName", split[1]));
                                    }
                                    else
                                    {
                                        excludeUser = true;
                                    }
                                }
                                else
                                {
                                    excludeUser = true;
                                }
                            }

                            // Full Name
                            if (child.Attributes["Name"] != null)
                            {
                                perUser.Add(new KeyValuePair<string, string>("fullName", child.Attributes["Name"].Value));
                            }

                            //Email Address
                            if (child.Attributes["Email"] != null)
                            {
                                perUser.Add(new KeyValuePair<string, string>("contactEmail", child.Attributes["Email"].Value));
                            }

                            // Finally Add when all appropriate conditions are met.
                            if (!excludeUser)
                            {
                                retValuePairs.Add(perUser);
                            }
                        }
                    }
                }
            }

            Utility.Log("Method ConvertSPUsersToBizAgiUsers ...completed");

            return retValuePairs;
        }

        public static List<List<KeyValuePair<string, string>>> ConvertSPUserInformationListToBizAgiUsers(XmlNode userDataXML)
        {
            Utility.Log("Method ConvertSPUserInformationListToBizAgiUsers");
            // All Users
            List<List<KeyValuePair<string, string>>> retValuePairs = new List<List<KeyValuePair<string, string>>>();

            if (userDataXML != null)
            {
                // Check if it has Users Node
                if (userDataXML.HasChildNodes)
                {
                    // Check if it has Per User Node
                    if (userDataXML.ChildNodes[1].HasChildNodes)
                    {
                        XmlNodeList UserNodes = userDataXML.ChildNodes[1].ChildNodes;
                        foreach (XmlNode child in UserNodes)
                        {
                            if (child.Name == "z:row" && child.Attributes.Count != 0)
                            {
                                // Per User
                                List<KeyValuePair<string, string>> perUser = new List<KeyValuePair<string, string>>();
                                bool excludeUser = false;
                                //User Name
                                if (child.Attributes["ows_Name"] != null)
                                {
                                    if (!string.IsNullOrEmpty(child.Attributes["ows_Name"].Value))
                                    {
                                        string[] split = child.Attributes["ows_Name"].Value.Split(new char[] { '\\' });
                                        if (split.Length == 2)
                                        {
                                            perUser.Add(new KeyValuePair<string, string>("domain", split[0]));
                                            perUser.Add(new KeyValuePair<string, string>("userName", split[1]));
                                        }
                                        else
                                        {
                                            excludeUser = true;
                                        }
                                    }
                                    else
                                    {
                                        excludeUser = true;
                                    }
                                }

                                // Full Name
                                if (child.Attributes["ows_Title"] != null)
                                {
                                    perUser.Add(new KeyValuePair<string, string>("fullName", child.Attributes["ows_Title"].Value));
                                }

                                //Email Address
                                if (child.Attributes["ows_EMail"] != null)
                                {
                                    if (!string.IsNullOrEmpty(child.Attributes["ows_EMail"].Value))
                                        perUser.Add(new KeyValuePair<string, string>("contactEmail", child.Attributes["ows_EMail"].Value));
                                }
                                else
                                {
                                    excludeUser = true;
                                }

                                //Job Title
                                if (child.Attributes["ows_JobTitle"] != null)
                                {
                                    if (!string.IsNullOrEmpty(child.Attributes["ows_JobTitle"].Value))
                                    {
                                        perUser.Add(new KeyValuePair<string, string>("Title", child.Attributes["ows_JobTitle"].Value));
                                    }
                                    else
                                    {
                                        perUser.Add(new KeyValuePair<string, string>("Title", "(no title)"));
                                    }
                                }
                                else
                                {
                                    perUser.Add(new KeyValuePair<string, string>("Title", "(no title)"));
                                }

                                //Department
                                if (child.Attributes["ows_Department"] != null)
                                {
                                    if (!string.IsNullOrEmpty(child.Attributes["ows_Department"].Value))
                                    {
                                        perUser.Add(new KeyValuePair<string, string>("Department", child.Attributes["ows_Department"].Value));
                                    }
                                    else
                                    {
                                        perUser.Add(new KeyValuePair<string, string>("Department", "(no department)"));
                                    }
                                }
                                else
                                {
                                    perUser.Add(new KeyValuePair<string, string>("Department", "(no department)"));
                                }

                                // Finally Add when all appropriate conditions are met.
                                if (!excludeUser)
                                {
                                    retValuePairs.Add(perUser);
                                }
                            }
                        }
                    }
                }
            }
            Utility.Log("Method ConvertSPUserInformationListToBizAgiUsers... completed");
            return retValuePairs;
        }

        public static string Between(this string value, string a, string b)
        {
            int posA = value.IndexOf(a);
            int posB = value.LastIndexOf(b);
            if (posA == -1)
            {
                return "";
            }
            if (posB == -1)
            {
                return "";
            }
            int adjustedPosA = posA + a.Length;
            if (adjustedPosA >= posB)
            {
                return "";
            }
            return value.Substring(adjustedPosA, posB - adjustedPosA);
        }

        public static string PolishValue(string value)
        {
            if (value.Contains('<') && value.Contains('>'))
            {
                return value.Between("<", ">");
            }
            else if (value.Contains('"'))
            {
                return value.Replace('"', ' ');
            }
            else
            {
                return value;
            }
        }

        public static string ExtractKey(string value, int position)
        {
            string phValue = ExtractPlaceHolderValue(value);
            string[] parts = phValue.Split(new char[] { '#' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 2)
            {
                return parts[position];
            }
            else
            {
                return string.Empty;
            }
            
        }

        private static string ExtractPlaceHolderValue(string value)
        {
            if (!string.IsNullOrEmpty(value) && value.Contains("##"))
            {
                return value.Between("##", "##");
            }
            else
            {
                return string.Empty;
            }
        }

        public static void SendEmail(string subject, string body)
        {
            try
            {
                string smtpURL = ConfigurationManager.AppSettings["SMTPGateWay"];
                string to = ConfigurationManager.AppSettings["SendErrorNotificationsTo"];
                if (!string.IsNullOrEmpty(smtpURL))
                {
                    MailMessage mail = new MailMessage("noreply@takeda.com", to, subject, body);
                    SmtpClient smtp = new SmtpClient(smtpURL);
                    smtp.Send(mail);
                    Log("Email sent to : " + to);
                }
            }
            catch (Exception ex)
            {
                Log("Error Sending Email : " + subject + " Message: " + ex.Message);
            }
        }

        public static void SendEmail(string to, string subject, string body)
        {
            try
            {
                string smtpURL = ConfigurationManager.AppSettings["SMTPGateWay"];
                if (!string.IsNullOrEmpty(smtpURL))
                {
                    MailMessage mail = new MailMessage("noreply@takeda.com", to, subject, body);
                    SmtpClient smtp = new SmtpClient(smtpURL);
                    smtp.Send(mail);
                    Log("Email sent to : " + to);
                }
            }
            catch (Exception ex)
            {
                Log("Error Sending Email to: " + to + " Message: " + ex.Message);
            }
        }
    }
}
