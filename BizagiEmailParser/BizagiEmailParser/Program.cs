using S22.Imap;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using takeda.bizagi.connector;

namespace BizagiEmailParser
{
    class Program
    {
        static AutoResetEvent reconnectEvent = new AutoResetEvent(false);
        static ImapClient client;
        static string userName = ConfigurationManager.AppSettings["userName"];
        static string password = ConfigurationManager.AppSettings["Password"];
        static string host = ConfigurationManager.AppSettings["Host"];
        static int port = Convert.ToInt32(ConfigurationManager.AppSettings["Port"]);
        static string TempMailsStoragePath = ConfigurationManager.AppSettings["TempMailsStoragePath"];
        static string bizagiUrl = ConfigurationManager.AppSettings["BizagiUrl"];
        static string bizagiUserName = ConfigurationManager.AppSettings["bizagiUserName"];
        static string bizagiDomain = ConfigurationManager.AppSettings["bizagiDomain"];
        static string bizagiProcessName = ConfigurationManager.AppSettings["bizagiProcessName"];
        static string bizagiEntityName = ConfigurationManager.AppSettings["bizagiEntityName"];
        static string bizagiEmailSubjectColumnName = ConfigurationManager.AppSettings["bizagiEmailSubjectColumnName"];
        static string bizagiEmailBodyCOlumnName = ConfigurationManager.AppSettings["bizagiEmailBodyCOlumnName"];
        static string bizagiEmailFileAttributeName = ConfigurationManager.AppSettings["bizagiEmailFileAttributeName"];
        static string readMessagesFilterAccount = ConfigurationManager.AppSettings["readMessagesFilterAccount"];
        static bool newMessageSatisfiesCondition = false;
        static uint newMessageUint;
        static MailMessage msg;

        static void Main(string[] args)
        {
            while (true)
            {
                try
                {
                    #region ReadUnreadMessages
                    Console.Write("Connecting...");
                    InitializeClient();
                    Console.Write("Ok");
                    EmptyTemporatyMailFolder();
                    var unreadMessages = GetUnreadFilteredMailMessages();
                    if(unreadMessages.Count() > 1)
                        Console.WriteLine("\nFound Some Unread Messages ...... storing them .... Just a sec\n");
                    foreach (var message in unreadMessages)
                    {
                        WriteMessageToFileWIthHelpOfSmtp(message.Value);
                        var files = ReadAllEmlFiles();
                        if (files.Length > 0)
                        {
                            var fileToRead = files.First();
                            var trialMessage = new KeyValuePair<uint, MailMessage>(12, InitiateSampleMailMessage());
                            //SaveDataToBizagi(fileToRead, trialMessage.Value.Subject);
                            SaveDataToBizagi(fileToRead, message.Value.Subject);
                            DeleteFile(fileToRead);
                            client.SetMessageFlags(message.Key, null, MessageFlag.Seen);
                        }
                    }
                    if (unreadMessages.Count() == 0)
                    {
                        Console.WriteLine("\nNo New Messages Found ..... Starting IMAP Service\n");
                    }
                    else
                    {
                        Console.Write("\nWoah ....Read All Unread Messages and have stored them ..... Tough job ain't it ... Dont worry .... I get things done :)\n");
                    }
                    #endregion

                    #region ListentoNewMessages
                    while (true)
                    {
                        Console.Write("Initiating IMAP Idle Protocol...");
                        InitializeImapClient();
                        Console.WriteLine("OK");
                        reconnectEvent.WaitOne();
                        if (newMessageSatisfiesCondition)
                        {
                            WriteMessageToFileWIthHelpOfSmtp(msg);
                            var files = ReadAllEmlFiles();
                            if (files.Length > 0)
                            {
                                var fileToRead = files.First();
                                var trialMessage = InitiateSampleMailMessage();
                                //SaveDataToBizagi(fileToRead, trialMessage.Subject);
                                SaveDataToBizagi(fileToRead, msg.Subject);
                                client.SetMessageFlags(newMessageUint, null, MessageFlag.Seen);
                                DeleteFile(fileToRead);
                            }
                        }
                        newMessageSatisfiesCondition = false;
                    }
                    #endregion
                }
                catch (Exception ex)
                {

                }
                finally
                {
                    if (client != null)
                        client.Dispose();
                }
            }
        }

        public static void EmptyTemporatyMailFolder()
        {
            Array.ForEach(Directory.GetFiles(TempMailsStoragePath),
              delegate (string path) { File.Delete(path); });
        }

        public static void DeleteFile(string file)
        {
            if (File.Exists(file))
            {
                File.Delete(file);
            }
        }

        public static void SaveDataToBizagi(string fileNameWithPath, string subject, string body = "")
        {
            string fileBytes = ConvertToBase64(fileNameWithPath);
            string bizagiUrl = ConfigurationManager.AppSettings["BizagiUrl"];
            Connection bizagicon = new Connection(bizagiUrl);
            var fileName = "EmailAttachment.eml";
            List<KeyValuePair<string, string>> keyvaluepair = new List<KeyValuePair<string, string>>() {
                new KeyValuePair<string, string>(bizagiEmailSubjectColumnName, subject),
                new KeyValuePair<string, string>(bizagiEmailBodyCOlumnName, body) };
            WorkflowEngine bizagi = new WorkflowEngine(bizagiUrl);
            var data = bizagi.CreateCase(bizagiUserName, bizagiDomain, bizagiProcessName, bizagiEntityName, keyvaluepair, bizagiEmailFileAttributeName, fileName, fileBytes);
        }

        public static string ConvertToBase64(string file)
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

        public static string[] ReadAllEmlFiles()
        {
            string[] files = System.IO.Directory.GetFiles(TempMailsStoragePath, "*.eml");
            return files;
        }

        public static string FileNameByTime
        {
            get
            {
                return string.Format("{0}.txt", DateTime.Now.ToString("HH-mm-ss"));
            }
        }

        
        static SearchCondition GetSearchCondition()
        {
            return SearchCondition.From(readMessagesFilterAccount).And(SearchCondition.Unseen());
        }
        static List<KeyValuePair<uint, MailMessage>> GetUnreadFilteredMailMessages()
        {
            IEnumerable<uint> uids = client.Search(GetSearchCondition());
            List<KeyValuePair<uint, MailMessage>> messages = uids.Select(x => new KeyValuePair<uint, MailMessage>(x, client.GetMessage(x, false))).ToList();
            return messages;
        }

        static void WriteMessageToFileWIthHelpOfSmtp(MailMessage message)
        {
            SmtpClient client = new SmtpClient();
            client.DeliveryMethod = SmtpDeliveryMethod.SpecifiedPickupDirectory;
            client.PickupDirectoryLocation = TempMailsStoragePath;
            client.Send(message);
        }

        static void InitializeClient()
        {
            // Dispose of existing instance, if any.
            if (client != null)
                client.Dispose();
            client = new ImapClient(host, port, userName, password, AuthMethod.Login, true);
        }

        static void InitializeImapClient()
        {
            // Dispose of existing instance, if any.
            if (client != null)
                client.Dispose();
            client = new ImapClient(host, port, userName, password, AuthMethod.Login, true);
            // Setup event handlers.
            client.NewMessage += client_NewMessage;
            client.IdleError += client_IdleError;
        }

        static void client_IdleError(object sender, IdleErrorEventArgs e)
        {
            Console.Write("An error occurred while idling: ");
            Console.WriteLine(e.Exception.Message);
            reconnectEvent.Set();
        }

        static void client_NewMessage(object sender, IdleMessageEventArgs e)
        {
            msg = client.GetMessage(e.MessageUID, false);
            newMessageUint = e.MessageUID;
            Console.WriteLine("Got a new message, = " + msg.Subject + "--" + msg.Body);
            reconnectEvent.Set();
            if (msg.From.Address == readMessagesFilterAccount)
                newMessageSatisfiesCondition = true;
        }

        static MailMessage InitiateSampleMailMessage()
        {
            MailAddress from = new MailAddress("test@example.com", "TestFromName");
            MailAddress to = new MailAddress("test2@example.com", "TestToName");
            MailMessage myMail = new System.Net.Mail.MailMessage(from, to);

            // add ReplyTo
            MailAddress replyTo = new MailAddress("reply@example.com");
            myMail.ReplyToList.Add(replyTo);

            // set subject and encoding
            myMail.Subject = "Test message";
            myMail.SubjectEncoding = System.Text.Encoding.UTF8;

            // set body-message and encoding
            myMail.Body = "<b>Test Mail</b><br>using <b>HTML</b>.";
            myMail.BodyEncoding = System.Text.Encoding.UTF8;
            // text or html
            myMail.IsBodyHtml = true;

            return myMail;
        }
    }
}


