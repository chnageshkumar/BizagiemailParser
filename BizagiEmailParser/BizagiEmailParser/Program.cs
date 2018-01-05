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
        static MailMessage msg;

        static void Main(string[] args)
        {
            try
            {
                Console.Write("Connecting...");
                InitializeClient();
                Console.Write("Ok");
                var unreadMessages = GetUnreadFilteredMailMessages();
                foreach (var message in unreadMessages)
                {
                    WriteMessageToFileWIthHelpOfSmtp(message.Value);
                    var files = ReadAllEmlFiles();
                    if (files.Length > 0)
                    {
                        var fileToRead = files.First();
                        FileStream fs = File.Open(fileToRead, FileMode.Open,
                             FileAccess.Read);
                        EMLReader reader = new EMLReader(fs);
                        SaveDataToBizagi(fileToRead, message.Value.Subject);
                        fs.Close();
                    }
                    if (unreadMessages.Count() == 0)
                    {
                        Console.WriteLine("\nNo New Messages Found ..... Starting IMAP Service\n");
                    }
                    else
                    {
                        Console.Write("\nRead All Unread Messages ... Done\n");
                    }
                }
            }
            catch(Exception ex)
            {

            }
            finally
            {
                if (client != null)
                    client.Dispose();
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

        static ImapClient client;

        static void ListenToNewMessages()
        {
            while (true)
            {
                Console.Write("Connecting...");
                InitializeClient();
                Console.WriteLine("OK");
                //reconnectEvent.WaitOne();
                //WriteMessage(msg);
            }
        }

        static IEnumerable<KeyValuePair<uint, MailMessage>> GetUnreadFilteredMailMessages()
        {
            IEnumerable<uint> uids = client.Search(SearchCondition.From(readMessagesFilterAccount).And(SearchCondition.Unseen()));
            IEnumerable<KeyValuePair<uint, MailMessage>> messages = uids.Select(x => new KeyValuePair<uint, MailMessage>(x, client.GetMessage(x)));
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
            msg = client.GetMessage(e.MessageUID);
            Console.WriteLine("Got a new message, = " + msg.Subject + "--" + msg.Body);
            reconnectEvent.Set();
        }
    }
}


