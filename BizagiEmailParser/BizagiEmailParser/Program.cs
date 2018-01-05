using S22.Imap;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
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
        static string mailSavePath = ConfigurationManager.AppSettings["SaveMailPath"];
        static string bizagiUrl = ConfigurationManager.AppSettings["BizagiUrl"];
        static string TempStorePathForMailMessage = ConfigurationManager.AppSettings["BizagiUrl"];
        static string bizagiUserName = "admon";
        static string bizagiDomain = "domain";
        static string bizagiProcessName = "trial2";
        static string bizagiEntityName = "trial2";
        static string bizagiEmailSubjectColumnName = "ssubject";
        static string bizagiEmailBodyCOlumnName = "sbody";
        static string bizagiEmailFileAttributeName = "ffileAttribute";
        static string readMessagesFilterAccount = "abhiram144tests@gmail.com";
        static MailMessage msg;

        static void Main(string[] args)
        {
            try
            {
                Console.Write("Connecting...");
                InitializeClient();
                var unreadMessages = GetUnreadFilteredMailMessages();
                foreach (var message in unreadMessages)
                {
                    WriteMessageToFileWIthHelpOfSmtp(message);
                }
                if(unreadMessages.Count() == 0)
                {
                    Console.WriteLine("No New Messages Found ..... Starting IMAP Service");
                }
                Console.Write("Read All Unread Messages ... Done");
            }
            finally
            {
                if (client != null)
                    client.Dispose();
            }
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

        static IEnumerable<MailMessage> GetUnreadFilteredMailMessages()
        {
            IEnumerable<uint> uids = client.Search(SearchCondition.From(readMessagesFilterAccount).And(SearchCondition.Unseen()));
            IEnumerable<MailMessage> messages = client.GetMessages(uids);
            return messages;
        }

        static void WriteMessageToFileWIthHelpOfSmtp(MailMessage message)
        {
            SmtpClient client = new SmtpClient();
            client.DeliveryMethod = SmtpDeliveryMethod.SpecifiedPickupDirectory;
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


