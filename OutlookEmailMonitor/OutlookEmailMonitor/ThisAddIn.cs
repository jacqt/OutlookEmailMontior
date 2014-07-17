using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Resources;
using System.Threading;
using Newtonsoft.Json;

namespace OutlookEmailMonitor
{
    public class Data : List<DataMember>
    {
    }

    public class DataMember : Tuple<DateTime, TimeSpan>
    {
        public DataMember(DateTime a, TimeSpan b) : base (a,b)
        {
        }
    }

    public class Parameter : Tuple<String,String>
    {
        public Parameter(String a, String b) : base (a,b)
        {
        }
           
    }

    public partial class ThisAddIn
    {
        Outlook.NameSpace outlook_namespace;
        Outlook.MAPIFolder inbox;
        Outlook.Items items;

        ServerCommunication client;
        String app_guid;
        String username;
        String domain;
        String email;
        int most_recent_mail_hashcode;

        private const String url = "http://107.170.81.203:8080/send_data";

        
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //inspectors = this.Application.Inspectors;
            //inspectors.NewInspector += new Microsoft.Office.Interop.Outlook.InspectorsEvents_NewInspectorEventHandler(Inspectors_NewInspector);
            //this.Application.NewMail += new Microsoft.Office.Interop.Outlook.MailItem;

            client = new ServerCommunication(url);

            //Load guid if exists. Otherwise create a new one.
            String app_guid = Storage.loadFile("app_guid");

            //Need a check to make sure that it is an exchange user
            Outlook.ExchangeUser current_user = Application.Session.CurrentUser.AddressEntry.GetExchangeUser();
            email = current_user.PrimarySmtpAddress;
            String[] temp = email.Split('@');

            username = temp[0];
            domain = temp[1];
            
            if (app_guid == null)
            {
                app_guid = Guid.NewGuid().ToString();
                Storage.saveFile("app_guid", app_guid);
            }

            outlook_namespace = this.Application.GetNamespace("MAPI");
            inbox = outlook_namespace.GetDefaultFolder(
                Outlook.OlDefaultFolders.olFolderInbox);

            items = inbox.Items;
            items.ItemAdd +=
                new Outlook.ItemsEvents_ItemAddEventHandler(items_ItemAdd);

#if DEBUG
            //Storage.saveFile("most_recent_email_hash", "foo"); //Force outlook to load everything again
#endif
            Thread process_emails_thread = new Thread(ProcessUnprocessedEmails);
            process_emails_thread.Start();
            Debug.WriteLine("Hello debugger - this is the application speaking!");
        }

        private void ProcessUnprocessedEmails()
        {
            object item = items.GetLast();
            String most_recent_email_hash = item.GetHashCode().ToString();
            Debug.WriteLine(most_recent_email_hash);

            Data data = new Data();
            int i = 0;
            while (isUnprocessed(item))
            {
                Outlook.MailItem mail_item = item as Outlook.MailItem;
                if (mail_item != null)
                {
                    DateTime received_time  = mail_item.ReceivedTime;
                    DateTime sent_time      = mail_item.SentOn;
                    TimeSpan latency        = received_time.Subtract(sent_time);
                    //Debug.WriteLine(received_time.ToString());
                    Debug.WriteLine(sent_time.ToString());
                    //Debug.WriteLine(mail_item);

                    //Debug.WriteLine("Latency: " + latency.ToString());
                    data.Add(new DataMember(sent_time, latency));
                }
                item = items.GetPrevious();
                ++i;
            }
            Storage.saveFile("most_recent_email_hash", most_recent_email_hash);
            sendData(data);
        }

        private bool isUnprocessed(object item)
        {
            if (item == null)
                return false;
            int hash = item.GetHashCode();
            Debug.WriteLine("Hashcode: " + item.GetHashCode().ToString());
            String most_recent_email_hash = Storage.loadFile("most_recent_email_hash");
            if (most_recent_email_hash != null)
            {
                if (most_recent_email_hash.Equals(item.GetHashCode().ToString()))
                {
                    return false;
                }
            }
            return true;
        }

        private void sendData(Data data)
        {
            String data_serialized = Newtonsoft.Json.JsonConvert.SerializeObject(data);
            Debug.WriteLine(data_serialized);

            List<Parameter> parameters = new List<Parameter>();
            parameters.Add(new Parameter("app_guid", app_guid));
            parameters.Add(new Parameter("username", username));
            parameters.Add(new Parameter("domain", domain));
            parameters.Add(new Parameter("data", data_serialized));
            try
            {
                client.performPostRequest(parameters);
            }
            catch (Exception e)
            {
                //TODO
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            //TODO
        }

        void items_ItemAdd(object item)
        {
            Outlook.MailItem mail_item = item as Outlook.MailItem;
            if (mail_item != null)
            {
                DateTime received_time  = mail_item.ReceivedTime;
                DateTime sent_time      = mail_item.SentOn;
                TimeSpan latency        = received_time.Subtract(sent_time);
                Debug.WriteLine(received_time.ToString());
                Debug.WriteLine(sent_time.ToString());
                Debug.WriteLine(latency.ToString());
                Debug.WriteLine(mail_item);
                Storage.saveFile("most_recent_email_hash", mail_item.GetHashCode().ToString());

                Data data = new Data();
                data.Add(new DataMember(sent_time, latency));
                sendData(data);
                    //Debug.WriteLine("Latency: " + latency.ToString());
            }
        }
        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
