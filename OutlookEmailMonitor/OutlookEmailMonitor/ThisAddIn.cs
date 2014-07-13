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

namespace OutlookEmailMonitor
{
    public partial class ThisAddIn
    {
        Outlook.Inspectors inspectors;
        Outlook.NameSpace outlook_namespace;
        Outlook.MAPIFolder inbox;
        Outlook.Items items;
        Storage storage;
        int most_recent_mail_hashcode;
        
        private async void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //inspectors = this.Application.Inspectors;
            //inspectors.NewInspector += new Microsoft.Office.Interop.Outlook.InspectorsEvents_NewInspectorEventHandler(Inspectors_NewInspector);
            //this.Application.NewMail += new Microsoft.Office.Interop.Outlook.MailItem;

            outlook_namespace = this.Application.GetNamespace("MAPI");
            inbox = outlook_namespace.GetDefaultFolder(
                Outlook.OlDefaultFolders.olFolderInbox);

            items = inbox.Items;
            items.ItemAdd +=
                new Outlook.ItemsEvents_ItemAddEventHandler(items_ItemAdd);

            await ProcessUnprocessedEmails(); 
            Debug.WriteLine("Hello debugger - this is the application speaking! We finished processing the unprocessed emails");
        }

        private async Task ProcessUnprocessedEmails(){
            object item = items.GetNext();
            int i = 0;
            while (isUnprocessed(item) && i != 30000)
            {
                Outlook.MailItem mail_item = item as Outlook.MailItem;
                if (mail_item != null)
                {
                    DateTime received_time  = mail_item.ReceivedTime;
                    DateTime sent_time      = mail_item.SentOn;
                    TimeSpan latency        = received_time.Subtract(sent_time);
                    //Debug.WriteLine(received_time.ToString());
                    //Debug.WriteLine(sent_time.ToString());
                    Debug.WriteLine("Latency: " + latency.ToString());
                    //Debug.WriteLine(mail_item);
                }
                item = items.GetNext();
                ++i;
            }
        }

        private bool isUnprocessed(object item)
        {
            Outlook.MailItem mail_item = item as Outlook.MailItem;
            if (mail_item != null)
            {
                int hash = mail_item.GetHashCode();
            }
            return false;
            
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
            }
        }
        void Inspectors_NewInspector(Microsoft.Office.Interop.Outlook.Inspector Inspector)
        {
            Outlook.MailItem mailItem = Inspector.CurrentItem as Outlook.MailItem;
            if (mailItem != null)
            {
                if (mailItem.EntryID == null)
                {
                    //mailItem.Subject = "This text was added by using code";
                    //mailItem.Body = "This text was added by using code";
                }

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
