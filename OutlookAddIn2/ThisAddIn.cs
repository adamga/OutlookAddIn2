using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace OutlookAddIn2
{
    
    public partial class ThisAddIn
    {
        public Outlook.MailItem  mi = null;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Application.ItemLoad += new Outlook.ApplicationEvents_11_ItemLoadEventHandler(Application_ItemLoad);
                      
        }

        private void Application_ItemLoad(object Item)
        {
            if (Item is Outlook.MailItem)
            {
                mi = Item as Outlook.MailItem;
                mi.BeforeAttachmentAdd += new Outlook.ItemEvents_10_BeforeAttachmentAddEventHandler(Application_BeforeAttach);
               
            }
           
        }


        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see http://go.microsoft.com/fwlink/?LinkId=506785
        }

        void Application_BeforeAttach(Microsoft.Office.Interop.Outlook.Attachment Attachment, ref bool Cancel)
        {

            //add  Watchdox specific dialog and/or functionality call here:

            //get the full path to the selected attachmentment
            string attachname = Attachment.PathName;
            
            //do something with the attachment
            mi.Subject = attachname;

            //cancel the attachment add
            Cancel = true;
            


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
