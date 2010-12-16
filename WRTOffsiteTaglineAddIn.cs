/* WRTOffsiteTaglineAddin.cs        */
/* Created by Larry G. Wapnitsky    */
/* August, 2010                     */


using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using System.Net;
using System.IO;
using System.Xml;
using Microsoft.VisualStudio.Tools.Office.Runtime.Interop;



namespace WRTOffsite_NET35
{
    public partial class WRTOffsiteTaglineAddIn
    {
        // Define variables for local and remote XML file from which to read taglines/links
        string EnvTempDir = Environment.GetEnvironmentVariable("Temp");
        string OffsiteXMLFile = "offsite.xml";
        string LocalXMLFile;
        string urlOffsiteRss = "http://www.wrtdesign.com/offsite/rss.xml";
        
        // Create a new instance of class that checks for/creates/modifies registry entries
        OLRegistryAddin olRegCheck = new OLRegistryAddin();
        
        // create inspectors to monitor Outlook window activity
        Outlook.Inspectors inspectors;
        Outlook.Inspector inspector;


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            olRegCheck.RegCheckExists();  // Check to see if the registry entry for the add-in exists.  If not, create it.

            string OffsiteXMLDir = EnvTempDir.Replace("\\", "\\\\");
            LocalXMLFile = OffsiteXMLDir + "\\" + OffsiteXMLFile;

            GetXMLFile();  // Download the RSS XML file from the 'offsite' website

            // Activate this add-in on a message window
            inspectors = Application.Inspectors;
            inspectors.NewInspector += new Outlook.InspectorsEvents_NewInspectorEventHandler(Inspectors_NewInspector);
        }

        public void Inspectors_NewInspector(Outlook.Inspector Inspector)
        {
            if (inspector == null)
            {
                inspector = Inspector;
                ((Outlook.InspectorEvents_10_Event)inspector).Activate += new Outlook.InspectorEvents_10_ActivateEventHandler(Inspector_Activate);
                ((Outlook.InspectorEvents_10_Event)inspector).Close += new Outlook.InspectorEvents_10_CloseEventHandler(Inspector_Close);
            }
        }

        void Inspector_Close()
        {
            ((Outlook.InspectorEvents_10_Event)inspector).Activate -= new Outlook.InspectorEvents_10_ActivateEventHandler(Inspector_Activate);
            ((Outlook.InspectorEvents_10_Event)inspector).Close -= new Outlook.InspectorEvents_10_CloseEventHandler(Inspector_Close);
            inspector = null;
        }

        public void Inspector_Activate()
        {
            Outlook.MailItem mi = inspector.CurrentItem as Outlook.MailItem;
            
            if (mi.Sent == false)  // activate the butttons on New, Reply and resumed Draft messages
            {
                OLRegistryAddin buttonSet = new OLRegistryAddin();
                                
                UpdateBody olMessage = new UpdateBody();
                olMessage.RemoveOffsiteMessage(mi);
                olMessage.updateTask(mi, buttonSet.RegCurrentValue());
            }
            
            Inspector_Close();
        }

        public void GetXMLFile()
        {
            System.Net.WebClient Client = new WebClient();
            try
            {
                do
                {
                    Client.DownloadFile(urlOffsiteRss, LocalXMLFile);
                }
                while (Client == null);
            }
            catch { };
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
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
