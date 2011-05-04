/* WRTOffsiteTaglineAddinRibbon.cs        */
/* Created by Larry G. Wapnitsky    */
/* August, 2010                     */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

namespace WRTOffsite_NET35
{
    public partial class WRTOffsiteTaglineAddInRibbon_old
    {
        string taglineActive;
        OLRegistryAddin buttonSet = new OLRegistryAddin();  // variable for reading the value of the registry key
        UpdateBody msgBody = new UpdateBody();  // method for adding/removing tagline from the message

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            taglineActive = buttonSet.RegCurrentValue();  // retrieve the current registry value

            if (taglineActive == "0")
            {
                // tagline is off for all messages
                ActiveAllMessages.Checked = false; // uncheck "All Messages" button
                ActiveAllMessages.Label = "Inactive - All Messages";  // change the label
                ActiveThisMessage.Visible = false;  // hide the "This Message" button
                ActiveThisMessage.Enabled = false;  // deactivate the "This Message" button
            }
            else if (taglineActive == "1")
            {
                // tagline is on for all messages
                ActiveAllMessages.Checked = true;   // check "All Messages" button
                ActiveAllMessages.Label = "Active - All Messages";  // change the label
                ActiveThisMessage.Visible = true;   // show the "This Message" button
                ActiveThisMessage.Enabled = true;   // activate the "This Message" button
            }
        }

        private void Tagline()
        {
            Outlook.Inspector inspector = this.Context as Outlook.Inspector;
            Outlook.MailItem mi = inspector.CurrentItem as Outlook.MailItem;
            msgBody.updateTask(mi, taglineActive);  // update the message body based on the value of taglineActive
        }

        private void ActiveAllMessages_Click(object sender, RibbonControlEventArgs e)
        {
            switch (ActiveAllMessages.Checked)
            {
                case true:
                    taglineActive = "1";                // tagline is on for all messages
                    ActiveAllMessages.Checked = true;   // check "All Messages" button
                    ActiveAllMessages.Label = "Active - All Messages";  // change the label
                    ActiveThisMessage.Visible = true;   // show the "This Message" button
                    ActiveThisMessage.Enabled = true;   // activate the "This Message" button
                    break;
                case false:
                    taglineActive = "0";                // tagline is off for all messages
                    ActiveAllMessages.Checked = false;  // uncheck "All Messages" button
                    ActiveAllMessages.Label = "Inactive - All Messages";  // change the label
                    ActiveThisMessage.Visible = false;  // hide the "This Message" button
                    ActiveThisMessage.Enabled = false;  // deactivate the "This Message" button
                    break;
            }
            buttonSet.SetCurrentValue(taglineActive);  

            Tagline();
        }

        private void ActiveThisMessage_Click(object sender, RibbonControlEventArgs e)
        {
            switch (ActiveThisMessage.Checked)
            {
                case true:
                    taglineActive = "1";
                    ActiveThisMessage.Label = "Active - This Message Only";
                    break;
                case false:
                    taglineActive = "0";
                    ActiveThisMessage.Label = "Inactive - This Message Only";
                    break;
            }

            Tagline();
        }


    }
}