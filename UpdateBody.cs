/* UpdateBody.cs        */
/* Created by Larry G. Wapnitsky    */
/* August, 2010                     */

using System;
using System.Linq;
using GenerateArrayFromRSSXML;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;

namespace WRTOffsite_NET35
{
    internal class UpdateBody
    {
        public void updateTask(Outlook.MailItem oMsg, string taglineActive)
        {
            if (taglineActive == "1")
            {
                // if the registry key and current item are active, add the tagline
                InsertOffsiteSig(oMsg);
            }
            else
            {
                // otherwise remove it from the message if it exists
                RemoveOffsiteMessage(oMsg);
            }
        }

        private void InsertOffsiteSig(Outlook.MailItem oMsg)
        {
            object oBookmarkName = "_MailAutoSig";  // Outlook internal bookmark for location of the e-mail signature
            string oOffsiteBookmark = "OffsiteBookmark";  // bookmark to be created in Outlook for the Offsite tagline
            object oOffsiteBookmarkObj = oOffsiteBookmark;

            Word.Document SigDoc = oMsg.GetInspector.WordEditor as Word.Document; // edit the message using Word

            string bf = oMsg.BodyFormat.ToString();  // determine the message body format (text, html, rtf)

            //  Go to the e-mail signature bookmark, then set the cursor to the very end of the range.
            //  This is where we will insert/remove our tagline, and the start of the new range of text

            Word.Range r = SigDoc.Bookmarks.get_Item(ref oBookmarkName).Range;
            object collapseEnd = Word.WdCollapseDirection.wdCollapseEnd;

            r.Collapse(ref collapseEnd);

            string[] taglines = GetRssItem();  // Get tagline information from the RSS XML file and place into an array

            // Loop through the array and insert each line of text separated by a newline

            foreach (string taglineText in taglines)
                r.InsertAfter(taglineText + "\n");
            r.InsertAfter("\n");

            // Add formatting to HTML/RTF messages

            if (bf != "olFormatPlain" && bf != "olFormatUnspecified")
            {
                SigDoc.Hyperlinks.Add(r, taglines[2]); // turn the link text into a hyperlink
                r.Font.Underline = 0;  // remove the hyperlink underline
                r.Font.Color = Word.WdColor.wdColorGray45;  // change all text to Gray45
                r.Font.Size = 8;  // Change the font size to 8 point
                r.Font.Name = "Arial";  // Change the font to Arial
            }

            r.NoProofing = -1;  // turn off spelling/grammar check for this range of text

            object range1 = r;
            SigDoc.Bookmarks.Add(oOffsiteBookmark, ref range1);  // define this range as our custom bookmark

            if (bf != "olFormatPlain" && bf != "olFormatUnspecified")
            {
                // Make the first line BOLD only for HTML/RTF messages

                Word.Find f = r.Find;
                f.Text = taglines[0];
                f.MatchWholeWord = true;
                f.Execute();
                while (f.Found)
                {
                    r.Font.Bold = -1;
                    f.Execute();
                }
            }
            else
            {
                // otherwise turn the plain text hyperlink into an active hyperlink
                // this is done here instead of above due to the extra formatting needed for HTML/RTF text

                Word.Find f = r.Find;
                f.Text = taglines[2];
                f.MatchWholeWord = true;
                f.Execute();
                SigDoc.Hyperlinks.Add(r, taglines[2]);
            }
            r.NoProofing = -1;  // disable spelling/grammar checking on the updated range
            r.Collapse(collapseEnd);
        }

        public void RemoveOffsiteMessage(Outlook.MailItem oMsg)
        {
            string oOffsiteBookmark = "OffsiteBookmark";
            object oOffsiteBookmarkObj = oOffsiteBookmark;

            Word.Document SigDoc = oMsg.GetInspector.WordEditor as Word.Document;

            if (SigDoc.Bookmarks.Exists(oOffsiteBookmark) == true)  // if the custom bookmark exists, remove it
            {
                Word.Range r = SigDoc.Bookmarks.get_Item(ref oOffsiteBookmarkObj).Range;
                r.Text = "";
            }
        }

        private string[] GetRssItem()
        {
            string[] oi2;

            RSSItemArray offsiteTaglines = new RSSItemArray();
            do
            {
                string oi = String.Format("{0}", offsiteTaglines.PickRssItem());  // get a random RSS item as a string

                oi2 = oi.Split('*');  // split it into an array using '*' as a delimiter
            } while (oi2[0] == null || oi2.Count() < 3);

            return oi2;
        }
    }
}