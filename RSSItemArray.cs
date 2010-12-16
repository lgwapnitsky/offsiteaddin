/* RSSItemArray.cs        */
/* Created by Larry G. Wapnitsky    */
/* August, 2010                     */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Net;
using System.Xml;

namespace GenerateArrayFromRSSXML
{
    public class RSSItemArray
    {
        private RSSItem[] offsite;  // array of RSSItems using our custom structure.
        int NUMBER_OF_RSS_ITEMS;
        string EnvTempDir = Environment.GetEnvironmentVariable("Temp");
        string OffsiteXMLFile = "offsite.xml";
        string LocalXMLFile;
        
        public RSSItemArray()
        {
            XmlTextReader rssReader;
            XmlDocument rssDoc;
            XmlNode nodeRss = null;
            XmlNode nodeChannel = null;
            XmlNode nodeItem;

            // define the location of the RSS XML file
            string OffsiteXMLDir = EnvTempDir.Replace("\\", "\\\\");
            LocalXMLFile = OffsiteXMLDir + "\\" + OffsiteXMLFile;

            // open the XML file and read in all values
            rssReader = new XmlTextReader(LocalXMLFile);
            rssDoc = new XmlDocument();
            rssDoc.Load(rssReader);

            // Populate the array of RSS items by finding valid RSS items...
            for (int i = 0; i < rssDoc.ChildNodes.Count; i++)
            {
                if (rssDoc.ChildNodes[i].Name == "rss")
                {
                    nodeRss = rssDoc.ChildNodes[i];
                }
            }

            // ...that have valid child names
            for (int i = 0; i < nodeRss.ChildNodes.Count; i++)
            {
                if (nodeRss.ChildNodes[i].Name == "channel")
                {
                    nodeChannel = nodeRss.ChildNodes[i];
                }
            }

            NUMBER_OF_RSS_ITEMS = nodeChannel.ChildNodes.Count;

            offsite = new RSSItem[NUMBER_OF_RSS_ITEMS];  // Resize the array to hold all valid items

            for (int i = 0; i < offsite.Length; i++)
            {
                if (nodeChannel.ChildNodes[i].Name == "item")
                {
                    nodeItem = nodeChannel.ChildNodes[i];
                    string title = nodeItem["title"].InnerText;

                    if (title != null)  // but do not include items that have no title
                    {
                        offsite[i] = new RSSItem(title, nodeItem["description"].InnerText, nodeItem["link"].InnerText);
                        // each new item consists of a title, description and URL
                    }
                }
            }
        }

        public RSSItem PickRssItem()
        {
            // generate a random number within the range of total RSS items and choose a random item
            Random r = new Random();
            int RandomItem = r.Next(NUMBER_OF_RSS_ITEMS);
            return offsite[RandomItem];
        }
    }
}
