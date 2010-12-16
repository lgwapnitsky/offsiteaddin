/* RSSItem.cs        */
/* Created by Larry G. Wapnitsky    */
/* August, 2010                     */


// This is a class that defines the structure of an item from the RSS XML file

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace GenerateArrayFromRSSXML
{
    public class RSSItem
    {
        private string title;
        private string description;
        private string link;

        public RSSItem(string offsiteTitle, string offsiteDescription, string offsiteLink)
        {
            title = offsiteTitle;
            description = offsiteDescription;
            link = offsiteLink;
        }

        public override string ToString()
        {
            return title + "*" + description + "*" + link;
            // return a string that has the variables separated by a '*' for parsing later on
        }
    }
}
