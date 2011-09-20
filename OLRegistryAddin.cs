/* OLRegistryAddin.cs               */
/* Created by Larry G. Wapnitsky    */
/* August, 2010                     */

using Microsoft.Win32;

namespace WRTOffsite_NET35
{
    internal class OLRegistryAddin
    {
        RegistryKey olAddinKey = Registry.CurrentUser;
        string OLAddinSubKey = @"Software\WRT\OutlookAddins\Offsite";
        string OLAddinValue = "OffsiteActive";

        public void RegCheckExists()  // Check to see if the registry key exists.  If not, create it and set as active
        {
            olAddinKey = olAddinKey.OpenSubKey(OLAddinSubKey);
            if (olAddinKey == null)
            {
                olAddinKey = Registry.CurrentUser.CreateSubKey(OLAddinSubKey);
                olAddinKey.SetValue(OLAddinValue, "1");
            }
        }

        public string RegCurrentValue()  // Retrieve the current value from the registry key
        {
            olAddinKey = olAddinKey.OpenSubKey(OLAddinSubKey, true);
            string currentValue = olAddinKey.GetValue(OLAddinValue).ToString() as string;

            return currentValue;
        }

        public void SetCurrentValue(string value)  // Set the value of the registry key
        {
            olAddinKey.SetValue(OLAddinValue, value);
        }
    }
}