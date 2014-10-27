using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.Win32;

namespace ReelImporter
{
    public partial class Program
    {
        public RegistryKey currUserKey;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                if (Registry.CurrentUser != null)
                    this.currUserKey = Registry.CurrentUser;
            }
            catch (Exception rException)
            {
                System.Windows.Forms.MessageBox.Show("Error code:\n" + rException.Message.ToString(), "Error Accessing Registry", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
            validateKey(currUserKey);
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
        private void validateKey(RegistryKey key)
        {
            String[] names = key.GetSubKeyNames();
            bool found = false;
            foreach (String s in names)
            {
                if (s == "Software\\Microsoft\\Office\\Excel\\Addins\\ReelImporter\\CurrentFolder")
                    found = true;
            }
            if (!found)
            {
                key.CreateSubKey("Software\\Microsoft\\Office\\Excel\\Addins\\ReelImporter\\CurrentFolder", RegistryKeyPermissionCheck.ReadWriteSubTree);
            }
        }
    }

}