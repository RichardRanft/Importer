using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using ReelImporter.Properties;

namespace ReelImporter
{
    public partial class ImporterRibbon
    {
        private FolderSelection select;
        private HeaderImport import;
        private Excel.Window excelWin;
        private void ImporterRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            this.select = new FolderSelection();
            this.select.ribbon = this;
            this.excelWin = Globals.Program.Application.ActiveWindow;
            this.import = new HeaderImport(excelWin);
        }

        private void selectButton_Click(object sender, RibbonControlEventArgs e)
        {
            select.ShowDialog();
        }

        private void importButton_Click(object sender, RibbonControlEventArgs e)
        {
            import.importFolder(select.importFolder);
        }

        public void EnableImport(bool flag)
        {
            this.importButton.Enabled = flag;
        }
    }
}
