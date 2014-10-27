using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.Win32;
using ReelImporter.Properties;

namespace ReelImporter
{
    public partial class ImporterRibbon
    {
        private FolderSelection select;
        private HeaderImport import;
        private CalcImport calc;
        private Excel.Window excelWin;
        private System.Windows.Forms.OpenFileDialog openFile;

        private void ImporterRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            this.select = new FolderSelection();
            this.select.ribbon = this;
            this.excelWin = Globals.Program.Application.ActiveWindow;
            this.import = new HeaderImport(excelWin);
            this.calc = new CalcImport(this);
            this.openFile = new System.Windows.Forms.OpenFileDialog();
            this.openFile.Filter = "Excel files(*.xls;*.xlsx;*.xlsm)|*.xls;*.xlsx;*.xlsm";
        }

        private void selectButton_Click(object sender, RibbonControlEventArgs e)
        {
            Object val = Globals.Program.currUserKey.GetValue("Folder", "");
            if (val != null && val.GetType() == this.openFile.Filter.GetType())
                select.setSelectedFolder(val.ToString());
            select.ShowDialog();
        }

        private void importButton_Click(object sender, RibbonControlEventArgs e)
        {
            import.importFolder(select.importFolder, select.reelType);
        }

        public void EnableImport(bool flag)
        {
            this.importButton.Enabled = flag;
            this.selectCalcsButton.Enabled = flag;
            if (!flag)
                this.calcsButton.Enabled = flag;
        }

        private void selectCalcsButton_Click(object sender, RibbonControlEventArgs e)
        {
            if (openFile.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                calc.setTargetWorkbook(Globals.Program.Application.ActiveWorkbook);
                calc.setTargetWorksheet(Globals.Program.Application.ActiveSheet);
                calc.openWorkbook(openFile.FileName);
                EnableCalcImport(true);
            }
        }

        public void EnableCalcImport(bool flag)
        {
            this.calcsButton.Enabled = flag;
        }

        private void calcsButton_Click(object sender, RibbonControlEventArgs e)
        {
            calc.setStartCell(Globals.Program.Application.Selection);
            calc.importReels();
        }
    }
}
