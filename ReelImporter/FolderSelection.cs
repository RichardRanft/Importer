using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;

namespace ReelImporter
{
    public partial class FolderSelection : Form
    {
        public ImporterRibbon ribbon;
        public String importFolder;
        public ReelDataType reelType;
        public FolderSelection()
        {
            InitializeComponent();
            importFolder = "";
            reelType = 0;
        }

        public FolderBrowserDialog getSelectDialog()
        {
            return reelFolderBrowserDialog;
        }

        private void folderBrowse_Click(object sender, EventArgs e)
        {
            // HKEY_CURRENT_USER\Software\Microsoft\Office\Excel\Addins\ReelImporter
            if (reelFolderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                selectedFolder.Text = reelFolderBrowserDialog.SelectedPath;
                importFolder = selectedFolder.Text;
            }
        }

        private void folderSelectOK_Click(object sender, EventArgs e)
        {
            if (selectedFolder.Text != "")
            {
                ribbon.EnableImport(true);
                importFolder = selectedFolder.Text;
                if (rbSHFLReels.Checked)
                    reelType = ReelDataType.SHFL ;
                if (rbBallyConfig.Checked)
                    reelType = ReelDataType.BALLY;
                Globals.Program.currUserKey.SetValue("Folder", selectedFolder.Text);
            }
            else
                ribbon.EnableImport(false);
            this.Close();
        }
    }
}
