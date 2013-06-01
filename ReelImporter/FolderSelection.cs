using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ReelImporter
{
    public partial class FolderSelection : Form
    {
        public ImporterRibbon ribbon;
        public String importFolder;
        public FolderSelection()
        {
            InitializeComponent();
        }

        public FolderBrowserDialog getSelectDialog()
        {
            return reelFolderBrowserDialog;
        }

        private void folderBrowse_Click(object sender, EventArgs e)
        {
            if (reelFolderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                selectedFolder.Text = reelFolderBrowserDialog.SelectedPath;
                importFolder = selectedFolder.Text;
            }
        }

        private void folderSelectOK_Click(object sender, EventArgs e)
        {
            if (selectedFolder.Text != "")
                ribbon.EnableImport(true);
            else
                ribbon.EnableImport(false);
            this.Close();
        }
    }
}
