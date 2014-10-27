﻿namespace ReelImporter
{
    partial class FolderSelection
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.selectedFolder = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.folderBrowse = new System.Windows.Forms.Button();
            this.folderSelectOK = new System.Windows.Forms.Button();
            this.reelFolderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.rbSHFLReels = new System.Windows.Forms.RadioButton();
            this.rbBallyConfig = new System.Windows.Forms.RadioButton();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // selectedFolder
            // 
            this.selectedFolder.Location = new System.Drawing.Point(13, 41);
            this.selectedFolder.Name = "selectedFolder";
            this.selectedFolder.Size = new System.Drawing.Size(421, 20);
            this.selectedFolder.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(101, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Reel Source Folder:";
            // 
            // folderBrowse
            // 
            this.folderBrowse.Location = new System.Drawing.Point(440, 39);
            this.folderBrowse.Name = "folderBrowse";
            this.folderBrowse.Size = new System.Drawing.Size(75, 23);
            this.folderBrowse.TabIndex = 2;
            this.folderBrowse.Text = "Browse";
            this.folderBrowse.UseVisualStyleBackColor = true;
            this.folderBrowse.Click += new System.EventHandler(this.folderBrowse_Click);
            // 
            // folderSelectOK
            // 
            this.folderSelectOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.folderSelectOK.Location = new System.Drawing.Point(440, 68);
            this.folderSelectOK.Name = "folderSelectOK";
            this.folderSelectOK.Size = new System.Drawing.Size(75, 23);
            this.folderSelectOK.TabIndex = 3;
            this.folderSelectOK.Text = "OK";
            this.folderSelectOK.UseVisualStyleBackColor = true;
            this.folderSelectOK.Click += new System.EventHandler(this.folderSelectOK_Click);
            // 
            // rbSHFLReels
            // 
            this.rbSHFLReels.AutoSize = true;
            this.rbSHFLReels.Checked = true;
            this.rbSHFLReels.Location = new System.Drawing.Point(6, 19);
            this.rbSHFLReels.Name = "rbSHFLReels";
            this.rbSHFLReels.Size = new System.Drawing.Size(82, 17);
            this.rbSHFLReels.TabIndex = 4;
            this.rbSHFLReels.TabStop = true;
            this.rbSHFLReels.Text = "SHFL Reels";
            this.rbSHFLReels.UseVisualStyleBackColor = true;
            // 
            // rbBallyConfig
            // 
            this.rbBallyConfig.AutoSize = true;
            this.rbBallyConfig.Location = new System.Drawing.Point(95, 19);
            this.rbBallyConfig.Name = "rbBallyConfig";
            this.rbBallyConfig.Size = new System.Drawing.Size(80, 17);
            this.rbBallyConfig.TabIndex = 5;
            this.rbBallyConfig.TabStop = true;
            this.rbBallyConfig.Text = "Bally Config";
            this.rbBallyConfig.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.rbSHFLReels);
            this.groupBox1.Controls.Add(this.rbBallyConfig);
            this.groupBox1.Location = new System.Drawing.Point(13, 68);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(421, 46);
            this.groupBox1.TabIndex = 6;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Reel Type";
            // 
            // FolderSelection
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(527, 126);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.folderSelectOK);
            this.Controls.Add(this.folderBrowse);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.selectedFolder);
            this.Name = "FolderSelection";
            this.Text = "Reel Folder Selection";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox selectedFolder;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button folderBrowse;
        private System.Windows.Forms.Button folderSelectOK;
        public System.Windows.Forms.FolderBrowserDialog reelFolderBrowserDialog;
        public void setSelectedFolder(string folder)
        {
            selectedFolder.Text = folder;
            reelFolderBrowserDialog.SelectedPath = folder;
        }

        private System.Windows.Forms.RadioButton rbSHFLReels;
        private System.Windows.Forms.RadioButton rbBallyConfig;
        private System.Windows.Forms.GroupBox groupBox1;
    }
}