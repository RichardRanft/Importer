namespace ReelImporter
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
            // FolderSelection
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(527, 103);
            this.Controls.Add(this.folderSelectOK);
            this.Controls.Add(this.folderBrowse);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.selectedFolder);
            this.Name = "FolderSelection";
            this.Text = "Reel Folder Selection";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox selectedFolder;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button folderBrowse;
        private System.Windows.Forms.Button folderSelectOK;
        public System.Windows.Forms.FolderBrowserDialog reelFolderBrowserDialog;
    }
}