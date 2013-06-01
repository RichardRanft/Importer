namespace ReelImporter
{
    partial class ImporterRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public ImporterRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

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

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.importGroup = this.Factory.CreateRibbonGroup();
            this.selectButton = this.Factory.CreateRibbonButton();
            this.importButton = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.importGroup.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.importGroup);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // importGroup
            // 
            this.importGroup.Items.Add(this.selectButton);
            this.importGroup.Items.Add(this.importButton);
            this.importGroup.Label = "Reel Import";
            this.importGroup.Name = "importGroup";
            // 
            // selectButton
            // 
            this.selectButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.selectButton.Label = "Choose Folder";
            this.selectButton.Name = "selectButton";
            this.selectButton.OfficeImageId = "OpenFolder";
            this.selectButton.ShowImage = true;
            this.selectButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.selectButton_Click);
            // 
            // importButton
            // 
            this.importButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.importButton.Enabled = false;
            this.importButton.Label = "Import Reels";
            this.importButton.Name = "importButton";
            this.importButton.OfficeImageId = "ImportTextFile";
            this.importButton.ShowImage = true;
            this.importButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.importButton_Click);
            // 
            // ImporterRibbon
            // 
            this.Name = "ImporterRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.ImporterRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.importGroup.ResumeLayout(false);
            this.importGroup.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup importGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton selectButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton importButton;
    }

    partial class ThisRibbonCollection
    {
        internal ImporterRibbon ImporterRibbon
        {
            get { return this.GetRibbon<ImporterRibbon>(); }
        }
    }
}
