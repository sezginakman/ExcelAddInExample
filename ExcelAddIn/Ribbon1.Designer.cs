namespace ExcelAddIn
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
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
            this.group1 = this.Factory.CreateRibbonGroup();
            this.hideButton = this.Factory.CreateRibbonButton();
            this.showButton = this.Factory.CreateRibbonButton();
            this.trimButton = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.showButton);
            this.group1.Items.Add(this.hideButton);
            this.group1.Items.Add(this.trimButton);
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // hideButton
            // 
            this.hideButton.Label = "Hide Comments";
            this.hideButton.Name = "hideButton";
            this.hideButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.hideButton_Click);
            // 
            // showButton
            // 
            this.showButton.Label = "Show Comments";
            this.showButton.Name = "showButton";
            this.showButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.showButton_Click);
            // 
            // trimButton
            // 
            this.trimButton.Label = "Auto Size Comments";
            this.trimButton.Name = "trimButton";
            this.trimButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.trimButton_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton hideButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton showButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton trimButton;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
