namespace GeollyExcelWorkbook
{
    partial class GeollyRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public GeollyRibbon()
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
            this.GeollyTab = this.Factory.CreateRibbonTab();
            this.GeollyGroup = this.Factory.CreateRibbonGroup();
            this.ResetTTAPIButton = this.Factory.CreateRibbonButton();
            this.StopTTAPIButton = this.Factory.CreateRibbonButton();
            this.buttonGroup1 = this.Factory.CreateRibbonButtonGroup();
            this.BuyButton = this.Factory.CreateRibbonButton();
            this.SellButton = this.Factory.CreateRibbonButton();
            this.buttonGroup2 = this.Factory.CreateRibbonButtonGroup();
            this.TakeProfitButton = this.Factory.CreateRibbonButton();
            this.CutLossButton = this.Factory.CreateRibbonButton();
            this.GeollyTab.SuspendLayout();
            this.GeollyGroup.SuspendLayout();
            this.buttonGroup1.SuspendLayout();
            this.buttonGroup2.SuspendLayout();
            // 
            // GeollyTab
            // 
            this.GeollyTab.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.GeollyTab.Groups.Add(this.GeollyGroup);
            this.GeollyTab.Label = "Geolly Pty Ltd";
            this.GeollyTab.Name = "GeollyTab";
            // 
            // GeollyGroup
            // 
            this.GeollyGroup.Items.Add(this.ResetTTAPIButton);
            this.GeollyGroup.Items.Add(this.StopTTAPIButton);
            this.GeollyGroup.Items.Add(this.buttonGroup1);
            this.GeollyGroup.Items.Add(this.buttonGroup2);
            this.GeollyGroup.Label = "Geolly Pty Ltd";
            this.GeollyGroup.Name = "GeollyGroup";
            // 
            // ResetTTAPIButton
            // 
            this.ResetTTAPIButton.Label = "Reset TT API";
            this.ResetTTAPIButton.Name = "ResetTTAPIButton";
            this.ResetTTAPIButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Reset_TT_API_Button_Click);
            // 
            // StopTTAPIButton
            // 
            this.StopTTAPIButton.Label = "Stop TT API";
            this.StopTTAPIButton.Name = "StopTTAPIButton";
            this.StopTTAPIButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.STOP_TT_API_Button_Click);
            // 
            // buttonGroup1
            // 
            this.buttonGroup1.Items.Add(this.BuyButton);
            this.buttonGroup1.Items.Add(this.SellButton);
            this.buttonGroup1.Name = "buttonGroup1";
            // 
            // BuyButton
            // 
            this.BuyButton.Label = "Buy";
            this.BuyButton.Name = "BuyButton";
            this.BuyButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BuyButton_Click);
            // 
            // SellButton
            // 
            this.SellButton.Label = "Sell";
            this.SellButton.Name = "SellButton";
            this.SellButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SellButton_Click);

            // 
            // buttonGroup2
            // 
            this.buttonGroup2.Items.Add(this.TakeProfitButton);
            this.buttonGroup2.Items.Add(this.CutLossButton);
            this.buttonGroup2.Name = "buttonGroup2";
            // 
            // BuyButton
            // 
            this.TakeProfitButton.Label = "Take Profit";
            this.TakeProfitButton.Name = "TakeProfitButton";
            this.TakeProfitButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.TakeProfitButton_Click);
            // 
            // SellButton
            // 
            this.CutLossButton.Label = "Cut Loss";
            this.CutLossButton.Name = "CutLossButton";
            this.CutLossButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.CutLossButton_Click);
            
            
            // 
            // GeollyRibbon
            // 
            this.Name = "GeollyRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.GeollyTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.GeollyRibbon_Load);
            this.GeollyTab.ResumeLayout(false);
            this.GeollyTab.PerformLayout();
            this.GeollyGroup.ResumeLayout(false);
            this.GeollyGroup.PerformLayout();
            this.buttonGroup1.ResumeLayout(false);
            this.buttonGroup1.PerformLayout();
            this.buttonGroup2.ResumeLayout(false);
            this.buttonGroup2.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab GeollyTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup GeollyGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton ResetTTAPIButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton StopTTAPIButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup buttonGroup1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BuyButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton SellButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButtonGroup buttonGroup2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton TakeProfitButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton CutLossButton;
    }

    partial class ThisRibbonCollection
    {
        internal GeollyRibbon GeollyRibbon
        {
            get { return this.GetRibbon<GeollyRibbon>(); }
        }
    }
}
