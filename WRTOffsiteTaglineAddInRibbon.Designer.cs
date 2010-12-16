namespace WRTOffsite_NET35
{
    partial class WRTOffsiteTaglineAddInRibbon : Microsoft.Office.Tools.Ribbon.OfficeRibbon
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public WRTOffsiteTaglineAddInRibbon()
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
            this.tab1 = new Microsoft.Office.Tools.Ribbon.RibbonTab();
            this.WRTOffsiteTagline = new Microsoft.Office.Tools.Ribbon.RibbonGroup();
            this.ActiveAllMessages = new Microsoft.Office.Tools.Ribbon.RibbonToggleButton();
            this.ActiveThisMessage = new Microsoft.Office.Tools.Ribbon.RibbonToggleButton();
            this.tab1.SuspendLayout();
            this.WRTOffsiteTagline.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.WRTOffsiteTagline);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // WRTOffsiteTagline
            // 
            this.WRTOffsiteTagline.Items.Add(this.ActiveAllMessages);
            this.WRTOffsiteTagline.Items.Add(this.ActiveThisMessage);
            this.WRTOffsiteTagline.Label = "WRT Offsite Tagline";
            this.WRTOffsiteTagline.Name = "WRTOffsiteTagline";
            // 
            // ActiveAllMessages
            // 
            this.ActiveAllMessages.Label = "Inactive - All Messages";
            this.ActiveAllMessages.Name = "ActiveAllMessages";
            this.ActiveAllMessages.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.ActiveAllMessages_Click);
            // 
            // ActiveThisMessage
            // 
            this.ActiveThisMessage.Checked = true;
            this.ActiveThisMessage.Label = "Active - This Message";
            this.ActiveThisMessage.Name = "ActiveThisMessage";
            this.ActiveThisMessage.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.ActiveThisMessage_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Outlook.Mail.Compose";
            this.Tabs.Add(this.tab1);
            this.Load += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonUIEventArgs>(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.WRTOffsiteTagline.ResumeLayout(false);
            this.WRTOffsiteTagline.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup WRTOffsiteTagline;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton ActiveAllMessages;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton ActiveThisMessage;
    }

    partial class ThisRibbonCollection : Microsoft.Office.Tools.Ribbon.RibbonReadOnlyCollection
    {
        internal WRTOffsiteTaglineAddInRibbon Ribbon1
        {
            get { return this.GetRibbon<WRTOffsiteTaglineAddInRibbon>(); }
        }
    }
}
