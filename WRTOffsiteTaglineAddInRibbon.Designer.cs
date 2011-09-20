namespace WRTOffsite_NET35
{
    partial class WRTOffsiteTaglineAddInRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public WRTOffsiteTaglineAddInRibbon()
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
            this.tab_WRTOffsite = this.Factory.CreateRibbonTab();
            this.grp_WRTOffsiteTagline = this.Factory.CreateRibbonGroup();
            this.ActiveAllMessages = this.Factory.CreateRibbonToggleButton();
            this.ActiveThisMessage = this.Factory.CreateRibbonToggleButton();
            this.tab_WRTOffsite.SuspendLayout();
            this.grp_WRTOffsiteTagline.SuspendLayout();
            // 
            // tab_WRTOffsite
            // 
            this.tab_WRTOffsite.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab_WRTOffsite.ControlId.OfficeId = "TabNewMailMessage";
            this.tab_WRTOffsite.Groups.Add(this.grp_WRTOffsiteTagline);
            this.tab_WRTOffsite.Label = "TabNewMailMessage";
            this.tab_WRTOffsite.Name = "tab_WRTOffsite";
            // 
            // grp_WRTOffsiteTagline
            // 
            this.grp_WRTOffsiteTagline.Items.Add(this.ActiveAllMessages);
            this.grp_WRTOffsiteTagline.Items.Add(this.ActiveThisMessage);
            this.grp_WRTOffsiteTagline.Label = "WRT Offsite Tagline";
            this.grp_WRTOffsiteTagline.Name = "grp_WRTOffsiteTagline";
            // 
            // ActiveAllMessages
            // 
            this.ActiveAllMessages.Label = "Active -All Messages";
            this.ActiveAllMessages.Name = "ActiveAllMessages";
            this.ActiveAllMessages.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ActiveAllMessages_Click);
            // 
            // ActiveThisMessage
            // 
            this.ActiveThisMessage.Label = "Active - This Message Only";
            this.ActiveThisMessage.Name = "ActiveThisMessage";
            this.ActiveThisMessage.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ActiveThisMessage_Click);
            // 
            // WRTOffsiteTaglineAddInRibbon
            // 
            this.Name = "WRTOffsiteTaglineAddInRibbon";
            this.RibbonType = "Microsoft.Outlook.Mail.Compose";
            this.Tabs.Add(this.tab_WRTOffsite);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab_WRTOffsite.ResumeLayout(false);
            this.tab_WRTOffsite.PerformLayout();
            this.grp_WRTOffsiteTagline.ResumeLayout(false);
            this.grp_WRTOffsiteTagline.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab_WRTOffsite;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grp_WRTOffsiteTagline;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton ActiveAllMessages;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton ActiveThisMessage;
    }

    partial class ThisRibbonCollection
    {
        internal WRTOffsiteTaglineAddInRibbon Ribbon1
        {
            get { return this.GetRibbon<WRTOffsiteTaglineAddInRibbon>(); }
        }
    }
}
