namespace RefusalReasonList
{
    partial class rrRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public rrRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージド リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region コンポーネント デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            this.JpoApi0 = this.Factory.CreateRibbonTab();
            this.rrGroup = this.Factory.CreateRibbonGroup();
            this.rrButton = this.Factory.CreateRibbonButton();
            this.buttonRR2Word = this.Factory.CreateRibbonButton();
            this.buttonAccount = this.Factory.CreateRibbonButton();
            this.JpoApi0.SuspendLayout();
            this.rrGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // JpoApi0
            // 
            this.JpoApi0.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.JpoApi0.Groups.Add(this.rrGroup);
            this.JpoApi0.Label = "JpoApi";
            this.JpoApi0.Name = "JpoApi0";
            // 
            // rrGroup
            // 
            this.rrGroup.Items.Add(this.rrButton);
            this.rrGroup.Items.Add(this.buttonRR2Word);
            this.rrGroup.Items.Add(this.buttonAccount);
            this.rrGroup.Label = "拒絶理由通知";
            this.rrGroup.Name = "rrGroup";
            // 
            // rrButton
            // 
            this.rrButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.rrButton.Label = "条文";
            this.rrButton.Name = "rrButton";
            this.rrButton.OfficeImageId = "GroupListActions";
            this.rrButton.ShowImage = true;
            this.rrButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // buttonRR2Word
            // 
            this.buttonRR2Word.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonRR2Word.Label = "拒理Word";
            this.buttonRR2Word.Name = "buttonRR2Word";
            this.buttonRR2Word.OfficeImageId = "FileSaveAsWordDocx";
            this.buttonRR2Word.ShowImage = true;
            this.buttonRR2Word.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonRR2Word_Click);
            // 
            // buttonAccount
            // 
            this.buttonAccount.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonAccount.Label = "アカウント";
            this.buttonAccount.Name = "buttonAccount";
            this.buttonAccount.OfficeImageId = "AccountSettings";
            this.buttonAccount.ShowImage = true;
            this.buttonAccount.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonAccount_Click);
            // 
            // rrRibbon
            // 
            this.Name = "rrRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.JpoApi0);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.rrRibbon_Load);
            this.JpoApi0.ResumeLayout(false);
            this.JpoApi0.PerformLayout();
            this.rrGroup.ResumeLayout(false);
            this.rrGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab JpoApi0;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup rrGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton rrButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonRR2Word;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonAccount;
    }

    partial class ThisRibbonCollection
    {
        internal rrRibbon rrRibbon
        {
            get { return this.GetRibbon<rrRibbon>(); }
        }
    }
}
