namespace ExcelToPaper
{
    partial class RibbonToolBar : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonToolBar()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RibbonToolBar));
            this.ExcelToPaper = this.Factory.CreateRibbonTab();
            this.ExcelToPaperGroup = this.Factory.CreateRibbonGroup();
            this.BatchPrint = this.Factory.CreateRibbonButton();
            this.ExcelToPaper.SuspendLayout();
            this.ExcelToPaperGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // ExcelToPaper
            // 
            this.ExcelToPaper.Groups.Add(this.ExcelToPaperGroup);
            this.ExcelToPaper.Label = "ExcelToPaper";
            this.ExcelToPaper.Name = "ExcelToPaper";
            // 
            // ExcelToPaperGroup
            // 
            this.ExcelToPaperGroup.Items.Add(this.BatchPrint);
            this.ExcelToPaperGroup.Label = "ExcelToPaper";
            this.ExcelToPaperGroup.Name = "ExcelToPaperGroup";
            // 
            // BatchPrint
            // 
            this.BatchPrint.Image = ((System.Drawing.Image)(resources.GetObject("BatchPrint.Image")));
            this.BatchPrint.Label = "BatchPrint";
            this.BatchPrint.Name = "BatchPrint";
            this.BatchPrint.ShowImage = true;
            this.BatchPrint.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BatchToPaper_Click);
            // 
            // RibbonToolBar
            // 
            this.Name = "RibbonToolBar";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.ExcelToPaper);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonToolBar_Load);
            this.ExcelToPaper.ResumeLayout(false);
            this.ExcelToPaper.PerformLayout();
            this.ExcelToPaperGroup.ResumeLayout(false);
            this.ExcelToPaperGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab ExcelToPaper;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup ExcelToPaperGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BatchPrint;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonToolBar RibbonToolBar
        {
            get { return this.GetRibbon<RibbonToolBar>(); }
        }
    }
}
