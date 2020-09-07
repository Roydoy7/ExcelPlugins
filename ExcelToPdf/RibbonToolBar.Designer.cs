namespace ExcelToPdf
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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.ExcelToPdfGroup = this.Factory.CreateRibbonGroup();
            this.BatchToPdf = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.ExcelToPdfGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.ExcelToPdfGroup);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // ExcelToPdfGroup
            // 
            this.ExcelToPdfGroup.Items.Add(this.BatchToPdf);
            this.ExcelToPdfGroup.Label = "ExcelToPdf";
            this.ExcelToPdfGroup.Name = "ExcelToPdfGroup";
            // 
            // BatchToPdf
            // 
            this.BatchToPdf.Image = global::ExcelToPdf.Properties.Resources.ExcelToPdf;
            this.BatchToPdf.Label = "BatchToPdf";
            this.BatchToPdf.Name = "BatchToPdf";
            this.BatchToPdf.ShowImage = true;
            this.BatchToPdf.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BatchToPdf_Click);
            // 
            // RibbonToolBar
            // 
            this.Name = "RibbonToolBar";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonToolBar_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.ExcelToPdfGroup.ResumeLayout(false);
            this.ExcelToPdfGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup ExcelToPdfGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton BatchToPdf;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonToolBar RibbonToolBar
        {
            get { return this.GetRibbon<RibbonToolBar>(); }
        }
    }
}
