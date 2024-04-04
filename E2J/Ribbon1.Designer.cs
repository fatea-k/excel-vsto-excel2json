namespace Excel2Json
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
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
            this.Excel2Json = this.Factory.CreateRibbonGroup();
            this.box1 = this.Factory.CreateRibbonBox();
            this.btnExportClient = this.Factory.CreateRibbonButton();
            this.btnExportServer = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.button2 = this.Factory.CreateRibbonButton();
            this.button1 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.Excel2Json.SuspendLayout();
            this.box1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.Excel2Json);
            this.tab1.Label = "Excel转Json";
            this.tab1.Name = "tab1";
            // 
            // Excel2Json
            // 
            this.Excel2Json.Items.Add(this.box1);
            this.Excel2Json.Items.Add(this.separator1);
            this.Excel2Json.Items.Add(this.button2);
            this.Excel2Json.Items.Add(this.button1);
            this.Excel2Json.Label = "Excel转Json";
            this.Excel2Json.Name = "Excel2Json";
            // 
            // box1
            // 
            this.box1.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            this.box1.Items.Add(this.btnExportClient);
            this.box1.Items.Add(this.btnExportServer);
            this.box1.Name = "box1";
            // 
            // btnExportClient
            // 
            this.btnExportClient.Label = "导出客户端JSON(自用非标准格式)";
            this.btnExportClient.Name = "btnExportClient";
            this.btnExportClient.ScreenTip = "导出客户端JSON";
            this.btnExportClient.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExportClient_Click);
            // 
            // btnExportServer
            // 
            this.btnExportServer.Label = "导出服务端Excel";
            this.btnExportServer.Name = "btnExportServer";
            this.btnExportServer.ScreenTip = "导出服务端Excel";
            this.btnExportServer.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExportServer_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // button2
            // 
            this.button2.Label = "导出客户端JSON";
            this.button2.Name = "button2";
            this.button2.ScreenTip = "导出客户端JSON";
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);
            // 
            // button1
            // 
            this.button1.Label = "导出服务端Json";
            this.button1.Name = "button1";
            this.button1.ScreenTip = "导出服务端Json";
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExportServerJson_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.Excel2Json.ResumeLayout(false);
            this.Excel2Json.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Excel2Json;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExportClient;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExportServer;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
