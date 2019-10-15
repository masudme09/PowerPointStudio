namespace PowerPointStudio
{
    partial class PowerPointStudioRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public PowerPointStudioRibbon()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(PowerPointStudioRibbon));
            this.tab2 = this.Factory.CreateRibbonTab();
            this.Extract = this.Factory.CreateRibbonGroup();
            this.preview = this.Factory.CreateRibbonGroup();
            this.btnExtractSlides = this.Factory.CreateRibbonButton();
            this.btnPreviewWeb = this.Factory.CreateRibbonButton();
            this.btnPreviewJSON = this.Factory.CreateRibbonButton();
            this.tab2.SuspendLayout();
            this.Extract.SuspendLayout();
            this.preview.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab2
            // 
            this.tab2.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab2.Groups.Add(this.Extract);
            this.tab2.Groups.Add(this.preview);
            this.tab2.Label = "PowerPointStudio";
            this.tab2.Name = "tab2";
            // 
            // Extract
            // 
            this.Extract.Items.Add(this.btnExtractSlides);
            this.Extract.Label = "Create             ";
            this.Extract.Name = "Extract";
            // 
            // preview
            // 
            this.preview.Items.Add(this.btnPreviewWeb);
            this.preview.Items.Add(this.btnPreviewJSON);
            this.preview.Label = "Preview";
            this.preview.Name = "preview";
            // 
            // btnExtractSlides
            // 
            this.btnExtractSlides.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnExtractSlides.Image = ((System.Drawing.Image)(resources.GetObject("btnExtractSlides.Image")));
            this.btnExtractSlides.Label = "ExtarctSlides ";
            this.btnExtractSlides.Name = "btnExtractSlides";
            this.btnExtractSlides.ShowImage = true;
            this.btnExtractSlides.SuperTip = "Extract Slides to Generate JSON & HTML";
            this.btnExtractSlides.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnExtractSlides_Click);
            // 
            // btnPreviewWeb
            // 
            this.btnPreviewWeb.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnPreviewWeb.Image = ((System.Drawing.Image)(resources.GetObject("btnPreviewWeb.Image")));
            this.btnPreviewWeb.Label = "Preview Web";
            this.btnPreviewWeb.Name = "btnPreviewWeb";
            this.btnPreviewWeb.ShowImage = true;
            this.btnPreviewWeb.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnPreviewWeb_Click);
            // 
            // btnPreviewJSON
            // 
            this.btnPreviewJSON.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnPreviewJSON.Image = ((System.Drawing.Image)(resources.GetObject("btnPreviewJSON.Image")));
            this.btnPreviewJSON.Label = "Preview JSON";
            this.btnPreviewJSON.Name = "btnPreviewJSON";
            this.btnPreviewJSON.ShowImage = true;
            this.btnPreviewJSON.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnPreviewJSON_Click);
            // 
            // PowerPointStudioRibbon
            // 
            this.Name = "PowerPointStudioRibbon";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab2);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.PowerPointStudioRibbon_Load);
            this.tab2.ResumeLayout(false);
            this.tab2.PerformLayout();
            this.Extract.ResumeLayout(false);
            this.Extract.PerformLayout();
            this.preview.ResumeLayout(false);
            this.preview.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Extract;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExtractSlides;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup preview;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPreviewWeb;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPreviewJSON;
    }

    partial class ThisRibbonCollection
    {
        internal PowerPointStudioRibbon PowerPointStudioRibbon
        {
            get { return this.GetRibbon<PowerPointStudioRibbon>(); }
        }
    }
}
