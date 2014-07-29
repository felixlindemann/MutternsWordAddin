namespace MutternsWordAddin
{
    partial class RibbonMuttern : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Erforderliche Designervariable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonMuttern()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Verwendete Ressourcen bereinigen.
        /// </summary>
        /// <param name="disposing">"true", wenn verwaltete Ressourcen gelöscht werden sollen; andernfalls "false".</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Vom Komponenten-Designer generierter Code

        /// <summary>
        /// Erforderliche Methode für Designerunterstützung -
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RibbonMuttern));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.tab2 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnStep1 = this.Factory.CreateRibbonButton();
            this.buttonFillContent = this.Factory.CreateRibbonButton();
            this.splitButton1 = this.Factory.CreateRibbonSplitButton();
            this.buttonAveryZweckform = this.Factory.CreateRibbonButton();
            this.button1 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.tab2.SuspendLayout();
            this.group1.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // tab2
            // 
            this.tab2.Groups.Add(this.group1);
            this.tab2.Label = "Mutterns Tools";
            this.tab2.Name = "tab2";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnStep1);
            this.group1.Items.Add(this.buttonFillContent);
            this.group1.Items.Add(this.splitButton1);
            this.group1.Label = "Etiketten erstellen";
            this.group1.Name = "group1";
            // 
            // btnStep1
            // 
            this.btnStep1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnStep1.Image = ((System.Drawing.Image)(resources.GetObject("btnStep1.Image")));
            this.btnStep1.Label = "Schritt 1: Vorlage laden";
            this.btnStep1.Name = "btnStep1";
            this.btnStep1.ShowImage = true;
            this.btnStep1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnStep1_Click);
            // 
            // buttonFillContent
            // 
            this.buttonFillContent.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.buttonFillContent.Image = ((System.Drawing.Image)(resources.GetObject("buttonFillContent.Image")));
            this.buttonFillContent.Label = "Schritt 2: Inhalt füllen";
            this.buttonFillContent.Name = "buttonFillContent";
            this.buttonFillContent.ShowImage = true;
            this.buttonFillContent.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonFillContent_Click);
            // 
            // splitButton1
            // 
            this.splitButton1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.splitButton1.Image = ((System.Drawing.Image)(resources.GetObject("splitButton1.Image")));
            this.splitButton1.Items.Add(this.buttonAveryZweckform);
            this.splitButton1.Items.Add(this.button1);
            this.splitButton1.Label = "Einstellungen";
            this.splitButton1.Name = "splitButton1";
            // 
            // buttonAveryZweckform
            // 
            this.buttonAveryZweckform.Description = "Installiere Avery Zweckform Assistent";
            this.buttonAveryZweckform.Image = ((System.Drawing.Image)(resources.GetObject("buttonAveryZweckform.Image")));
            this.buttonAveryZweckform.Label = "Installiere Avery Zweckform Assistent";
            this.buttonAveryZweckform.Name = "buttonAveryZweckform";
            this.buttonAveryZweckform.ScreenTip = "Installiere Avery Zweckform Assistent";
            this.buttonAveryZweckform.ShowImage = true;
            this.buttonAveryZweckform.SuperTip = "Installiere Avery Zweckform Assistent";
            this.buttonAveryZweckform.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonAveryZweckform_Click);
            // 
            // button1
            // 
            this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
            this.button1.Label = "Über dieses Addin";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click_1);
            // 
            // RibbonMuttern
            // 
            this.Name = "RibbonMuttern";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Tabs.Add(this.tab2);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RibbonMuttern_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.tab2.ResumeLayout(false);
            this.tab2.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        private Microsoft.Office.Tools.Ribbon.RibbonTab tab2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton splitButton1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonAveryZweckform;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnStep1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonFillContent;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonMuttern RibbonMuttern
        {
            get { return this.GetRibbon<RibbonMuttern>(); }
        }
    }
}
