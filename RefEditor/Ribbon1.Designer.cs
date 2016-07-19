namespace RefEditor
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon1));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.button3 = this.Factory.CreateRibbonButton();
            this.button4 = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.button5 = this.Factory.CreateRibbonButton();
            this.button6 = this.Factory.CreateRibbonButton();
            this.button7 = this.Factory.CreateRibbonButton();
            this.button8 = this.Factory.CreateRibbonButton();
            this.button9 = this.Factory.CreateRibbonButton();
            this.button10 = this.Factory.CreateRibbonButton();
            this.button12 = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.button11 = this.Factory.CreateRibbonButton();
            this.button14 = this.Factory.CreateRibbonButton();
            this.button13 = this.Factory.CreateRibbonButton();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.group5 = this.Factory.CreateRibbonGroup();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group3.SuspendLayout();
            this.group2.SuspendLayout();
            this.group4.SuspendLayout();
            this.group5.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group5);
            this.tab1.Groups.Add(this.group4);
            this.tab1.Label = "Reference Editor";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.button3);
            this.group1.Items.Add(this.button4);
            this.group1.Label = "References List";
            this.group1.Name = "group1";
            // 
            // button3
            // 
            this.button3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button3.Image = ((System.Drawing.Image)(resources.GetObject("button3.Image")));
            this.button3.Label = "Open Reference Table";
            this.button3.Name = "button3";
            this.button3.ShowImage = true;
            this.button3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button3_Click);
            // 
            // button4
            // 
            this.button4.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button4.Image = ((System.Drawing.Image)(resources.GetObject("button4.Image")));
            this.button4.Label = "Initialize Reference Table";
            this.button4.Name = "button4";
            this.button4.ShowImage = true;
            this.button4.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button4_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.button5);
            this.group3.Items.Add(this.button6);
            this.group3.Items.Add(this.button7);
            this.group3.Label = "References in Active Document";
            this.group3.Name = "group3";
            // 
            // button5
            // 
            this.button5.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button5.Image = ((System.Drawing.Image)(resources.GetObject("button5.Image")));
            this.button5.Label = "Insert New Reference";
            this.button5.Name = "button5";
            this.button5.ShowImage = true;
            this.button5.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button5_Click);
            // 
            // button6
            // 
            this.button6.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button6.Image = ((System.Drawing.Image)(resources.GetObject("button6.Image")));
            this.button6.Label = "Operate References";
            this.button6.Name = "button6";
            this.button6.ShowImage = true;
            this.button6.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button6_Click);
            // 
            // button7
            // 
            this.button7.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button7.Image = ((System.Drawing.Image)(resources.GetObject("button7.Image")));
            this.button7.Label = "Operate All Docs";
            this.button7.Name = "button7";
            this.button7.ShowImage = true;
            this.button7.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button7_Click_1);
            // 
            // button8
            // 
            this.button8.Label = "Collect all refs";
            this.button8.Name = "button8";
            this.button8.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button8_Click);
            // 
            // button9
            // 
            this.button9.Label = "FindRefsBlocks";
            this.button9.Name = "button9";
            this.button9.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button9_Click);
            // 
            // button10
            // 
            this.button10.Label = "Bookmarks 2 Refs";
            this.button10.Name = "button10";
            this.button10.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button10_Click);
            // 
            // button12
            // 
            this.button12.Label = "Bookmarks 2 table";
            this.button12.Name = "button12";
            this.button12.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button12_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.button1);
            this.group2.Items.Add(this.button2);
            this.group2.Label = "References in Clowd Database";
            this.group2.Name = "group2";
            // 
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
            this.button1.Label = "Connect Clowd Database";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button2.Image = ((System.Drawing.Image)(resources.GetObject("button2.Image")));
            this.button2.Label = "Insert New Reference to Clowd Database";
            this.button2.Name = "button2";
            this.button2.ShowImage = true;
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);
            // 
            // button11
            // 
            this.button11.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button11.Image = ((System.Drawing.Image)(resources.GetObject("button11.Image")));
            this.button11.Label = "Connect Local Database";
            this.button11.Name = "button11";
            this.button11.ShowImage = true;
            this.button11.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button11_Click);
            // 
            // button14
            // 
            this.button14.Label = "Clear double #-s";
            this.button14.Name = "button14";
            this.button14.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button14_Click);
            // 
            // button13
            // 
            this.button13.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button13.Image = ((System.Drawing.Image)(resources.GetObject("button13.Image")));
            this.button13.Label = "Update Footnotes";
            this.button13.Name = "button13";
            this.button13.ShowImage = true;
            this.button13.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button13_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            this.openFileDialog1.FileOk += new System.ComponentModel.CancelEventHandler(this.openFileDialog1_FileOk);
            // 
            // group4
            // 
            this.group4.Items.Add(this.button8);
            this.group4.Items.Add(this.button9);
            this.group4.Items.Add(this.button10);
            this.group4.Items.Add(this.button12);
            this.group4.Items.Add(this.button14);
            this.group4.Items.Add(this.button13);
            this.group4.Label = "Supporting Operations";
            this.group4.Name = "group4";
            // 
            // group5
            // 
            this.group5.Items.Add(this.button11);
            this.group5.Label = "References in Local Database";
            this.group5.Name = "group5";
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.group5.ResumeLayout(false);
            this.group5.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button4;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button6;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button7;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button8;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button9;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button10;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button11;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button12;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button13;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button14;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group5;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
