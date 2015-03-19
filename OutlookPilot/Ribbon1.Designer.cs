namespace OutlookPilot
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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.button4 = this.Factory.CreateRibbonButton();
            this.button7 = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.button5 = this.Factory.CreateRibbonButton();
            this.button8 = this.Factory.CreateRibbonButton();
            this.button3 = this.Factory.CreateRibbonButton();
            this.button6 = this.Factory.CreateRibbonButton();
            this.button9 = this.Factory.CreateRibbonButton();
            this.box1 = this.Factory.CreateRibbonBox();
            this.Today = this.Factory.CreateRibbonButton();
            this.EOW = this.Factory.CreateRibbonButton();
            this.Whenever = this.Factory.CreateRibbonButton();
            this.box2 = this.Factory.CreateRibbonBox();
            this.Tomorrow = this.Factory.CreateRibbonButton();
            this.PickDate = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.box1.SuspendLayout();
            this.box2.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabMail";
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabMail";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.button1);
            this.group1.Items.Add(this.button4);
            this.group1.Items.Add(this.button7);
            this.group1.Items.Add(this.button2);
            this.group1.Items.Add(this.button5);
            this.group1.Items.Add(this.button8);
            this.group1.Items.Add(this.button3);
            this.group1.Items.Add(this.button6);
            this.group1.Items.Add(this.button9);
            this.group1.Items.Add(this.box1);
            this.group1.Items.Add(this.box2);
            this.group1.Label = "Defer";
            this.group1.Name = "group1";
            // 
            // button1
            // 
            this.button1.KeyTip = "CMA";
            this.button1.Label = "1";
            this.button1.Name = "button1";
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // button4
            // 
            this.button4.Label = "4";
            this.button4.Name = "button4";
            this.button4.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button4_Click);
            // 
            // button7
            // 
            this.button7.Label = "7";
            this.button7.Name = "button7";
            this.button7.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button7_Click);
            // 
            // button2
            // 
            this.button2.Label = "2";
            this.button2.Name = "button2";
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);
            // 
            // button5
            // 
            this.button5.Label = "5";
            this.button5.Name = "button5";
            this.button5.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button5_Click);
            // 
            // button8
            // 
            this.button8.Label = "8";
            this.button8.Name = "button8";
            this.button8.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button8_Click);
            // 
            // button3
            // 
            this.button3.Label = "3";
            this.button3.Name = "button3";
            this.button3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button3_Click);
            // 
            // button6
            // 
            this.button6.Label = "6";
            this.button6.Name = "button6";
            this.button6.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button6_Click);
            // 
            // button9
            // 
            this.button9.KeyTip = "9";
            this.button9.Label = "9";
            this.button9.Name = "button9";
            this.button9.ScreenTip = "Defer 9 Days";
            this.button9.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button9_Click);
            // 
            // box1
            // 
            this.box1.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            this.box1.Items.Add(this.Today);
            this.box1.Items.Add(this.Tomorrow);
            this.box1.Items.Add(this.EOW);
            this.box1.Name = "box1";
            // 
            // Today
            // 
            this.Today.Label = "Today";
            this.Today.Name = "Today";
            this.Today.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.today_Click);
            // 
            // EOW
            // 
            this.EOW.Label = "This Week";
            this.EOW.Name = "EOW";
            this.EOW.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.eow_Click);
            // 
            // Whenever
            // 
            this.Whenever.Label = "Whenever";
            this.Whenever.Name = "Whenever";
            this.Whenever.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Whenever_Click);
            // 
            // box2
            // 
            this.box2.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical;
            this.box2.Items.Add(this.PickDate);
            this.box2.Items.Add(this.Whenever);
            this.box2.Name = "box2";
            // 
            // Tomorrow
            // 
            this.Tomorrow.Label = "Tomorrow";
            this.Tomorrow.Name = "Tomorrow";
            this.Tomorrow.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.tomorrow_Click);
            // 
            // PickDate
            // 
            this.PickDate.Label = "Pick Date";
            this.PickDate.Name = "PickDate";
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Outlook.Explorer";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.box1.ResumeLayout(false);
            this.box1.PerformLayout();
            this.box2.ResumeLayout(false);
            this.box2.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button6;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button7;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button8;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button9;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Today;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Tomorrow;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton EOW;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Whenever;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box1;
        internal Microsoft.Office.Tools.Ribbon.RibbonBox box2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton PickDate;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
