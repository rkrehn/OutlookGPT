using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;

namespace OutlookGPT
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
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl3 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl4 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl5 = this.Factory.CreateRibbonDropDownItem();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.dropDown1 = this.Factory.CreateRibbonDropDown();
            this.button1 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.dropDown1);
            this.group1.Items.Add(this.button1);
            this.group1.Label = "OutlookGPT";
            this.group1.Name = "group1";
            // 
            // dropDown1
            // 
            ribbonDropDownItemImpl1.Label = "Positive";
            ribbonDropDownItemImpl1.ScreenTip = "Alter the language of your email to appear positive.";
            ribbonDropDownItemImpl2.Label = "Conscionable";
            ribbonDropDownItemImpl2.ScreenTip = "Rewrite your email to be acceptable and proper.";
            ribbonDropDownItemImpl3.Label = "Politically Correct";
            ribbonDropDownItemImpl3.ScreenTip = "Rewrite your email to be politically correct for those difficult scenarios.";
            ribbonDropDownItemImpl4.Label = "Stern - Gently";
            ribbonDropDownItemImpl4.ScreenTip = "Rewrite your email to be stern, but in a gentle manner";
            ribbonDropDownItemImpl5.Label = "Stern - Direct";
            ribbonDropDownItemImpl5.ScreenTip = "Rewrite your email to be stern and very direct.";
            this.dropDown1.Items.Add(ribbonDropDownItemImpl1);
            this.dropDown1.Items.Add(ribbonDropDownItemImpl2);
            this.dropDown1.Items.Add(ribbonDropDownItemImpl3);
            this.dropDown1.Items.Add(ribbonDropDownItemImpl4);
            this.dropDown1.Items.Add(ribbonDropDownItemImpl5);
            this.dropDown1.Label = "Mood";
            this.dropDown1.Name = "dropDown1";
            this.dropDown1.ButtonClick += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.dropDown1_ButtonClick);
            this.dropDown1.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.dropDown1_SelectionChanged);
            // 
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Image = global::OutlookGPT.Properties.Resources.logo;
            this.button1.Label = "Correct";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Outlook.Mail.Compose";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal RibbonDropDown dropDown1;
        internal RibbonButton button1;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
