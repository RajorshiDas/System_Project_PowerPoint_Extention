namespace PowerPointAddIn1
{
    partial class MyRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public MyRibbon()
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
            this.grpMyTools = this.Factory.CreateRibbonGroup();
            this.btnAddSlide = this.Factory.CreateRibbonButton();
            this.button1 = this.Factory.CreateRibbonButton();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.lblTotal = this.Factory.CreateRibbonLabel();
            this.lblSection = this.Factory.CreateRibbonLabel();
            this.lblSlide = this.Factory.CreateRibbonLabel();
            this.lblTotalValue = this.Factory.CreateRibbonLabel();
            this.lblSectionValue = this.Factory.CreateRibbonLabel();
            this.lblSlideValue = this.Factory.CreateRibbonLabel();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.SubSectionStart = this.Factory.CreateRibbonEditBox();
            this.lblSubsectionName = this.Factory.CreateRibbonLabel();
            this.btnCreateSubsection = this.Factory.CreateRibbonButton();
            this.SubSectionEnd = this.Factory.CreateRibbonEditBox();
            this.valueSubsectionName = this.Factory.CreateRibbonLabel();
            this.NavbarGruop = this.Factory.CreateRibbonGroup();
            this.btnCreateNav = this.Factory.CreateRibbonButton();
            this.btnRefreshNav = this.Factory.CreateRibbonButton();
            this.btnRemoveNav = this.Factory.CreateRibbonButton();
            this.btnNavBarSetting = this.Factory.CreateRibbonButton();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.btnZoomToArea = this.Factory.CreateRibbonButton();
            this.btnZoomSettings = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.grpMyTools.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.NavbarGruop.SuspendLayout();
            this.group4.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.grpMyTools);
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.NavbarGruop);
            this.tab1.Groups.Add(this.group4);
            this.tab1.Label = "My Tools";
            this.tab1.Name = "tab1";
            // 
            // grpMyTools
            // 
            this.grpMyTools.Items.Add(this.btnAddSlide);
            this.grpMyTools.Items.Add(this.button1);
            this.grpMyTools.Label = "Slide Operations";
            this.grpMyTools.Name = "grpMyTools";
            // 
            // btnAddSlide
            // 
            this.btnAddSlide.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnAddSlide.Label = "Add Slide";
            this.btnAddSlide.Name = "btnAddSlide";
            this.btnAddSlide.ShowImage = true;
            this.btnAddSlide.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAddSlide_Click);
            // 
            // button1
            // 
            this.button1.Label = "Insert Hello World";
            this.button1.Name = "button1";
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // group1
            // 
            this.group1.Items.Add(this.lblTotal);
            this.group1.Items.Add(this.lblSection);
            this.group1.Items.Add(this.lblSlide);
            this.group1.Items.Add(this.lblTotalValue);
            this.group1.Items.Add(this.lblSectionValue);
            this.group1.Items.Add(this.lblSlideValue);
            this.group1.Label = "Information";
            this.group1.Name = "group1";
            // 
            // lblTotal
            // 
            this.lblTotal.Label = "Total Sections:";
            this.lblTotal.Name = "lblTotal";
            // 
            // lblSection
            // 
            this.lblSection.Label = "Current Section:";
            this.lblSection.Name = "lblSection";
            // 
            // lblSlide
            // 
            this.lblSlide.Label = "Slides in Section:";
            this.lblSlide.Name = "lblSlide";
            // 
            // lblTotalValue
            // 
            this.lblTotalValue.Label = "Value";
            this.lblTotalValue.Name = "lblTotalValue";
            // 
            // lblSectionValue
            // 
            this.lblSectionValue.Label = "Value";
            this.lblSectionValue.Name = "lblSectionValue";
            // 
            // lblSlideValue
            // 
            this.lblSlideValue.Label = "Value";
            this.lblSlideValue.Name = "lblSlideValue";
            // 
            // group2
            // 
            this.group2.Items.Add(this.SubSectionStart);
            this.group2.Items.Add(this.lblSubsectionName);
            this.group2.Items.Add(this.btnCreateSubsection);
            this.group2.Items.Add(this.SubSectionEnd);
            this.group2.Items.Add(this.valueSubsectionName);
            this.group2.Label = "Subsection";
            this.group2.Name = "group2";
            // 
            // SubSectionStart
            // 
            this.SubSectionStart.Label = "Start:";
            this.SubSectionStart.Name = "SubSectionStart";
            this.SubSectionStart.Text = null;
            this.SubSectionStart.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.editBox2_TextChanged);
            // 
            // lblSubsectionName
            // 
            this.lblSubsectionName.Label = "Current subsection name:";
            this.lblSubsectionName.Name = "lblSubsectionName";
            // 
            // btnCreateSubsection
            // 
            this.btnCreateSubsection.Label = "Create Subsection";
            this.btnCreateSubsection.Name = "btnCreateSubsection";
            this.btnCreateSubsection.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCreateSubsection_Click);
            // 
            // SubSectionEnd
            // 
            this.SubSectionEnd.Label = "End:";
            this.SubSectionEnd.Name = "SubSectionEnd";
            this.SubSectionEnd.Text = null;
            // 
            // valueSubsectionName
            // 
            this.valueSubsectionName.Label = "Value";
            this.valueSubsectionName.Name = "valueSubsectionName";
            // 
            // NavbarGruop
            // 
            this.NavbarGruop.Items.Add(this.btnCreateNav);
            this.NavbarGruop.Items.Add(this.btnRefreshNav);
            this.NavbarGruop.Items.Add(this.btnRemoveNav);
            this.NavbarGruop.Items.Add(this.btnNavBarSetting);
            this.NavbarGruop.Label = "Navigation Bar";
            this.NavbarGruop.Name = "NavbarGruop";
            // 
            // btnCreateNav
            // 
            this.btnCreateNav.Label = "Create Nav Bar";
            this.btnCreateNav.Name = "btnCreateNav";
            this.btnCreateNav.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click_1);
            // 
            // btnRefreshNav
            // 
            this.btnRefreshNav.Label = "Refresh Nav Bar";
            this.btnRefreshNav.Name = "btnRefreshNav";
            this.btnRefreshNav.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button3_Click);
            // 
            // btnRemoveNav
            // 
            this.btnRemoveNav.Label = "Remove Nav Bar";
            this.btnRemoveNav.Name = "btnRemoveNav";
            this.btnRemoveNav.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button4_Click);
            // 
            // btnNavBarSetting
            // 
            this.btnNavBarSetting.Label = "Customize Nav Bar";
            this.btnNavBarSetting.Name = "btnNavBarSetting";
            this.btnNavBarSetting.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnNavBarSettings_Click);
            // 
            // group4
            // 
            this.group4.Items.Add(this.btnZoomToArea);
            this.group4.Items.Add(this.btnZoomSettings);
            this.group4.Label = "Zoom Lab";
            this.group4.Name = "group4";
            // 
            // btnZoomToArea
            // 
            this.btnZoomToArea.Label = "Zoom to Area";
            this.btnZoomToArea.Name = "btnZoomToArea";
            this.btnZoomToArea.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click_2);
            // 
            // btnZoomSettings
            // 
            this.btnZoomSettings.Label = "Zoom Settings";
            this.btnZoomSettings.Name = "btnZoomSettings";
            this.btnZoomSettings.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnZoomSettings_Click_1);
            // 
            // MyRibbon
            // 
            this.Name = "MyRibbon";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MyRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.grpMyTools.ResumeLayout(false);
            this.grpMyTools.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.NavbarGruop.ResumeLayout(false);
            this.NavbarGruop.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grpMyTools;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddSlide;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lblTotal;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lblSection;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lblSlide;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lblTotalValue;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lblSectionValue;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lblSlideValue;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lblSubsectionName;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCreateSubsection;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel valueSubsectionName;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup NavbarGruop;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCreateNav;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRefreshNav;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRemoveNav;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox SubSectionStart;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox SubSectionEnd;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnNavBarSetting;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnZoomToArea;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnZoomSettings;
    }
}
