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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MyRibbon));
            this.tab1 = this.Factory.CreateRibbonTab();
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
            this.linkbtn = this.Factory.CreateRibbonButton();
            this.adjustbtn = this.Factory.CreateRibbonButton();
            this.ZoomGroup = this.Factory.CreateRibbonGroup();
            this.zoomaddbtn = this.Factory.CreateRibbonButton();
            this.selectzoombtn = this.Factory.CreateRibbonButton();
            this.createzoombtn = this.Factory.CreateRibbonButton();
            this.effectGruop = this.Factory.CreateRibbonGroup();
            this.effectsplitbtn = this.Factory.CreateRibbonSplitButton();
            this.focusMenu = this.Factory.CreateRibbonMenu();
            this.focusbtn = this.Factory.CreateRibbonButton();
            this.spotlightMenu = this.Factory.CreateRibbonMenu();
            this.createspotlightbtn = this.Factory.CreateRibbonButton();
            this.spotlightSettingsBtn = this.Factory.CreateRibbonButton();
            this.blurMenu = this.Factory.CreateRibbonMenu();
            this.MagnifyingGlassMenu = this.Factory.CreateRibbonMenu();
            this.magniaddbtn = this.Factory.CreateRibbonButton();
            this.magsetingsbtn = this.Factory.CreateRibbonButton();
            this.selectEffecctdtn = this.Factory.CreateRibbonButton();
            this.hyperlinkgroup = this.Factory.CreateRibbonGroup();
            this.SlideNobox = this.Factory.CreateRibbonEditBox();
            this.crtHypBtn = this.Factory.CreateRibbonButton();
            this.rmvHypbtn = this.Factory.CreateRibbonButton();
            this.selecthypbtn = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.NavbarGruop.SuspendLayout();
            this.ZoomGroup.SuspendLayout();
            this.effectGruop.SuspendLayout();
            this.hyperlinkgroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.NavbarGruop);
            this.tab1.Groups.Add(this.ZoomGroup);
            this.tab1.Groups.Add(this.effectGruop);
            this.tab1.Groups.Add(this.hyperlinkgroup);
            this.tab1.Label = "My Tools";
            this.tab1.Name = "tab1";
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
            this.NavbarGruop.Items.Add(this.linkbtn);
            this.NavbarGruop.Items.Add(this.adjustbtn);
            this.NavbarGruop.Label = "Navigation Bar";
            this.NavbarGruop.Name = "NavbarGruop";
            // 
            // btnCreateNav
            // 
            this.btnCreateNav.Image = ((System.Drawing.Image)(resources.GetObject("btnCreateNav.Image")));
            this.btnCreateNav.Label = "Create Nav Bar";
            this.btnCreateNav.Name = "btnCreateNav";
            this.btnCreateNav.ShowImage = true;
            this.btnCreateNav.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click_1);
            // 
            // btnRefreshNav
            // 
            this.btnRefreshNav.Image = ((System.Drawing.Image)(resources.GetObject("btnRefreshNav.Image")));
            this.btnRefreshNav.Label = "Refresh Nav Bar";
            this.btnRefreshNav.Name = "btnRefreshNav";
            this.btnRefreshNav.ShowImage = true;
            this.btnRefreshNav.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button3_Click);
            // 
            // btnRemoveNav
            // 
            this.btnRemoveNav.Image = ((System.Drawing.Image)(resources.GetObject("btnRemoveNav.Image")));
            this.btnRemoveNav.Label = "Remove Nav Bar";
            this.btnRemoveNav.Name = "btnRemoveNav";
            this.btnRemoveNav.ShowImage = true;
            this.btnRemoveNav.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button4_Click);
            // 
            // btnNavBarSetting
            // 
            this.btnNavBarSetting.Image = ((System.Drawing.Image)(resources.GetObject("btnNavBarSetting.Image")));
            this.btnNavBarSetting.Label = "Customize Nav Bar";
            this.btnNavBarSetting.Name = "btnNavBarSetting";
            this.btnNavBarSetting.ShowImage = true;
            this.btnNavBarSetting.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnNavBarSettings_Click);
            // 
            // linkbtn
            // 
            this.linkbtn.Image = ((System.Drawing.Image)(resources.GetObject("linkbtn.Image")));
            this.linkbtn.Label = "Add Link";
            this.linkbtn.Name = "linkbtn";
            this.linkbtn.ShowImage = true;
            this.linkbtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.linkbtn_Click);
            // 
            // adjustbtn
            // 
            this.adjustbtn.Image = ((System.Drawing.Image)(resources.GetObject("adjustbtn.Image")));
            this.adjustbtn.Label = "Adjust";
            this.adjustbtn.Name = "adjustbtn";
            this.adjustbtn.ShowImage = true;
            // 
            // ZoomGroup
            // 
            this.ZoomGroup.Items.Add(this.zoomaddbtn);
            this.ZoomGroup.Items.Add(this.selectzoombtn);
            this.ZoomGroup.Items.Add(this.createzoombtn);
            this.ZoomGroup.Label = "Zoom Features";
            this.ZoomGroup.Name = "ZoomGroup";
            // 
            // zoomaddbtn
            // 
            this.zoomaddbtn.Image = ((System.Drawing.Image)(resources.GetObject("zoomaddbtn.Image")));
            this.zoomaddbtn.Label = "Zoom  Add";
            this.zoomaddbtn.Name = "zoomaddbtn";
            this.zoomaddbtn.ShowImage = true;
            this.zoomaddbtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.zoomaddbtn_Click);
            // 
            // selectzoombtn
            // 
            this.selectzoombtn.Image = ((System.Drawing.Image)(resources.GetObject("selectzoombtn.Image")));
            this.selectzoombtn.Label = "Select Zoom Areas";
            this.selectzoombtn.Name = "selectzoombtn";
            this.selectzoombtn.ShowImage = true;
            this.selectzoombtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.selectzoombtn_Click);
            // 
            // createzoombtn
            // 
            this.createzoombtn.Image = ((System.Drawing.Image)(resources.GetObject("createzoombtn.Image")));
            this.createzoombtn.Label = "Create";
            this.createzoombtn.Name = "createzoombtn";
            this.createzoombtn.ShowImage = true;
            this.createzoombtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.createzoombtn_Click);
            // 
            // effectGruop
            // 
            this.effectGruop.Items.Add(this.effectsplitbtn);
            this.effectGruop.Items.Add(this.selectEffecctdtn);
            this.effectGruop.Label = "Effects";
            this.effectGruop.Name = "effectGruop";
            // 
            // effectsplitbtn
            // 
            this.effectsplitbtn.Image = ((System.Drawing.Image)(resources.GetObject("effectsplitbtn.Image")));
            this.effectsplitbtn.Items.Add(this.focusMenu);
            this.effectsplitbtn.Items.Add(this.spotlightMenu);
            this.effectsplitbtn.Items.Add(this.blurMenu);
            this.effectsplitbtn.Items.Add(this.MagnifyingGlassMenu);
            this.effectsplitbtn.Label = "Effects Menu";
            this.effectsplitbtn.Name = "effectsplitbtn";
            this.effectsplitbtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.effectsplitbtn_Click);
            // 
            // focusMenu
            // 
            this.focusMenu.Image = ((System.Drawing.Image)(resources.GetObject("focusMenu.Image")));
            this.focusMenu.Items.Add(this.focusbtn);
            this.focusMenu.Label = "Focus";
            this.focusMenu.Name = "focusMenu";
            this.focusMenu.ShowImage = true;
            // 
            // focusbtn
            // 
            this.focusbtn.Label = "Add Focus";
            this.focusbtn.Name = "focusbtn";
            this.focusbtn.ShowImage = true;
            // 
            // spotlightMenu
            // 
            this.spotlightMenu.Image = ((System.Drawing.Image)(resources.GetObject("spotlightMenu.Image")));
            this.spotlightMenu.Items.Add(this.createspotlightbtn);
            this.spotlightMenu.Items.Add(this.spotlightSettingsBtn);
            this.spotlightMenu.Label = "Spotlight";
            this.spotlightMenu.Name = "spotlightMenu";
            this.spotlightMenu.ShowImage = true;
            // 
            // createspotlightbtn
            // 
            this.createspotlightbtn.Label = "Create Spotlight";
            this.createspotlightbtn.Name = "createspotlightbtn";
            this.createspotlightbtn.ShowImage = true;
            this.createspotlightbtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.createspotlightbtn_Click);
            // 
            // spotlightSettingsBtn
            // 
            this.spotlightSettingsBtn.Label = "Settings";
            this.spotlightSettingsBtn.Name = "spotlightSettingsBtn";
            this.spotlightSettingsBtn.ShowImage = true;
            this.spotlightSettingsBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.spotlightSettingsBtn_Click);
            // 
            // blurMenu
            // 
            this.blurMenu.Image = ((System.Drawing.Image)(resources.GetObject("blurMenu.Image")));
            this.blurMenu.Label = "Blur";
            this.blurMenu.Name = "blurMenu";
            this.blurMenu.ShowImage = true;
            // 
            // MagnifyingGlassMenu
            // 
            this.MagnifyingGlassMenu.Image = ((System.Drawing.Image)(resources.GetObject("MagnifyingGlassMenu.Image")));
            this.MagnifyingGlassMenu.Items.Add(this.magniaddbtn);
            this.MagnifyingGlassMenu.Items.Add(this.magsetingsbtn);
            this.MagnifyingGlassMenu.Label = "Magnifying Glass";
            this.MagnifyingGlassMenu.Name = "MagnifyingGlassMenu";
            this.MagnifyingGlassMenu.ShowImage = true;
            // 
            // magniaddbtn
            // 
            this.magniaddbtn.Label = "Add Effect";
            this.magniaddbtn.Name = "magniaddbtn";
            this.magniaddbtn.ShowImage = true;
            this.magniaddbtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.magniaddbtn_Click);
            // 
            // magsetingsbtn
            // 
            this.magsetingsbtn.Label = "Settings";
            this.magsetingsbtn.Name = "magsetingsbtn";
            this.magsetingsbtn.ShowImage = true;
            this.magsetingsbtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.magsetingsbtn_Click);
            // 
            // selectEffecctdtn
            // 
            this.selectEffecctdtn.Image = ((System.Drawing.Image)(resources.GetObject("selectEffecctdtn.Image")));
            this.selectEffecctdtn.Label = "Select Effect Areas";
            this.selectEffecctdtn.Name = "selectEffecctdtn";
            this.selectEffecctdtn.ShowImage = true;
            this.selectEffecctdtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.selectEffecctdtn_Click);
            // 
            // hyperlinkgroup
            // 
            this.hyperlinkgroup.Items.Add(this.SlideNobox);
            this.hyperlinkgroup.Items.Add(this.crtHypBtn);
            this.hyperlinkgroup.Items.Add(this.rmvHypbtn);
            this.hyperlinkgroup.Items.Add(this.selecthypbtn);
            this.hyperlinkgroup.Label = "Hyper Link";
            this.hyperlinkgroup.Name = "hyperlinkgroup";
            // 
            // SlideNobox
            // 
            this.SlideNobox.Label = "Slide No :";
            this.SlideNobox.Name = "SlideNobox";
            this.SlideNobox.Text = null;
            // 
            // crtHypBtn
            // 
            this.crtHypBtn.Image = ((System.Drawing.Image)(resources.GetObject("crtHypBtn.Image")));
            this.crtHypBtn.Label = "Create Hyperlink";
            this.crtHypBtn.Name = "crtHypBtn";
            this.crtHypBtn.ShowImage = true;
            // 
            // rmvHypbtn
            // 
            this.rmvHypbtn.Image = ((System.Drawing.Image)(resources.GetObject("rmvHypbtn.Image")));
            this.rmvHypbtn.Label = "Remove Hyperlink";
            this.rmvHypbtn.Name = "rmvHypbtn";
            this.rmvHypbtn.ShowImage = true;
            // 
            // selecthypbtn
            // 
            this.selecthypbtn.Label = "Select";
            this.selecthypbtn.Name = "selecthypbtn";
            // 
            // MyRibbon
            // 
            this.Name = "MyRibbon";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MyRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.NavbarGruop.ResumeLayout(false);
            this.NavbarGruop.PerformLayout();
            this.ZoomGroup.ResumeLayout(false);
            this.ZoomGroup.PerformLayout();
            this.effectGruop.ResumeLayout(false);
            this.effectGruop.PerformLayout();
            this.hyperlinkgroup.ResumeLayout(false);
            this.hyperlinkgroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
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
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup ZoomGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup effectGruop;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton effectsplitbtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu focusMenu;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu spotlightMenu;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu blurMenu;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu MagnifyingGlassMenu;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton selectEffecctdtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup hyperlinkgroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox SlideNobox;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton crtHypBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton rmvHypbtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton selecthypbtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton focusbtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton createspotlightbtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton spotlightSettingsBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton magniaddbtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton magsetingsbtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton linkbtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton zoomaddbtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton selectzoombtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton createzoombtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton adjustbtn;
    }
}
