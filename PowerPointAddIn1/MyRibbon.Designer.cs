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
            this.group2 = this.Factory.CreateRibbonGroup();
            this.SubSectionStart = this.Factory.CreateRibbonEditBox();
            this.btnCreateSubsection = this.Factory.CreateRibbonButton();
            this.lblSubsectionName = this.Factory.CreateRibbonLabel();
            this.SubSectionEnd = this.Factory.CreateRibbonEditBox();
            this.valueSubsectionName = this.Factory.CreateRibbonLabel();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.addAgendaBtn = this.Factory.CreateRibbonButton();
            this.removeAgendaBtn = this.Factory.CreateRibbonButton();
            this.refreshAgendaBtn = this.Factory.CreateRibbonButton();
            this.NavbarGruop = this.Factory.CreateRibbonGroup();
            this.btnCreateNav = this.Factory.CreateRibbonButton();
            this.btnRefreshNav = this.Factory.CreateRibbonButton();
            this.btnRemoveNav = this.Factory.CreateRibbonButton();
            this.btnNavBarSetting = this.Factory.CreateRibbonButton();
            this.adjustbtn = this.Factory.CreateRibbonButton();
            this.linkbtn = this.Factory.CreateRibbonButton();
            this.ZoomGroup = this.Factory.CreateRibbonGroup();
            this.zoomTransparentSelectBtn = this.Factory.CreateRibbonButton();
            this.selectzoombtn = this.Factory.CreateRibbonButton();
            this.effectGruop = this.Factory.CreateRibbonGroup();
            this.spotlightSplitBtn = this.Factory.CreateRibbonSplitButton();
            this.spotlightSettingsBtn = this.Factory.CreateRibbonButton();
            this.blurSplitBtn = this.Factory.CreateRibbonSplitButton();
            this.blurModeRemainderToggle = this.Factory.CreateRibbonToggleButton();
            this.blurModeAllExceptToggle = this.Factory.CreateRibbonToggleButton();
            this.blursettingbtn = this.Factory.CreateRibbonButton();
            this.magnifySplitBtn = this.Factory.CreateRibbonSplitButton();
            this.magsetingsbtn = this.Factory.CreateRibbonButton();
            this.hyperlinkgroup = this.Factory.CreateRibbonGroup();
            this.SlideNobox = this.Factory.CreateRibbonEditBox();
            this.crtHypBtn = this.Factory.CreateRibbonButton();
            this.rmvHypbtn = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.btnToggleQR = this.Factory.CreateRibbonToggleButton();
            this.positionsLabGroup = this.Factory.CreateRibbonGroup();
            this.positionsLabBtn = this.Factory.CreateRibbonButton();
            this.Resize_group = this.Factory.CreateRibbonGroup();
            this.resizeBtn = this.Factory.CreateRibbonButton();
            this.lblSlideValue = this.Factory.CreateRibbonLabel();
            this.lblSectionValue = this.Factory.CreateRibbonLabel();
            this.lblTotalValue = this.Factory.CreateRibbonLabel();
            this.lblSlide = this.Factory.CreateRibbonLabel();
            this.lblSection = this.Factory.CreateRibbonLabel();
            this.lblTotal = this.Factory.CreateRibbonLabel();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.blur_remainbtn = this.Factory.CreateRibbonButton();
            this.blur_allexceptbtn = this.Factory.CreateRibbonButton();
            this.createspotlightbtn = this.Factory.CreateRibbonButton();
            this.blur_selectbtn = this.Factory.CreateRibbonButton();
            this.magniaddbtn = this.Factory.CreateRibbonButton();
            this.effectsplitbtn = this.Factory.CreateRibbonSplitButton();
            this.spotlightMenu = this.Factory.CreateRibbonMenu();
            this.blurMenu = this.Factory.CreateRibbonMenu();
            this.MagnifyingGlassMenu = this.Factory.CreateRibbonMenu();
            this.selectEffecctdtn = this.Factory.CreateRibbonButton();
            this.createzoombtn = this.Factory.CreateRibbonButton();
            this.zoomaddbtn = this.Factory.CreateRibbonButton();
            this.zoomTransparentBtn = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group4.SuspendLayout();
            this.NavbarGruop.SuspendLayout();
            this.ZoomGroup.SuspendLayout();
            this.effectGruop.SuspendLayout();
            this.hyperlinkgroup.SuspendLayout();
            this.group3.SuspendLayout();
            this.positionsLabGroup.SuspendLayout();
            this.Resize_group.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group4);
            this.tab1.Groups.Add(this.NavbarGruop);
            this.tab1.Groups.Add(this.ZoomGroup);
            this.tab1.Groups.Add(this.effectGruop);
            this.tab1.Groups.Add(this.hyperlinkgroup);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.positionsLabGroup);
            this.tab1.Groups.Add(this.Resize_group);
            this.tab1.Label = "My Tools";
            this.tab1.Name = "tab1";
            // 
            // group2
            // 
            this.group2.Items.Add(this.SubSectionStart);
            this.group2.Items.Add(this.btnCreateSubsection);
            this.group2.Items.Add(this.lblSubsectionName);
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
            // btnCreateSubsection
            // 
            this.btnCreateSubsection.Label = "Create Subsection";
            this.btnCreateSubsection.Name = "btnCreateSubsection";
            this.btnCreateSubsection.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCreateSubsection_Click);
            // 
            // lblSubsectionName
            // 
            this.lblSubsectionName.Label = "Current subsection name : ";
            this.lblSubsectionName.Name = "lblSubsectionName";
            // 
            // SubSectionEnd
            // 
            this.SubSectionEnd.Label = "End:";
            this.SubSectionEnd.Name = "SubSectionEnd";
            this.SubSectionEnd.Text = null;
            this.SubSectionEnd.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SubSectionEnd_TextChanged);
            // 
            // valueSubsectionName
            // 
            this.valueSubsectionName.Label = "Value";
            this.valueSubsectionName.Name = "valueSubsectionName";
            // 
            // group4
            // 
            this.group4.Items.Add(this.addAgendaBtn);
            this.group4.Items.Add(this.removeAgendaBtn);
            this.group4.Items.Add(this.refreshAgendaBtn);
            this.group4.Label = "Agenda";
            this.group4.Name = "group4";
            // 
            // addAgendaBtn
            // 
            this.addAgendaBtn.Label = "Add Agenda";
            this.addAgendaBtn.Name = "addAgendaBtn";
            this.addAgendaBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.addAgendaBtn_Click);
            // 
            // removeAgendaBtn
            // 
            this.removeAgendaBtn.Label = "Remove Agenda";
            this.removeAgendaBtn.Name = "removeAgendaBtn";
            this.removeAgendaBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.removeAgendaBtn_Click);
            // 
            // refreshAgendaBtn
            // 
            this.refreshAgendaBtn.Label = "Refresh Agenda";
            this.refreshAgendaBtn.Name = "refreshAgendaBtn";
            this.refreshAgendaBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.refreshAgendaBtn_Click);
            // 
            // NavbarGruop
            // 
            this.NavbarGruop.Items.Add(this.btnCreateNav);
            this.NavbarGruop.Items.Add(this.btnRefreshNav);
            this.NavbarGruop.Items.Add(this.btnRemoveNav);
            this.NavbarGruop.Items.Add(this.btnNavBarSetting);
            this.NavbarGruop.Items.Add(this.adjustbtn);
            this.NavbarGruop.Items.Add(this.linkbtn);
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
            // adjustbtn
            // 
            this.adjustbtn.Image = ((System.Drawing.Image)(resources.GetObject("adjustbtn.Image")));
            this.adjustbtn.Label = "Adjust";
            this.adjustbtn.Name = "adjustbtn";
            this.adjustbtn.ShowImage = true;
            this.adjustbtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.adjustbtn_Click);
            // 
            // linkbtn
            // 
            this.linkbtn.Image = ((System.Drawing.Image)(resources.GetObject("linkbtn.Image")));
            this.linkbtn.Label = "Add Link";
            this.linkbtn.Name = "linkbtn";
            this.linkbtn.ShowImage = true;
            this.linkbtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.linkbtn_Click);
            // 
            // ZoomGroup
            // 
            this.ZoomGroup.Items.Add(this.zoomTransparentSelectBtn);
            this.ZoomGroup.Items.Add(this.selectzoombtn);
            this.ZoomGroup.Label = "Zoom Features";
            this.ZoomGroup.Name = "ZoomGroup";
            // 
            // zoomTransparentSelectBtn
            // 
            this.zoomTransparentSelectBtn.Label = "Zoom Area";
            this.zoomTransparentSelectBtn.Name = "zoomTransparentSelectBtn";
            this.zoomTransparentSelectBtn.ShowImage = true;
            this.zoomTransparentSelectBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.zoomTransparentSelectBtn_Click);
            // 
            // selectzoombtn
            // 
            this.selectzoombtn.Image = ((System.Drawing.Image)(resources.GetObject("selectzoombtn.Image")));
            this.selectzoombtn.Label = "Select Object";
            this.selectzoombtn.Name = "selectzoombtn";
            this.selectzoombtn.ShowImage = true;
            this.selectzoombtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.selectzoombtn_Click);
            // 
            // effectGruop
            // 
            this.effectGruop.Items.Add(this.spotlightSplitBtn);
            this.effectGruop.Items.Add(this.blurSplitBtn);
            this.effectGruop.Items.Add(this.magnifySplitBtn);
            this.effectGruop.Label = "Effects";
            this.effectGruop.Name = "effectGruop";
            // 
            // spotlightSplitBtn
            // 
            this.spotlightSplitBtn.Image = ((System.Drawing.Image)(resources.GetObject("spotlightSplitBtn.Image")));
            this.spotlightSplitBtn.Items.Add(this.spotlightSettingsBtn);
            this.spotlightSplitBtn.Label = "Spotlight";
            this.spotlightSplitBtn.Name = "spotlightSplitBtn";
            this.spotlightSplitBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.spotlightSplitBtn_Click);
            // 
            // spotlightSettingsBtn
            // 
            this.spotlightSettingsBtn.Label = "Settings";
            this.spotlightSettingsBtn.Name = "spotlightSettingsBtn";
            this.spotlightSettingsBtn.ShowImage = true;
            this.spotlightSettingsBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.spotlightSettingsBtn_Click);
            // 
            // blurSplitBtn
            // 
            this.blurSplitBtn.Image = ((System.Drawing.Image)(resources.GetObject("blurSplitBtn.Image")));
            this.blurSplitBtn.Items.Add(this.blurModeRemainderToggle);
            this.blurSplitBtn.Items.Add(this.blurModeAllExceptToggle);
            this.blurSplitBtn.Items.Add(this.blursettingbtn);
            this.blurSplitBtn.Label = "Blur";
            this.blurSplitBtn.Name = "blurSplitBtn";
            this.blurSplitBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.blurSplitBtn_Click);
            // 
            // blurModeRemainderToggle
            // 
            this.blurModeRemainderToggle.Label = "Blur Remainder";
            this.blurModeRemainderToggle.Name = "blurModeRemainderToggle";
            this.blurModeRemainderToggle.ShowImage = true;
            this.blurModeRemainderToggle.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.blurModeRemainderToggle_Click);
            // 
            // blurModeAllExceptToggle
            // 
            this.blurModeAllExceptToggle.Label = "Blur Everything except selected";
            this.blurModeAllExceptToggle.Name = "blurModeAllExceptToggle";
            this.blurModeAllExceptToggle.ShowImage = true;
            this.blurModeAllExceptToggle.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.blurModeAllExceptToggle_Click);
            // 
            // blursettingbtn
            // 
            this.blursettingbtn.Label = "Settings";
            this.blursettingbtn.Name = "blursettingbtn";
            this.blursettingbtn.ShowImage = true;
            this.blursettingbtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.blursettingbtn_Click);
            // 
            // magnifySplitBtn
            // 
            this.magnifySplitBtn.Image = ((System.Drawing.Image)(resources.GetObject("magnifySplitBtn.Image")));
            this.magnifySplitBtn.Items.Add(this.magsetingsbtn);
            this.magnifySplitBtn.Label = "Magnifying Glass";
            this.magnifySplitBtn.Name = "magnifySplitBtn";
            this.magnifySplitBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.magnifySplitBtn_Click);
            // 
            // magsetingsbtn
            // 
            this.magsetingsbtn.Label = "Settings";
            this.magsetingsbtn.Name = "magsetingsbtn";
            this.magsetingsbtn.ShowImage = true;
            this.magsetingsbtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.magsetingsbtn_Click);
            // 
            // hyperlinkgroup
            // 
            this.hyperlinkgroup.Items.Add(this.SlideNobox);
            this.hyperlinkgroup.Items.Add(this.crtHypBtn);
            this.hyperlinkgroup.Items.Add(this.rmvHypbtn);
            this.hyperlinkgroup.Label = "Hyper Link";
            this.hyperlinkgroup.Name = "hyperlinkgroup";
            // 
            // SlideNobox
            // 
            this.SlideNobox.Label = "Slide No :";
            this.SlideNobox.Name = "SlideNobox";
            this.SlideNobox.Text = null;
            this.SlideNobox.TextChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SlideNobox_TextChanged);
            // 
            // crtHypBtn
            // 
            this.crtHypBtn.Image = ((System.Drawing.Image)(resources.GetObject("crtHypBtn.Image")));
            this.crtHypBtn.Label = "Create Hyperlink";
            this.crtHypBtn.Name = "crtHypBtn";
            this.crtHypBtn.ShowImage = true;
            this.crtHypBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.crtHypBtn_Click);
            // 
            // rmvHypbtn
            // 
            this.rmvHypbtn.Image = ((System.Drawing.Image)(resources.GetObject("rmvHypbtn.Image")));
            this.rmvHypbtn.Label = "Remove Hyperlink";
            this.rmvHypbtn.Name = "rmvHypbtn";
            this.rmvHypbtn.ShowImage = true;
            this.rmvHypbtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.rmvHypbtn_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.btnToggleQR);
            this.group3.Label = "QR Code";
            this.group3.Name = "group3";
            // 
            // btnToggleQR
            // 
            this.btnToggleQR.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnToggleQR.Image = ((System.Drawing.Image)(resources.GetObject("btnToggleQR.Image")));
            this.btnToggleQR.Label = "QR Code Pane";
            this.btnToggleQR.Name = "btnToggleQR";
            this.btnToggleQR.ShowImage = true;
            this.btnToggleQR.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnToggleQR_Click);
            // 
            // positionsLabGroup
            // 
            this.positionsLabGroup.Items.Add(this.positionsLabBtn);
            this.positionsLabGroup.Label = "Allignment";
            this.positionsLabGroup.Name = "positionsLabGroup";
            // 
            // positionsLabBtn
            // 
            this.positionsLabBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.positionsLabBtn.Image = ((System.Drawing.Image)(resources.GetObject("positionsLabBtn.Image")));
            this.positionsLabBtn.Label = "Positions Lab";
            this.positionsLabBtn.Name = "positionsLabBtn";
            this.positionsLabBtn.ShowImage = true;
            this.positionsLabBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.positionsLabBtn_Click);
            // 
            // Resize_group
            // 
            this.Resize_group.Items.Add(this.resizeBtn);
            this.Resize_group.Label = "Resize Lab";
            this.Resize_group.Name = "Resize_group";
            // 
            // resizeBtn
            // 
            this.resizeBtn.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.resizeBtn.Image = ((System.Drawing.Image)(resources.GetObject("resizeBtn.Image")));
            this.resizeBtn.Label = "Resize Lab";
            this.resizeBtn.Name = "resizeBtn";
            this.resizeBtn.ShowImage = true;
            this.resizeBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.resizeBtn_Click);
            // 
            // lblSlideValue
            // 
            this.lblSlideValue.Label = "Value";
            this.lblSlideValue.Name = "lblSlideValue";
            // 
            // lblSectionValue
            // 
            this.lblSectionValue.Label = "Value";
            this.lblSectionValue.Name = "lblSectionValue";
            // 
            // lblTotalValue
            // 
            this.lblTotalValue.Label = "Value";
            this.lblTotalValue.Name = "lblTotalValue";
            // 
            // lblSlide
            // 
            this.lblSlide.Label = "Slides in Section:";
            this.lblSlide.Name = "lblSlide";
            // 
            // lblSection
            // 
            this.lblSection.Label = "Current Section:";
            this.lblSection.Name = "lblSection";
            // 
            // lblTotal
            // 
            this.lblTotal.Label = "Total Sections:";
            this.lblTotal.Name = "lblTotal";
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
            // blur_remainbtn
            // 
            this.blur_remainbtn.Label = "Blur Remainder";
            this.blur_remainbtn.Name = "blur_remainbtn";
            this.blur_remainbtn.ShowImage = true;
            this.blur_remainbtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.blur_remainbtn_Click);
            // 
            // blur_allexceptbtn
            // 
            this.blur_allexceptbtn.Label = "Blur Everything except selected";
            this.blur_allexceptbtn.Name = "blur_allexceptbtn";
            this.blur_allexceptbtn.ShowImage = true;
            this.blur_allexceptbtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.blur_allexceptbtn_Click);
            // 
            // createspotlightbtn
            // 
            this.createspotlightbtn.Label = "Spotlight";
            this.createspotlightbtn.Name = "createspotlightbtn";
            this.createspotlightbtn.ShowImage = true;
            this.createspotlightbtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.createspotlightbtn_Click);
            // 
            // blur_selectbtn
            // 
            this.blur_selectbtn.Label = "Blur";
            this.blur_selectbtn.Name = "blur_selectbtn";
            this.blur_selectbtn.ShowImage = true;
            this.blur_selectbtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.blur_selectbtn_Click);
            // 
            // magniaddbtn
            // 
            this.magniaddbtn.Label = "Magnifying Glass";
            this.magniaddbtn.Name = "magniaddbtn";
            this.magniaddbtn.ShowImage = true;
            this.magniaddbtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.magniaddbtn_Click);
            // 
            // effectsplitbtn
            // 
            this.effectsplitbtn.Image = ((System.Drawing.Image)(resources.GetObject("effectsplitbtn.Image")));
            this.effectsplitbtn.Items.Add(this.spotlightMenu);
            this.effectsplitbtn.Items.Add(this.blurMenu);
            this.effectsplitbtn.Items.Add(this.MagnifyingGlassMenu);
            this.effectsplitbtn.Label = "Effects Menu";
            this.effectsplitbtn.Name = "effectsplitbtn";
            this.effectsplitbtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.effectsplitbtn_Click);
            // 
            // spotlightMenu
            // 
            this.spotlightMenu.Image = ((System.Drawing.Image)(resources.GetObject("spotlightMenu.Image")));
            this.spotlightMenu.Label = "Spotlight";
            this.spotlightMenu.Name = "spotlightMenu";
            this.spotlightMenu.ShowImage = true;
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
            this.MagnifyingGlassMenu.Label = "Magnifying Glass";
            this.MagnifyingGlassMenu.Name = "MagnifyingGlassMenu";
            this.MagnifyingGlassMenu.ShowImage = true;
            // 
            // selectEffecctdtn
            // 
            this.selectEffecctdtn.Image = ((System.Drawing.Image)(resources.GetObject("selectEffecctdtn.Image")));
            this.selectEffecctdtn.Label = "Select Effect Areas";
            this.selectEffecctdtn.Name = "selectEffecctdtn";
            this.selectEffecctdtn.ShowImage = true;
            this.selectEffecctdtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.selectEffecctdtn_Click);
            // 
            // createzoombtn
            // 
            this.createzoombtn.Image = ((System.Drawing.Image)(resources.GetObject("createzoombtn.Image")));
            this.createzoombtn.Label = "Add Zoom";
            this.createzoombtn.Name = "createzoombtn";
            this.createzoombtn.ShowImage = true;
            this.createzoombtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.createzoombtn_Click);
            // 
            // zoomaddbtn
            // 
            this.zoomaddbtn.Image = ((System.Drawing.Image)(resources.GetObject("zoomaddbtn.Image")));
            this.zoomaddbtn.Label = "Zoom  Add";
            this.zoomaddbtn.Name = "zoomaddbtn";
            this.zoomaddbtn.ShowImage = true;
            this.zoomaddbtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.zoomaddbtn_Click);
            // 
            // zoomTransparentBtn
            // 
            this.zoomTransparentBtn.Label = "Transparent Selected";
            this.zoomTransparentBtn.Name = "zoomTransparentBtn";
            this.zoomTransparentBtn.ShowImage = true;
            this.zoomTransparentBtn.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.zoomTransparentBtn_Click);
            // 
            // MyRibbon
            // 
            this.Name = "MyRibbon";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MyRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.NavbarGruop.ResumeLayout(false);
            this.NavbarGruop.PerformLayout();
            this.ZoomGroup.ResumeLayout(false);
            this.ZoomGroup.PerformLayout();
            this.effectGruop.ResumeLayout(false);
            this.effectGruop.PerformLayout();
            this.hyperlinkgroup.ResumeLayout(false);
            this.hyperlinkgroup.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.positionsLabGroup.ResumeLayout(false);
            this.positionsLabGroup.PerformLayout();
            this.Resize_group.ResumeLayout(false);
            this.Resize_group.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
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
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton spotlightSplitBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton blurSplitBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton magnifySplitBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton effectsplitbtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu spotlightMenu;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu blurMenu;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu MagnifyingGlassMenu;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton selectEffecctdtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup hyperlinkgroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonEditBox SlideNobox;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton crtHypBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton rmvHypbtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton createspotlightbtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton spotlightSettingsBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton magniaddbtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton magsetingsbtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton linkbtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton zoomTransparentBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton zoomTransparentSelectBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton zoomaddbtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton selectzoombtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton createzoombtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton adjustbtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton blurModeRemainderToggle;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton blurModeAllExceptToggle;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton blur_remainbtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton blur_selectbtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton blur_allexceptbtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton blursettingbtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton btnToggleQR;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton addAgendaBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton removeAgendaBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton refreshAgendaBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup positionsLabGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton positionsLabBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton resizeBtn;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Resize_group;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lblTotal;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lblSection;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lblSlide;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lblTotalValue;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lblSectionValue;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lblSlideValue;
    }
}
