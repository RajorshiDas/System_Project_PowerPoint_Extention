using System;
using System.Drawing;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PowerPointAddIn1
{
    public class NavBarColorSettings : Form
    {
        private NavBarSettings settings;
        private PowerPoint.Slide previewSlide = null;
        private bool previewActive = false;
        
        public NavBarSettings Settings => settings;

        public NavBarColorSettings(NavBarSettings currentSettings)
        {
            settings = new NavBarSettings
            {
                BackgroundColor = currentSettings.BackgroundColor,
                SectionNameColor = currentSettings.SectionNameColor,
                CurrentSectionNameColor = currentSettings.CurrentSectionNameColor,
                OtherSectionNameColor = currentSettings.OtherSectionNameColor,
                CurrentSlideColor = currentSettings.CurrentSlideColor,
                SameSubsectionBorderColor = currentSettings.SameSubsectionBorderColor,
                SameSubsectionFillColor = currentSettings.SameSubsectionFillColor,
                OtherSlidesBorderColor = currentSettings.OtherSlidesBorderColor,
                SubsectionBoxTransparency = currentSettings.SubsectionBoxTransparency,
                SlideShapeType = currentSettings.SlideShapeType,
                ShowSlideNumbers = currentSettings.ShowSlideNumbers,
                SlideNumberColor = currentSettings.SlideNumberColor
            };

            InitializeComponents();
            
            // Handle form closing to clean up preview
            this.FormClosing += NavBarColorSettings_FormClosing;
        }

        private void NavBarColorSettings_FormClosing(object sender, FormClosingEventArgs e)
        {
            // If preview was active and user clicks Cancel, remove it
            if (previewActive && this.DialogResult == DialogResult.Cancel && previewSlide != null)
            {
                RemovePreviewFromSlide(previewSlide);
            }
        }

        private void InitializeComponents()
        {
            this.Text = "Customize Navigation Bar";
            this.Size = new Size(450, 660);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;

            int y = 20;
            
            // === SHAPE OPTIONS ===
            Label lblShapeHeader = new Label 
            { 
                Text = "SHAPE OPTIONS", 
                Location = new Point(20, y), 
                Size = new Size(400, 20),
                Font = new Font(this.Font, FontStyle.Bold)
            };
            this.Controls.Add(lblShapeHeader);
            y += 30;

            // Shape Type: Circle or Square
            Label lblShape = new Label { Text = "Shape Type:", Location = new Point(20, y + 5), Size = new Size(180, 20) };
            ComboBox cmbShape = new ComboBox
            {
                Location = new Point(210, y),
                Size = new Size(180, 30),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cmbShape.Items.AddRange(new object[] { "Circle", "Square", "Number Only" });
            cmbShape.SelectedIndex = settings.SlideShapeType == NavBarSettings.ShapeType.Circle ? 0 : 
                                     settings.SlideShapeType == NavBarSettings.ShapeType.Square ? 1 : 2;
            cmbShape.SelectedIndexChanged += (s, e) => {
                settings.SlideShapeType = cmbShape.SelectedIndex == 0 ? NavBarSettings.ShapeType.Circle :
                                         cmbShape.SelectedIndex == 1 ? NavBarSettings.ShapeType.Square :
                                         NavBarSettings.ShapeType.NumberOnly;
            };
            this.Controls.AddRange(new Control[] { lblShape, cmbShape });
            y += 45;

            // Show Slide Numbers
            CheckBox chkShowNumbers = new CheckBox
            {
                Text = "Show Slide Numbers",
                Location = new Point(20, y),
                Size = new Size(200, 30),
                Checked = settings.ShowSlideNumbers
            };
            chkShowNumbers.CheckedChanged += (s, e) => settings.ShowSlideNumbers = chkShowNumbers.Checked;
            this.Controls.Add(chkShowNumbers);
            y += 45;

            // === COLOR OPTIONS ===
            Label lblColorHeader = new Label 
            { 
                Text = "COLOR OPTIONS", 
                Location = new Point(20, y), 
                Size = new Size(400, 20),
                Font = new Font(this.Font, FontStyle.Bold)
            };
            this.Controls.Add(lblColorHeader);
            y += 30;

            AddColorOption("Background Color:", settings.BackgroundColor, y, (c) => settings.BackgroundColor = c); y += 45;
            AddColorOption("Current Section Name:", settings.CurrentSectionNameColor, y, (c) => settings.CurrentSectionNameColor = c); y += 45;
            AddColorOption("Other Section Names:", settings.OtherSectionNameColor, y, (c) => settings.OtherSectionNameColor = c); y += 45;
            AddColorOption("Current Slide:", settings.CurrentSlideColor, y, (c) => settings.CurrentSlideColor = c); y += 45;
            AddColorOption("Slide Numbers:", settings.SlideNumberColor, y, (c) => settings.SlideNumberColor = c); y += 45;
            AddColorOption("Same Subsection Border:", settings.SameSubsectionBorderColor, y, (c) => settings.SameSubsectionBorderColor = c); y += 45;
            AddColorOption("Same Subsection Fill:", settings.SameSubsectionFillColor, y, (c) => settings.SameSubsectionFillColor = c); y += 45;
            AddColorOption("Other Slides Border:", settings.OtherSlidesBorderColor, y, (c) => settings.OtherSlidesBorderColor = c); y += 45;

            // Buttons
            Button btnPreview = new Button { Text = "Preview", Location = new Point(150, y), Size = new Size(90, 30) };
            btnPreview.Click += (s, e) => ShowPreview();
            
            Button btnReset = new Button { Text = "Reset Defaults", Location = new Point(20, y), Size = new Size(120, 30) };
            btnReset.Click += (s, e) => { 
                settings.ResetToDefaults(); 
                this.Close(); 
                new NavBarColorSettings(settings).ShowDialog(); 
            };

            Button btnOK = new Button { Text = "OK", Location = new Point(250, y), Size = new Size(80, 30), DialogResult = DialogResult.OK };
            Button btnCancel = new Button { Text = "Cancel", Location = new Point(340, y), Size = new Size(80, 30), DialogResult = DialogResult.Cancel };
            this.Controls.AddRange(new Control[] { btnPreview, btnReset, btnOK, btnCancel });
            this.AcceptButton = btnOK;
            this.CancelButton = btnCancel;
        }

        private void ShowPreview()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                
                if (app.ActivePresentation == null)
                {
                    MessageBox.Show("Please open a presentation first.", "No Presentation",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var presentation = app.ActivePresentation;
                var sections = presentation.SectionProperties;

                if (sections.Count == 0)
                {
                    MessageBox.Show("Please add sections to your presentation first.", "No Sections",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                try
                {
                    previewSlide = app.ActiveWindow.View.Slide;
                }
                catch
                {
                    MessageBox.Show("Please select a slide first.", "No Slide Selected",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Remove old nav bar and preview
                RemoveOldNavBar(previewSlide);
                RemovePreviewFromSlide(previewSlide);

                // Add preview nav bar
                AddPreviewNavigationBar(previewSlide, sections, presentation, settings);
                previewActive = true;

                MessageBox.Show("Preview applied to current slide!\n\n" +
                    "The preview shows your new settings.\n" +
                    "Click 'Preview' again to update after changing settings.\n" +
                    "Click 'OK' to apply to all slides.\n" +
                    "Click 'Cancel' to discard and restore original.",
                    "Preview Active", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error showing preview: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void RemoveOldNavBar(PowerPoint.Slide slide)
        {
            for (int i = slide.Shapes.Count; i >= 1; i--)
            {
                try
                {
                    if (slide.Shapes[i].Tags["NavBar"] == "True")
                    {
                        slide.Shapes[i].Delete();
                    }
                }
                catch { }
            }
        }

        private void RemovePreviewFromSlide(PowerPoint.Slide slide)
        {
            for (int i = slide.Shapes.Count; i >= 1; i--)
            {
                try
                {
                    if (slide.Shapes[i].Tags["NavBarPreview"] == "True")
                    {
                        slide.Shapes[i].Delete();
                    }
                }
                catch { }
            }
        }

        private void AddPreviewNavigationBar(PowerPoint.Slide slide, PowerPoint.SectionProperties sections, 
            PowerPoint.Presentation presentation, NavBarSettings previewSettings)
        {
            // This is a simplified version - you can copy the full logic from MyRibbon.cs AddNavigationBarToSlide
            float barHeight = 60;
            float slideWidth = presentation.PageSetup.SlideWidth;
            
            // Create background bar
            PowerPoint.Shape navBackground = slide.Shapes.AddShape(
                Office.MsoAutoShapeType.msoShapeRectangle,
                0, 0, slideWidth, barHeight);
            navBackground.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(previewSettings.BackgroundColor);
            navBackground.Line.Visible = Office.MsoTriState.msoFalse;
            navBackground.Tags.Add("NavBarPreview", "True");

            float currentX = 20;
            float topY = 8;
            float circleSize = 12;
            float circleSpacing = 6;

            int currentSlideIndex = slide.SlideIndex;
            int currentSectionIndex = 0;
            for (int si = 1; si <= sections.Count; si++)
            {
                int firstSlide = sections.FirstSlide(si);
                int slideCount = sections.SlidesCount(si);
                int lastSlide = firstSlide + slideCount - 1;

                if (currentSlideIndex >= firstSlide && currentSlideIndex <= lastSlide)
                {
                    currentSectionIndex = si;
                    break;
                }
            }

            for (int i = 1; i <= sections.Count; i++)
            {
                string sectionName = sections.Name(i);
                int slideCountInSection = sections.SlidesCount(i);
                int firstSlideInSection = sections.FirstSlide(i);

                // Add section name
                PowerPoint.Shape sectionLabel = slide.Shapes.AddTextbox(
                    Office.MsoTextOrientation.msoTextOrientationHorizontal,
                    currentX, topY, 200, 20);
                sectionLabel.TextFrame.TextRange.Text = sectionName;
                var sectionNameColor = i == currentSectionIndex
                    ? previewSettings.CurrentSectionNameColor
                    : previewSettings.OtherSectionNameColor;
                sectionLabel.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(sectionNameColor);
                sectionLabel.TextFrame.TextRange.Font.Size = 12;
                sectionLabel.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                sectionLabel.Line.Visible = Office.MsoTriState.msoFalse;
                sectionLabel.Fill.Visible = Office.MsoTriState.msoFalse;
                sectionLabel.Tags.Add("NavBarPreview", "True");

                float circleY = topY + 25;
                float circleX = currentX;

                // Draw sample shapes
                for (int j = 0; j < Math.Min(slideCountInSection, 10); j++) // Show max 10 for preview
                {
                    int slideIndexInPresentation = firstSlideInSection + j;
                    PowerPoint.Shape shape = null;
                    
                    if (previewSettings.SlideShapeType == NavBarSettings.ShapeType.NumberOnly)
                    {
                        shape = slide.Shapes.AddTextbox(
                            Office.MsoTextOrientation.msoTextOrientationHorizontal,
                            circleX, circleY, circleSize, circleSize);
                        shape.Line.Visible = Office.MsoTriState.msoFalse;
                        shape.Fill.Visible = Office.MsoTriState.msoFalse;
                    }
                    else if (previewSettings.SlideShapeType == NavBarSettings.ShapeType.Square)
                    {
                        shape = slide.Shapes.AddShape(
                            Office.MsoAutoShapeType.msoShapeRectangle,
                            circleX, circleY, circleSize, circleSize);
                    }
                    else
                    {
                        shape = slide.Shapes.AddShape(
                            Office.MsoAutoShapeType.msoShapeOval,
                            circleX, circleY, circleSize, circleSize);
                    }

                    if (previewSettings.SlideShapeType != NavBarSettings.ShapeType.NumberOnly)
                    {
                        if (slideIndexInPresentation == currentSlideIndex)
                        {
                            shape.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(previewSettings.CurrentSlideColor);
                            shape.Line.Visible = Office.MsoTriState.msoFalse;
                        }
                        else
                        {
                            shape.Fill.Visible = Office.MsoTriState.msoFalse;
                            shape.Line.Visible = Office.MsoTriState.msoTrue;
                            shape.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(previewSettings.OtherSlidesBorderColor);
                            shape.Line.Weight = 1.5f;
                        }
                    }

                    if (previewSettings.ShowSlideNumbers || previewSettings.SlideShapeType == NavBarSettings.ShapeType.NumberOnly)
                    {
                        shape.TextFrame.TextRange.Text = slideIndexInPresentation.ToString();
                        shape.TextFrame.TextRange.Font.Size = previewSettings.SlideShapeType == NavBarSettings.ShapeType.NumberOnly ? 10 : 8;
                        shape.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(previewSettings.SlideNumberColor);
                        shape.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                        shape.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;
                        shape.TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;
                        shape.TextFrame.MarginLeft = 0;
                        shape.TextFrame.MarginRight = 0;
                        shape.TextFrame.MarginTop = 0;
                        shape.TextFrame.MarginBottom = 0;
                    }

                    shape.Tags.Add("NavBarPreview", "True");
                    circleX += circleSize + circleSpacing;
                }

                currentX += Math.Max(150, (circleSize + circleSpacing) * Math.Min(slideCountInSection, 10)) + 30;
            }
        }

        private void AddColorOption(string label, Color color, int y, Action<Color> updateColor)
        {
            Label lbl = new Label { Text = label, Location = new Point(20, y + 5), Size = new Size(180, 20) };
            Panel preview = new Panel { BackColor = color, Location = new Point(210, y), Size = new Size(80, 30), BorderStyle = BorderStyle.FixedSingle };
            Button btn = new Button { Text = "Change", Location = new Point(300, y), Size = new Size(80, 30) };

            btn.Click += (s, e) => {
                using (ColorDialog dlg = new ColorDialog { Color = preview.BackColor, FullOpen = true })
                {
                    if (dlg.ShowDialog() == DialogResult.OK)
                    {
                        preview.BackColor = dlg.Color;
                        updateColor(dlg.Color);
                    }
                }
            };

            this.Controls.AddRange(new Control[] { lbl, preview, btn });
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            // 
            // NavBarColorSettings
            // 
            this.ClientSize = new System.Drawing.Size(282, 253);
            this.Name = "NavBarColorSettings";
            this.Load += new System.EventHandler(this.NavBarColorSettings_Load);
            this.ResumeLayout(false);

        }

        private void NavBarColorSettings_Load(object sender, EventArgs e)
        {

        }
    }
}
