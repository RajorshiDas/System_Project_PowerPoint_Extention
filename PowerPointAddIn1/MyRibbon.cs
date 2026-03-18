using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using QRCoder;

namespace PowerPointAddIn1
{
    public partial class MyRibbon
    {
        // Navigation Bar Customization Settings
        private NavBarSettings navBarSettings = new NavBarSettings();

        // Hyperlink feature state
        private int _hypSourceSlideIndex = -1;
        private readonly List<string> _hypShapeNames = new List<string>();

        private void MyRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            UpdateSectionInfo();
        }

        public void RefreshInfo()
        {
            UpdateSectionInfo();
        }

        private void UpdateSectionInfo()
        {
            try
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;

                // ActivePresentation throws a COM exception when no file is open,
                // so check Presentations.Count first to avoid that.
                if (app.Presentations.Count == 0)
                {
                    lblTotalValue.Label = "N/A";
                    lblSectionValue.Label = "N/A";
                    lblSlideValue.Label = "N/A";
                    valueSubsectionName.Label = "N/A";
                    return;
                }

                PowerPoint.Presentation presentation = app.ActivePresentation;
                PowerPoint.SectionProperties sections = presentation.SectionProperties;

                int totalSections = sections.Count;
                lblTotalValue.Label = totalSections.ToString();

                if (totalSections > 0)
                {
                    try
                    {
                        PowerPoint.Slide activeSlide = app.ActiveWindow.View.Slide;
                        int activeSlideIndex = activeSlide.SlideIndex;
                        int sectionIndex = GetSectionIndexForSlide(sections, activeSlideIndex);

                        if (sectionIndex > 0)
                        {
                            string currentSectionName = sections.Name(sectionIndex);
                            int slidesInSection = sections.SlidesCount(sectionIndex);
                            int firstSlideInSection = sections.FirstSlide(sectionIndex);

                            lblSectionValue.Label = currentSectionName;
                            lblSlideValue.Label = slidesInSection.ToString();

                            string subsectionTag = activeSlide.Tags["Subsection"];
                            int subsectionCount = CountSubsectionsInSection(presentation, firstSlideInSection, slidesInSection);

                            if (string.IsNullOrEmpty(subsectionTag))
                            {
                                valueSubsectionName.Label = $"None ({subsectionCount} total)";
                            }
                            else
                            {
                                valueSubsectionName.Label = $"{subsectionTag} ({subsectionCount} total)";
                            }
                        }
                        else
                        {
                            lblSectionValue.Label = "No section";
                            lblSlideValue.Label = "N/A";
                            valueSubsectionName.Label = "None";
                        }
                    }
                    catch
                    {
                        lblSectionValue.Label = "No slide selected";
                        lblSlideValue.Label = "N/A";
                        valueSubsectionName.Label = "N/A";
                    }
                }
                else
                {
                    lblSectionValue.Label = "No sections";
                    lblSlideValue.Label = presentation.Slides.Count.ToString();

                    try
                    {
                        PowerPoint.Slide activeSlide = app.ActiveWindow.View.Slide;
                        string subsectionTag = activeSlide.Tags["Subsection"];
                        int subsectionCount = CountSubsectionsInSection(presentation, 1, presentation.Slides.Count);

                        if (string.IsNullOrEmpty(subsectionTag))
                        {
                            valueSubsectionName.Label = $"None ({subsectionCount} total)";
                        }
                        else
                        {
                            valueSubsectionName.Label = $"{subsectionTag} ({subsectionCount} total)";
                        }
                    }
                    catch
                    {
                        valueSubsectionName.Label = "N/A";
                    }
                }
            }
            catch (Exception ex)
            {
                // If there is simply no active presentation (e.g. at startup),
                // reset labels quietly without showing an error dialog.
                if (ex.Message.IndexOf("no active presentation", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    ex.Message.IndexOf("Invalid request", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    lblTotalValue.Label = "N/A";
                    lblSectionValue.Label = "N/A";
                    lblSlideValue.Label = "N/A";
                    valueSubsectionName.Label = "N/A";
                    return;
                }

                lblTotalValue.Label = "Error";
                lblSectionValue.Label = "Error";
                lblSlideValue.Label = "Error";
                valueSubsectionName.Label = "Error";
                MessageBox.Show(
                    "Error updating section info: " + ex.Message,
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private int CountSubsectionsInSection(PowerPoint.Presentation presentation, int firstSlide, int slideCount)
        {
            HashSet<string> uniqueSubsections = new HashSet<string>();

            for (int i = firstSlide; i < firstSlide + slideCount; i++)
            {
                try
                {
                    PowerPoint.Slide slide = presentation.Slides[i];
                    string subsectionTag = slide.Tags["Subsection"];

                    if (!string.IsNullOrEmpty(subsectionTag))
                    {
                        uniqueSubsections.Add(subsectionTag);
                    }
                }
                catch
                {
                    continue;
                }
            }

            return uniqueSubsections.Count;
        }

        private int GetSectionIndexForSlide(PowerPoint.SectionProperties sections, int slideIndex)
        {
            for (int i = 1; i <= sections.Count; i++)
            {
                int firstSlide = sections.FirstSlide(i);
                int slideCount = sections.SlidesCount(i);
                int lastSlide = firstSlide + slideCount - 1;

                if (slideIndex >= firstSlide && slideIndex <= lastSlide)
                {
                    return i;
                }
            }
            return 0;
        }

        private void btnCreateSubsection_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;

                if (app.ActivePresentation == null)
                {
                    MessageBox.Show(
                        "Please open a presentation first.",
                        "No Presentation",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                PowerPoint.Presentation presentation = app.ActivePresentation;

                // Get start and end slide numbers from editboxes
                string startText = SubSectionStart.Text.Trim();
                string endText = SubSectionEnd.Text.Trim();

                if (string.IsNullOrEmpty(startText) || string.IsNullOrEmpty(endText))
                {
                    MessageBox.Show(
                        "Please enter both start and end slide numbers.",
                        "Missing Input",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                // Parse slide numbers
                if (!int.TryParse(startText, out int startSlide) || !int.TryParse(endText, out int endSlide))
                {
                    MessageBox.Show(
                        "Please enter valid slide numbers.",
                        "Invalid Input",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                // Validate slide range
                int totalSlides = presentation.Slides.Count;
                if (startSlide < 1 || endSlide > totalSlides || startSlide > endSlide)
                {
                    MessageBox.Show(
                        $"Invalid slide range. Please enter numbers between 1 and {totalSlides}, with start <= end.",
                        "Invalid Range",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                // Get subsection name
                string subsectionName = InputDialog.Show("Enter subsection name:", "Create Subsection");

                if (string.IsNullOrEmpty(subsectionName))
                {
                    return;
                }

                // Apply subsection tag to all slides in range
                int slideCount = 0;
                for (int i = startSlide; i <= endSlide; i++)
                {
                    PowerPoint.Slide slide = presentation.Slides[i];

                    if (slide.Tags["Subsection"] != "")
                    {
                        slide.Tags.Delete("Subsection");
                    }

                    slide.Tags.Add("Subsection", subsectionName);
                    slideCount++;
                }

                // Update the display
                UpdateSectionInfo();

                // Clear the input boxes
                SubSectionStart.Text = "";
                SubSectionEnd.Text = "";

                MessageBox.Show(
                    $"Successfully added subsection '{subsectionName}' to slides {startSlide} to {endSlide} ({slideCount} slide(s)).",
                    "Success",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Error creating subsection: " + ex.Message,
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void CreateSectionNavigationBar()
        {
            try
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;

                if (app.ActivePresentation == null)
                {
                    MessageBox.Show("Please open a presentation first.", "No Presentation",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                PowerPoint.Presentation presentation = app.ActivePresentation;
                PowerPoint.SectionProperties sections = presentation.SectionProperties;

                if (sections.Count == 0)
                {
                    MessageBox.Show("This presentation has no sections. Please add sections first.",
                        "No Sections", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                int slidesProcessed = 0;
                int totalSubsections = 0;

                // Count total subsections
                foreach (PowerPoint.Slide checkSlide in presentation.Slides)
                {
                    try
                    {
                        string subsectionTag = checkSlide.Tags["Subsection"];
                        if (!string.IsNullOrEmpty(subsectionTag))
                        {
                            totalSubsections++;
                        }
                    }
                    catch { }
                }

                // Loop through all slides in the presentation
                foreach (PowerPoint.Slide slide in presentation.Slides)
                {
                    // Remove old navigation bar if it exists
                    RemoveNavigationBar(slide);

                    // Add the navigation bar to the top of the slide
                    AddNavigationBarToSlide(slide, sections, presentation);
                    slidesProcessed++;
                }

                string message = $"Navigation bar added to {slidesProcessed} slide(s)!";
                if (totalSubsections > 0)
                {
                    message += $"\n{totalSubsections} slide(s) have subsections with colored backgrounds.";
                }
                else
                {
                    message += "\n\nNote: No subsections found. Create subsections to see colored grouping.";
                }

                MessageBox.Show(message, "Success",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error creating navigation bar: {ex.Message}\n\nStack: {ex.StackTrace}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void RemoveNavigationBar(PowerPoint.Slide slide)
        {
            // Restore content shapes to their pre-nav-bar geometry
            for (int i = 1; i <= slide.Shapes.Count; i++)
            {
                try
                {
                    var sh = slide.Shapes[i];
                    if (sh.Tags["NavBar"] == "True") continue;

                    string savedTop = sh.Tags["NavBarOrigTop"];
                    if (string.IsNullOrEmpty(savedTop)) continue;

                    sh.Left = float.Parse(sh.Tags["NavBarOrigLeft"], System.Globalization.CultureInfo.InvariantCulture);
                    sh.Top = float.Parse(savedTop, System.Globalization.CultureInfo.InvariantCulture);
                    sh.Width = float.Parse(sh.Tags["NavBarOrigWidth"], System.Globalization.CultureInfo.InvariantCulture);
                    sh.Height = float.Parse(sh.Tags["NavBarOrigHeight"], System.Globalization.CultureInfo.InvariantCulture);

                    try { sh.Tags.Delete("NavBarOrigLeft"); } catch { }
                    try { sh.Tags.Delete("NavBarOrigTop"); } catch { }
                    try { sh.Tags.Delete("NavBarOrigWidth"); } catch { }
                    try { sh.Tags.Delete("NavBarOrigHeight"); } catch { }
                }
                catch { }
            }

            // Remove shapes tagged as navigation bar
            for (int i = slide.Shapes.Count; i >= 1; i--)
            {
                try
                {
                    if (slide.Shapes[i].Tags["NavBar"] == "True")
                        slide.Shapes[i].Delete();
                }
                catch { }
            }
        }

        private void AddNavigationBarToSlide(PowerPoint.Slide slide, PowerPoint.SectionProperties sections, PowerPoint.Presentation presentation)
        {
            float slideWidth  = presentation.PageSetup.SlideWidth;
            float slideHeight = presentation.PageSetup.SlideHeight;

            // Layout constants
            // Each section is a column:   SECTION NAME
            //                             ● ● ● ● ●
            // Columns are packed left-to-right and wrap to a new band when they overflow.
            const float startX          = 10f;
            const float baseCircleSize  = 12f;
            const float circleSpacing   = 5f;
            const float sectionGap      = 14f;   // horizontal gap between section columns
            const float labelFontSize   = 12f;
            const float labelHeight     = 16f;   // textbox height for the section name
            const float labelCircleGap  = 3f;    // vertical gap between name and circles
            const float bandPadTop      = 4f;    // padding above name inside each band
            const float bandPadBottom   = 4f;    // padding below circles inside each band
            // Approximate pt-width per character at 12pt bold Calibri
            float approxCharWidth = labelFontSize * 0.65f;

            int   sectionCount   = sections.Count;

            // Dynamically size circles so multi-digit slide numbers (10, 11, 12…)
            // fit horizontally instead of stacking vertically.
            bool showNums = navBarSettings.ShowSlideNumbers
                         || navBarSettings.SlideShapeType == NavBarSettings.ShapeType.NumberOnly;
            int maxSlideIdx = 0;
            if (showNums)
            {
                for (int si = 1; si <= sectionCount; si++)
                {
                    int last = sections.FirstSlide(si) + sections.SlidesCount(si) - 1;
                    if (last > maxSlideIdx) maxSlideIdx = last;
                }
            }
            int digitCount = Math.Max(1, maxSlideIdx.ToString().Length);
            float circleSize = showNums
                ? Math.Max(baseCircleSize, 6f + digitCount * 5f)
                : baseCircleSize;

            // Height of one band = pad + label + gap + circles + pad
            float bandHeight = bandPadTop + labelHeight + labelCircleGap + circleSize + bandPadBottom;

            float availableWidth = slideWidth - startX * 2f;

            // Pass 1 – column width for each section = max(name width, circles width)
            float[] nameWidths = new float[sectionCount + 1];
            float[] colWidths  = new float[sectionCount + 1];
            for (int i = 1; i <= sectionCount; i++)
            {
                nameWidths[i] = Math.Max(30f, sections.Name(i).Length * approxCharWidth);
                float circlesW = sections.SlidesCount(i) * (circleSize + circleSpacing);
                colWidths[i]  = Math.Max(nameWidths[i], circlesW);
            }

            // Pass 2 – greedily pack sections into bands (wrap when next column would overflow)
            var bands   = new List<List<int>>();
            var curBand = new List<int>();
            float curBandW = 0f;
            for (int i = 1; i <= sectionCount; i++)
            {
                float needed = colWidths[i] + (curBand.Count > 0 ? sectionGap : 0f);
                if (curBand.Count > 0 && curBandW + needed > availableWidth)
                {
                    bands.Add(curBand);
                    curBand = new List<int>();
                    curBandW = 0f;
                }
                curBandW += (curBand.Count > 0 ? sectionGap : 0f) + colWidths[i];
                curBand.Add(i);
            }
            if (curBand.Count > 0)
                bands.Add(curBand);

            // Bar height grows with the number of bands
            float barHeight = bands.Count * bandHeight + 6f; // 3pt top + 3pt bottom margin

            // Scale existing slide content to fit below the bar
            float contentScale = (slideHeight - barHeight) / slideHeight;
            for (int k = 1; k <= slide.Shapes.Count; k++)
            {
                try
                {
                    var sh = slide.Shapes[k];
                    if (sh.Tags["NavBar"] == "True") continue;
                    if (!string.IsNullOrEmpty(sh.Tags["NavBarOrigTop"])) continue;

                    sh.Tags.Add("NavBarOrigLeft",   sh.Left.ToString("R",   System.Globalization.CultureInfo.InvariantCulture));
                    sh.Tags.Add("NavBarOrigTop",    sh.Top.ToString("R",    System.Globalization.CultureInfo.InvariantCulture));
                    sh.Tags.Add("NavBarOrigWidth",  sh.Width.ToString("R",  System.Globalization.CultureInfo.InvariantCulture));
                    sh.Tags.Add("NavBarOrigHeight", sh.Height.ToString("R", System.Globalization.CultureInfo.InvariantCulture));

                    sh.Top    = barHeight + sh.Top * contentScale;
                    sh.Height = sh.Height * contentScale;
                }
                catch { }
            }

            // Background bar - USE CUSTOM COLOR
            PowerPoint.Shape navBackground = slide.Shapes.AddShape(
                Office.MsoAutoShapeType.msoShapeRectangle,
                0, 0, slideWidth, barHeight);
            navBackground.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(navBarSettings.BackgroundColor);
            navBackground.Line.Visible = Office.MsoTriState.msoFalse;
            navBackground.Tags.Add("NavBar", "True");

            int currentSlideIndex = slide.SlideIndex;

            // Pass 3 – draw each band: columns are  SECTION NAME
            //                                        ● ● ● ●
            for (int bandIdx = 0; bandIdx < bands.Count; bandIdx++)
            {
                float bandTop  = 3f + bandIdx * bandHeight;   // 3pt top bar margin
                float labelY   = bandTop + bandPadTop;
                float circleY  = bandTop + bandPadTop + labelHeight + labelCircleGap;
                float currentX = startX;

                foreach (int i in bands[bandIdx])
                {
                    string sectionName         = sections.Name(i);
                    int    slideCountInSection  = sections.SlidesCount(i);
                    int    firstSlideInSection  = sections.FirstSlide(i);
                    float  cw = colWidths[i];   // full column width

                    // Section name label – uses the full column width so it aligns with circles - USE CUSTOM COLOR
                    PowerPoint.Shape sectionLabel = slide.Shapes.AddTextbox(
                        Office.MsoTextOrientation.msoTextOrientationHorizontal,
                        currentX, labelY, cw, labelHeight);
                    sectionLabel.TextFrame.TextRange.Text = sectionName;
                    sectionLabel.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(navBarSettings.SectionNameColor);
                    sectionLabel.TextFrame.TextRange.Font.Size = labelFontSize;
                    sectionLabel.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                    sectionLabel.Line.Visible = Office.MsoTriState.msoFalse;
                    sectionLabel.Fill.Visible = Office.MsoTriState.msoFalse;
                    sectionLabel.TextFrame.WordWrap = Office.MsoTriState.msoFalse;
                    sectionLabel.Tags.Add("NavBar", "True");

                    // Circles are directly below the name, left-aligned with the column
                    float circleX = currentX;

                    // Subsection group boxes - USE CUSTOM COLORS
                    var subsectionGroups = GetSubsectionGroups(presentation, firstSlideInSection, slideCountInSection);
                    int colorIndex = 0;
                    foreach (var group in subsectionGroups)
                    {
                        if (!string.IsNullOrEmpty(group.SubsectionName))
                        {
                            float boxX      = circleX + group.StartIndex * (circleSize + circleSpacing) - 2f;
                            float boxY      = circleY - 2f;
                            float boxWidth  = group.Count * (circleSize + circleSpacing) - circleSpacing + 4f;
                            float boxHeight = circleSize + 4f;

                            PowerPoint.Shape subsectionBox = slide.Shapes.AddShape(
                                Office.MsoAutoShapeType.msoShapeRoundedRectangle,
                                boxX, boxY, boxWidth, boxHeight);

                            System.Drawing.Color fillColor = navBarSettings.SubsectionBoxColors[colorIndex % navBarSettings.SubsectionBoxColors.Length];
                            subsectionBox.Fill.Visible        = Office.MsoTriState.msoTrue;
                            subsectionBox.Fill.ForeColor.RGB  = System.Drawing.ColorTranslator.ToOle(fillColor);
                            subsectionBox.Fill.Transparency   = navBarSettings.SubsectionBoxTransparency;
                            subsectionBox.Line.Visible        = Office.MsoTriState.msoTrue;
                            subsectionBox.Line.ForeColor.RGB  = System.Drawing.ColorTranslator.ToOle(fillColor);
                            subsectionBox.Line.Weight         = 2.0f;
                            subsectionBox.Line.Transparency   = 0.0f;
                            try { subsectionBox.Adjustments[1] = 0.25f; } catch { }
                            subsectionBox.Tags.Add("NavBar", "True");
                            subsectionBox.Tags.Add("SubsectionBox", group.SubsectionName);
                            subsectionBox.ZOrder(Office.MsoZOrderCmd.msoSendToBack);

                            System.Diagnostics.Debug.WriteLine($"Created colored box for subsection '{group.SubsectionName}' at position {boxX}, color index {colorIndex}");
                            colorIndex++;
                        }
                    }

                    // Draw slide circles
                    string currentSlideSubsection = "";
                    try { currentSlideSubsection = presentation.Slides[currentSlideIndex].Tags["Subsection"]; } catch { }

                    float cx = circleX;
                    for (int j = 0; j < slideCountInSection; j++)
                    {
                        int slideIndexInPresentation = firstSlideInSection + j;
                        PowerPoint.Shape shape = null;

                        if (navBarSettings.SlideShapeType == NavBarSettings.ShapeType.NumberOnly)
                        {
                            shape = slide.Shapes.AddTextbox(
                                Office.MsoTextOrientation.msoTextOrientationHorizontal,
                                cx, circleY, circleSize, circleSize);
                            shape.Line.Visible = Office.MsoTriState.msoFalse;
                            shape.Fill.Visible = Office.MsoTriState.msoFalse;
                        }
                        else if (navBarSettings.SlideShapeType == NavBarSettings.ShapeType.Square)
                        {
                            shape = slide.Shapes.AddShape(
                                Office.MsoAutoShapeType.msoShapeRectangle,
                                cx, circleY, circleSize, circleSize);
                        }
                        else
                        {
                            shape = slide.Shapes.AddShape(
                                Office.MsoAutoShapeType.msoShapeOval,
                                cx, circleY, circleSize, circleSize);
                        }

                        string thisSlideSubsection = "";
                        try { thisSlideSubsection = presentation.Slides[slideIndexInPresentation].Tags["Subsection"]; } catch { }

                        if (navBarSettings.SlideShapeType != NavBarSettings.ShapeType.NumberOnly)
                        {
                            if (slideIndexInPresentation == currentSlideIndex)
                            {
                                // CURRENT SLIDE - USE CUSTOM COLOR
                                shape.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(navBarSettings.CurrentSlideColor);
                                shape.Line.Visible = Office.MsoTriState.msoFalse;
                            }
                            else if (!string.IsNullOrEmpty(currentSlideSubsection) &&
                                     !string.IsNullOrEmpty(thisSlideSubsection) &&
                                     thisSlideSubsection == currentSlideSubsection)
                            {
                                // SAME SUBSECTION - USE CUSTOM COLORS
                                shape.Fill.ForeColor.RGB  = System.Drawing.ColorTranslator.ToOle(navBarSettings.SameSubsectionFillColor);
                                shape.Fill.Visible        = Office.MsoTriState.msoTrue;
                                shape.Line.Visible        = Office.MsoTriState.msoTrue;
                                shape.Line.ForeColor.RGB  = System.Drawing.ColorTranslator.ToOle(navBarSettings.SameSubsectionBorderColor);
                                shape.Line.Weight         = 2.0f;
                            }
                            else
                            {
                                // OTHER SLIDES - USE CUSTOM COLOR
                                shape.Fill.Visible        = Office.MsoTriState.msoFalse;
                                shape.Line.Visible        = Office.MsoTriState.msoTrue;
                                shape.Line.ForeColor.RGB  = System.Drawing.ColorTranslator.ToOle(navBarSettings.OtherSlidesBorderColor);
                                shape.Line.Weight         = 1.5f;
                            }
                        }

                        shape.Tags.Add("NavBar", "True");
                        shape.Tags.Add("NavBarSlide", slideIndexInPresentation.ToString());

                        // ADD SLIDE NUMBERS if enabled OR if NumberOnly mode
                        if (navBarSettings.ShowSlideNumbers || navBarSettings.SlideShapeType == NavBarSettings.ShapeType.NumberOnly)
                        {
                            shape.TextFrame.TextRange.Text = slideIndexInPresentation.ToString();
                            shape.TextFrame.TextRange.Font.Size = navBarSettings.SlideShapeType == NavBarSettings.ShapeType.NumberOnly ? 10 : 8;
                            shape.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(navBarSettings.SlideNumberColor);
                            shape.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                            shape.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;
                            shape.TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;
                            shape.TextFrame.MarginLeft   = 0;
                            shape.TextFrame.MarginRight  = 0;
                            shape.TextFrame.MarginTop    = 0;
                            shape.TextFrame.MarginBottom = 0;
                        }

                        cx += circleSize + circleSpacing;
                    }

                    currentX += cw + sectionGap;
                }
            }
        }

        private class SubsectionGroup
        {
            public string SubsectionName { get; set; }
            public int StartIndex { get; set; }
            public int Count { get; set; }
        }

        private List<SubsectionGroup> GetSubsectionGroups(PowerPoint.Presentation presentation, int firstSlide, int slideCount)
        {
            List<SubsectionGroup> groups = new List<SubsectionGroup>();

            string currentSubsection = null;
            int groupStart = 0;
            int groupCount = 0;

            string debugInfo = $"DEBUG: Checking section starting at slide {firstSlide}, {slideCount} slides total\n";

            for (int i = 0; i < slideCount; i++)
            {
                try
                {
                    PowerPoint.Slide slide = presentation.Slides[firstSlide + i];
                    string subsectionTag = slide.Tags["Subsection"];

                    debugInfo += $"Slide {firstSlide + i}: Subsection = '{subsectionTag}'\n";

                    if (!string.IsNullOrEmpty(subsectionTag))
                    {
                        if (subsectionTag == currentSubsection)
                        {
                            // Continue current group
                            groupCount++;
                        }
                        else
                        {
                            // Save previous group if exists
                            if (currentSubsection != null && groupCount > 0)
                            {
                                var newGroup = new SubsectionGroup
                                {
                                    SubsectionName = currentSubsection,
                                    StartIndex = groupStart,
                                    Count = groupCount
                                };
                                groups.Add(newGroup);
                                debugInfo += $"  -> Added group: {currentSubsection}, Start={groupStart}, Count={groupCount}\n";
                            }

                            // Start new group
                            currentSubsection = subsectionTag;
                            groupStart = i;
                            groupCount = 1;
                        }
                    }
                    else
                    {
                        // Save previous group if exists
                        if (currentSubsection != null && groupCount > 0)
                        {
                            var newGroup = new SubsectionGroup
                            {
                                SubsectionName = currentSubsection,
                                StartIndex = groupStart,
                                Count = groupCount
                            };
                            groups.Add(newGroup);
                            debugInfo += $"  -> Added group: {currentSubsection}, Start={groupStart}, Count={groupCount}\n";
                        }
                        currentSubsection = null;
                        groupCount = 0;
                    }
                }
                catch (Exception ex)
                {
                    debugInfo += $"  ERROR on slide {firstSlide + i}: {ex.Message}\n";
                    continue;
                }
            }

            // Add last group if exists
            if (currentSubsection != null && groupCount > 0)
            {
                var newGroup = new SubsectionGroup
                {
                    SubsectionName = currentSubsection,
                    StartIndex = groupStart,
                    Count = groupCount
                };
                groups.Add(newGroup);
                debugInfo += $"  -> Added LAST group: {currentSubsection}, Start={groupStart}, Count={groupCount}\n";
            }

            debugInfo += $"\nTotal groups found: {groups.Count}\n";

            // Show debug info
            if (groups.Count > 0)
            {
                System.Diagnostics.Debug.WriteLine(debugInfo);
                // Uncomment to see in MessageBox:
                // MessageBox.Show(debugInfo, "Subsection Debug", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            return groups;
        }

        private void editBox1_TextChanged_1(object sender, RibbonControlEventArgs e)
        {

        }

        private void button2_Click_1(object sender, RibbonControlEventArgs e)
        {
            CreateSectionNavigationBar();
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            CreateSectionNavigationBar();
        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            RemoveAllNavigationBars();
        }

        private void RemoveAllNavigationBars()
        {
            try
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;

                if (app.ActivePresentation == null)
                {
                    MessageBox.Show("Please open a presentation first.", "No Presentation",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                PowerPoint.Presentation presentation = app.ActivePresentation;
                int slidesProcessed = 0;

                // Loop through all slides and remove navigation bars
                foreach (PowerPoint.Slide slide in presentation.Slides)
                {
                    RemoveNavigationBar(slide);
                    slidesProcessed++;
                }

                MessageBox.Show($"Navigation bar removed from {slidesProcessed} slide(s)!",
                    "Success",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error removing navigation bar: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void editBox2_TextChanged(object sender, RibbonControlEventArgs e)
        {

        }



        private void btnNavBarSettings_Click(object sender, RibbonControlEventArgs e)
        {
            using (var settingsDialog = new NavBarColorSettings(navBarSettings))
            {
                if (settingsDialog.ShowDialog() == DialogResult.OK)
                {
                    navBarSettings = settingsDialog.Settings;
                    MessageBox.Show("Colors updated! Click 'Refresh Nav Bar' to apply changes.",
                        "Settings Saved", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }







        private void effectsplitbtn_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void linkbtn_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.Presentations.Count == 0)
                {
                    MessageBox.Show("Please open a presentation first.", "No Presentation",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var pres = app.ActivePresentation;
                int linksApplied = 0;
                int slidesVisited = 0;

                foreach (PowerPoint.Slide slide in pres.Slides)
                {
                    slidesVisited++;
                    for (int i = 1; i <= slide.Shapes.Count; i++)
                    {
                        try
                        {
                            var shape = slide.Shapes[i];
                            if (shape.Tags["NavBar"] != "True") continue;

                            string navSlideTag = shape.Tags["NavBarSlide"];
                            if (string.IsNullOrEmpty(navSlideTag)) continue;
                            if (!int.TryParse(navSlideTag, out int targetSlideNum)) continue;
                            if (targetSlideNum < 1 || targetSlideNum > pres.Slides.Count) continue;

                            var targetSlide = pres.Slides[targetSlideNum];
                            string subAddress = $"{targetSlide.SlideID},{targetSlide.SlideIndex},{targetSlide.Name}";

                            // Apply click-action hyperlink to the target slide
                            var click = shape.ActionSettings[PowerPoint.PpMouseActivation.ppMouseClick];
                            click.Action = PowerPoint.PpActionType.ppActionNone;
                            click.Action = PowerPoint.PpActionType.ppActionHyperlink;
                            click.Hyperlink.Address = "";
                            click.Hyperlink.SubAddress = subAddress;

                            // Clear mouse-over so it cannot interfere
                            var over = shape.ActionSettings[PowerPoint.PpMouseActivation.ppMouseOver];
                            over.Action = PowerPoint.PpActionType.ppActionNone;

                            linksApplied++;
                        }
                        catch { }
                    }
                }

                if (linksApplied > 0)
                    MessageBox.Show(
                        $"{linksApplied} hyperlink(s) added across {slidesVisited} slide(s).\n\n" +
                        "\u26a0 Links only fire during Slide Show (F5) \u2014 not in Normal Edit view.",
                        "Links Added", MessageBoxButtons.OK, MessageBoxIcon.Information);
                else
                    MessageBox.Show(
                        "No navigation bar circles found.\n\n" +
                        "Please generate (or refresh) the navigation bar first, then click 'Add Links'.",
                        "No Nav Bar Found", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Add Links error: " + ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void selectEffecctdtn_Click(object sender, RibbonControlEventArgs e)
        {
            try { SpotlightFeature.SelectAreas(Globals.ThisAddIn.Application); }
            catch (Exception ex)
            {
                MessageBox.Show("Select Effect Areas error: " + ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void spotlightSettingsBtn_Click(object sender, RibbonControlEventArgs e)
        {
            try { SpotlightFeature.ShowSettings(); }
            catch (Exception ex)
            {
                MessageBox.Show("Settings error: " + ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void createspotlightbtn_Click(object sender, RibbonControlEventArgs e)
        {
            try { SpotlightFeature.CreateSpotlight(Globals.ThisAddIn.Application); }
            catch (Exception ex)
            {
                MessageBox.Show("Create Spotlight error: " + ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void selectzoombtn_Click(object sender, RibbonControlEventArgs e)
        {
            try { ZoomFeature.SelectZoom(Globals.ThisAddIn.Application); }
            catch (Exception ex)
            {
                MessageBox.Show("Select Zoom error: " + ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void zoomaddbtn_Click(object sender, RibbonControlEventArgs e)
        {
            try { ZoomFeature.ZoomAdd(Globals.ThisAddIn.Application); }
            catch (Exception ex)
            {
                MessageBox.Show("Zoom Add error: " + ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void createzoombtn_Click(object sender, RibbonControlEventArgs e)
        {
            try { ZoomFeature.Create(Globals.ThisAddIn.Application); }
            catch (Exception ex)
            {
                MessageBox.Show("Create error: " + ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void magniaddbtn_Click(object sender, RibbonControlEventArgs e)
        {
            try { MagnifyFeature.AddMagnify(Globals.ThisAddIn.Application); }
            catch (Exception ex)
            {
                MessageBox.Show("Magnify error: " + ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void magsetingsbtn_Click(object sender, RibbonControlEventArgs e)
        {
            try { MagnifyFeature.ShowSettings(); }
            catch (Exception ex)
            {
                MessageBox.Show("Settings error: " + ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SubSectionEnd_TextChanged(object sender, RibbonControlEventArgs e)
        {

        }

        private void adjustbtn_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.Presentations.Count == 0)
                {
                    MessageBox.Show("Please open a presentation first.", "No Presentation",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var pres = app.ActivePresentation;
                float slideHeight = pres.PageSetup.SlideHeight;
                int shapesFixed = 0;
                int slidesFixed = 0;

                foreach (PowerPoint.Slide slide in pres.Slides)
                {
                    // Find the nav bar background on this slide to read its height
                    float barHeight = 0f;
                    for (int i = 1; i <= slide.Shapes.Count; i++)
                    {
                        try
                        {
                            var sh = slide.Shapes[i];
                            if (sh.Tags["NavBar"] == "True" && sh.Top < 1f && sh.Left < 1f)
                            {
                                barHeight = sh.Height;
                                break;
                            }
                        }
                        catch { }
                    }

                    if (barHeight <= 0f) continue; // no nav bar on this slide

                    float contentScale = (slideHeight - barHeight) / slideHeight;
                    bool slideChanged = false;

                    for (int k = 1; k <= slide.Shapes.Count; k++)
                    {
                        try
                        {
                            var sh = slide.Shapes[k];
                            if (sh.Tags["NavBar"] == "True") continue;           // skip nav bar shapes
                            if (!string.IsNullOrEmpty(sh.Tags["NavBarOrigTop"])) continue; // already adjusted

                            // Save original geometry so RemoveNavigationBar can restore it later
                            sh.Tags.Add("NavBarOrigLeft", sh.Left.ToString("R", System.Globalization.CultureInfo.InvariantCulture));
                            sh.Tags.Add("NavBarOrigTop", sh.Top.ToString("R", System.Globalization.CultureInfo.InvariantCulture));
                            sh.Tags.Add("NavBarOrigWidth", sh.Width.ToString("R", System.Globalization.CultureInfo.InvariantCulture));
                            sh.Tags.Add("NavBarOrigHeight", sh.Height.ToString("R", System.Globalization.CultureInfo.InvariantCulture));

                            // Shift down and compress vertically into the content area
                            sh.Top = barHeight + sh.Top * contentScale;
                            sh.Height = sh.Height * contentScale;

                            shapesFixed++;
                            slideChanged = true;
                        }
                        catch { }
                    }

                    if (slideChanged) slidesFixed++;
                }

                if (shapesFixed > 0)
                    MessageBox.Show(
                        $"Adjusted {shapesFixed} shape(s) across {slidesFixed} slide(s) to fit below the navigation bar.",
                        "Adjust Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                else
                    MessageBox.Show(
                        "No new shapes found to adjust.\n\n" +
                        "All content is already positioned below the navigation bar,\n" +
                        "or no navigation bar exists yet.",
                        "Nothing to Adjust", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Adjust error: " + ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void focusbtn_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void SlideNobox_TextChanged(object sender, RibbonControlEventArgs e)
        {

        }

        private PowerPoint.Shape FindShapeOnSlide(PowerPoint.Slide slide, string name)
        {
            for (int i = 1; i <= slide.Shapes.Count; i++)
                if (string.Equals(slide.Shapes[i].Name, name, StringComparison.OrdinalIgnoreCase))
                    return slide.Shapes[i];
            return null;
        }

        private void crtHypBtn_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.Presentations.Count == 0) return;

                var pres = app.ActivePresentation;

                string slideText = SlideNobox.Text.Trim();
                if (string.IsNullOrEmpty(slideText) || !int.TryParse(slideText, out int targetSlideNum))
                {
                    MessageBox.Show("Please enter a valid slide number in the 'Slide No' box.",
                        "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (targetSlideNum < 1 || targetSlideNum > pres.Slides.Count)
                {
                    MessageBox.Show(
                        $"Slide {targetSlideNum} does not exist. " +
                        $"The presentation has {pres.Slides.Count} slide(s).",
                        "Invalid Slide", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Use stored shapes; fall back to current selection if nothing stored
                List<string> shapeNames = new List<string>(_hypShapeNames);
                int srcSlideIdx = _hypSourceSlideIndex;

                if (shapeNames.Count == 0)
                {
                    var win = app.ActiveWindow;
                    if (win.Selection != null &&
                        (win.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes ||
                         win.Selection.Type == PowerPoint.PpSelectionType.ppSelectionText))
                    {
                        srcSlideIdx = (win.View.Slide as PowerPoint.Slide).SlideIndex;
                        var sr = win.Selection.ShapeRange;
                        for (int i = 1; i <= sr.Count; i++)
                            shapeNames.Add(sr[i].Name);
                    }
                    else
                    {
                        MessageBox.Show(
                            "No object stored or selected.\n\n" +
                            "Select an object then click 'Select Text or Area' first.",
                            "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                }

                if (srcSlideIdx < 1 || srcSlideIdx > pres.Slides.Count)
                {
                    MessageBox.Show("Source slide no longer exists.", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                var srcSlide = pres.Slides[srcSlideIdx];
                var targetSlide = pres.Slides[targetSlideNum];

                // SlideID,SlideIndex,SlideName is the full internal format PowerPoint uses
                // for "Place in This Document" hyperlinks — most robust across versions.
                string subAddress = $"{targetSlide.SlideID},{targetSlide.SlideIndex},{targetSlide.Name}";

                int applied = 0;
                string diagnosticInfo = "";
                foreach (string shapeName in shapeNames)
                {
                    var sh = FindShapeOnSlide(srcSlide, shapeName);
                    if (sh == null) continue;

                    // Reset click action so the Hyperlink object is clean
                    var click = sh.ActionSettings[PowerPoint.PpMouseActivation.ppMouseClick];
                    click.Action = PowerPoint.PpActionType.ppActionNone;
                    click.Action = PowerPoint.PpActionType.ppActionHyperlink;
                    click.Hyperlink.Address = "";
                    click.Hyperlink.SubAddress = subAddress;

                    // Also clear mouse-over so it cannot interfere
                    var over = sh.ActionSettings[PowerPoint.PpMouseActivation.ppMouseOver];
                    over.Action = PowerPoint.PpActionType.ppActionNone;

                    try
                    {
                        diagnosticInfo = $"Address='{click.Hyperlink.Address}'  SubAddress='{click.Hyperlink.SubAddress}'";
                    }
                    catch { }
                    applied++;
                }

                if (applied > 0)
                    MessageBox.Show(
                        $"Hyperlink to Slide {targetSlideNum} ('{subAddress}') applied to {applied} object(s).\n\n" +
                        $"Stored as: {diagnosticInfo}\n\n" +
                        "⚠ The link only fires during Slide Show (F5) — not in Normal Edit view.",
                        "Hyperlink Created", MessageBoxButtons.OK, MessageBoxIcon.Information);
                else
                    MessageBox.Show(
                        "Could not find the stored shapes on the slide.\n" +
                        "Please re-select the object and try again.",
                        "Shapes Not Found", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Create Hyperlink error: " + ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void rmvHypbtn_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.Presentations.Count == 0) return;

                var pres = app.ActivePresentation;

                // Prefer current selection; fall back to stored shapes
                List<string> shapeNames = new List<string>();
                int srcSlideIdx = -1;

                var win = app.ActiveWindow;
                if (win.Selection != null &&
                    (win.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes ||
                     win.Selection.Type == PowerPoint.PpSelectionType.ppSelectionText))
                {
                    srcSlideIdx = (win.View.Slide as PowerPoint.Slide).SlideIndex;
                    var sr = win.Selection.ShapeRange;
                    for (int i = 1; i <= sr.Count; i++)
                        shapeNames.Add(sr[i].Name);
                }
                else if (_hypShapeNames.Count > 0)
                {
                    shapeNames = new List<string>(_hypShapeNames);
                    srcSlideIdx = _hypSourceSlideIndex;
                }
                else
                {
                    MessageBox.Show(
                        "No object selected or stored.\n\n" +
                        "Select an object (or click 'Select Text or Area' first) then try again.",
                        "No Selection", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (srcSlideIdx < 1 || srcSlideIdx > pres.Slides.Count) return;

                var srcSlide = pres.Slides[srcSlideIdx];
                int removed = 0;

                foreach (string shapeName in shapeNames)
                {
                    var sh = FindShapeOnSlide(srcSlide, shapeName);
                    if (sh == null) continue;

                    var action = sh.ActionSettings[PowerPoint.PpMouseActivation.ppMouseClick];
                    if (action.Action != PowerPoint.PpActionType.ppActionNone)
                    {
                        action.Action = PowerPoint.PpActionType.ppActionNone;
                        removed++;
                    }
                }

                MessageBox.Show(
                    removed > 0
                        ? $"Hyperlink removed from {removed} object(s)."
                        : "No hyperlink found on the selected object(s).",
                    "Remove Hyperlink", MessageBoxButtons.OK,
                    removed > 0 ? MessageBoxIcon.Information : MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Remove Hyperlink error: " + ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void selecthypbtn_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.Presentations.Count == 0)
                {
                    MessageBox.Show("Please open a presentation first.", "No Presentation",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var win = app.ActiveWindow;
                if (win.Selection == null ||
                    (win.Selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes &&
                     win.Selection.Type != PowerPoint.PpSelectionType.ppSelectionText))
                {
                    MessageBox.Show(
                        "Please select a shape, image, text box or object on the slide first.",
                        "Nothing Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                _hypShapeNames.Clear();
                _hypSourceSlideIndex = (win.View.Slide as PowerPoint.Slide).SlideIndex;

                var sr = win.Selection.ShapeRange;
                for (int i = 1; i <= sr.Count; i++)
                    _hypShapeNames.Add(sr[i].Name);

                MessageBox.Show(
                    $"{_hypShapeNames.Count} object(s) stored from Slide {_hypSourceSlideIndex}.\n\n" +
                    "• Enter a slide number in the box.\n" +
                    "• Click 'Create Hyperlink' to apply the link.",
                    "Selection Stored", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Select error: " + ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void blursettingbtn_Click(object sender, RibbonControlEventArgs e)
        {
            try { BlurFeature.ShowSettings(); }
            catch (Exception ex)
            {
                MessageBox.Show("Blur Settings error: " + ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void blur_allexceptbtn_Click(object sender, RibbonControlEventArgs e)
        {
            try { BlurFeature.BlurAllExcept(Globals.ThisAddIn.Application); }
            catch (Exception ex)
            {
                MessageBox.Show("Blur All Except error: " + ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void blur_selectbtn_Click(object sender, RibbonControlEventArgs e)
        {
            try { BlurFeature.BlurSelected(Globals.ThisAddIn.Application); }
            catch (Exception ex)
            {
                MessageBox.Show("Blur Selected error: " + ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void blur_remainbtn_Click(object sender, RibbonControlEventArgs e)
        {
            try { BlurFeature.BlurRemainder(Globals.ThisAddIn.Application); }
            catch (Exception ex)
            {
                MessageBox.Show("Blur Remainder error: " + ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnToggleQR_Click(object sender, RibbonControlEventArgs e)
        {
            var taskPane = Globals.ThisAddIn.QRTaskPane;
            if (taskPane != null)
            {
                taskPane.Visible = btnToggleQR.Checked;
            }
        }

        public void SyncQRToggleButton(bool visible)
        {
            btnToggleQR.Checked = visible;
        }

        private void positionsLabBtn_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.TogglePositionsLabPane();
        }

        private void resizeBtn_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.ToggleResizeLabPane();
        }

        private void addAgendaBtn_Click(object sender, RibbonControlEventArgs e)
        {
            try { AgendaGenerator.AddAgenda(Globals.ThisAddIn.Application); }
            catch (Exception ex)
            {
                MessageBox.Show("Add Agenda error: " + ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void removeAgendaBtn_Click(object sender, RibbonControlEventArgs e)
        {
            try { AgendaGenerator.RemoveAgenda(Globals.ThisAddIn.Application); }
            catch (Exception ex)
            {
                MessageBox.Show("Remove Agenda error: " + ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void refreshAgendaBtn_Click(object sender, RibbonControlEventArgs e)
        {
            try { AgendaGenerator.RefreshAgenda(Globals.ThisAddIn.Application); }
            catch (Exception ex)
            {
                MessageBox.Show("Refresh Agenda error: " + ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
