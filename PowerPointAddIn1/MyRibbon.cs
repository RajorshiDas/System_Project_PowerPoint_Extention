using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PowerPointAddIn1
{
    public partial class MyRibbon
    {
        // Navigation Bar Customization Settings
        private NavBarSettings navBarSettings = new NavBarSettings();

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
            // Remove shapes tagged as navigation bar
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

        private void AddNavigationBarToSlide(PowerPoint.Slide slide, PowerPoint.SectionProperties sections, PowerPoint.Presentation presentation)
        {
            float barHeight = 60;
            float slideWidth = presentation.PageSetup.SlideWidth;
            
            // Create background bar - USE CUSTOM COLOR
            PowerPoint.Shape navBackground = slide.Shapes.AddShape(
                Office.MsoAutoShapeType.msoShapeRectangle,
                0, 0, slideWidth, barHeight);
            navBackground.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(navBarSettings.BackgroundColor);
            navBackground.Line.Visible = Office.MsoTriState.msoFalse;
            navBackground.Tags.Add("NavBar", "True");

            float startX = 20;
            float currentX = startX;
            float topY = 8;
            float circleSize = 12;
            float circleSpacing = 6;

            int currentSlideIndex = slide.SlideIndex;
            int currentSectionIndex = GetSectionIndexForSlide(sections, currentSlideIndex);

            for (int i = 1; i <= sections.Count; i++)
            {
                string sectionName = sections.Name(i);
                int slideCountInSection = sections.SlidesCount(i);
                int firstSlideInSection = sections.FirstSlide(i);

                // Add section name - USE CUSTOM COLOR
                PowerPoint.Shape sectionLabel = slide.Shapes.AddTextbox(
                    Office.MsoTextOrientation.msoTextOrientationHorizontal,
                    currentX, topY, 200, 20);
                sectionLabel.TextFrame.TextRange.Text = sectionName;
                sectionLabel.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(navBarSettings.SectionNameColor);
                sectionLabel.TextFrame.TextRange.Font.Size = 12;
                sectionLabel.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                sectionLabel.Line.Visible = Office.MsoTriState.msoFalse;
                sectionLabel.Fill.Visible = Office.MsoTriState.msoFalse;
                sectionLabel.Tags.Add("NavBar", "True");

                // Add circles for slides with subsection grouping
                float circleY = topY + 25;
                float circleX = currentX;

                // Get subsections for this section
                var subsectionGroups = GetSubsectionGroups(presentation, firstSlideInSection, slideCountInSection);

                // Draw colored backgrounds for subsection groups
                int colorIndex = 0;
                foreach (var group in subsectionGroups)
                {
                    if (!string.IsNullOrEmpty(group.SubsectionName))
                    {
                        // Calculate box position and size
                        float boxX = currentX + (group.StartIndex * (circleSize + circleSpacing)) - 2;
                        float boxY = circleY - 2;
                        float boxWidth = (group.Count * (circleSize + circleSpacing)) - circleSpacing + 4;
                        float boxHeight = circleSize + 4;

                        // Draw rounded rectangle with colored fill - USE CUSTOM COLORS
                        PowerPoint.Shape subsectionBox = slide.Shapes.AddShape(
                            Office.MsoAutoShapeType.msoShapeRoundedRectangle,
                            boxX, boxY, boxWidth, boxHeight);
                        
                        // Set fill color (cycling through colors)
                        System.Drawing.Color fillColor = navBarSettings.SubsectionBoxColors[colorIndex % navBarSettings.SubsectionBoxColors.Length];
                        subsectionBox.Fill.Visible = Office.MsoTriState.msoTrue;
                        subsectionBox.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(fillColor);
                        subsectionBox.Fill.Transparency = navBarSettings.SubsectionBoxTransparency;
                        
                        // Add thick visible border
                        subsectionBox.Line.Visible = Office.MsoTriState.msoTrue;
                        subsectionBox.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(fillColor);
                        subsectionBox.Line.Weight = 2.0f;
                        subsectionBox.Line.Transparency = 0.0f;
                        
                        try
                        {
                            subsectionBox.Adjustments[1] = 0.25f;
                        }
                        catch { }
                        
                        subsectionBox.Tags.Add("NavBar", "True");
                        subsectionBox.Tags.Add("SubsectionBox", group.SubsectionName);
                        subsectionBox.ZOrder(Office.MsoZOrderCmd.msoSendToBack);
                        
                        System.Diagnostics.Debug.WriteLine($"Created colored box for subsection '{group.SubsectionName}' at position {boxX}, color index {colorIndex}");
                        
                        colorIndex++;
                    }
                }

                // Draw circles
                string currentSlideSubsection = "";
                try
                {
                    PowerPoint.Slide currentSlide = presentation.Slides[currentSlideIndex];
                    currentSlideSubsection = currentSlide.Tags["Subsection"];
                }
                catch { }

                for (int j = 0; j < slideCountInSection; j++)
                {
                    int slideIndexInPresentation = firstSlideInSection + j;
                    
                    PowerPoint.Shape shape = null;
                    
                    // Check if we're using Number Only mode
                    if (navBarSettings.SlideShapeType == NavBarSettings.ShapeType.NumberOnly)
                    {
                        // Create just a text box for number only
                        shape = slide.Shapes.AddTextbox(
                            Office.MsoTextOrientation.msoTextOrientationHorizontal,
                            circleX, circleY, circleSize, circleSize);
                        shape.Line.Visible = Office.MsoTriState.msoFalse;
                        shape.Fill.Visible = Office.MsoTriState.msoFalse;
                    }
                    else
                    {
                        // Create shape based on settings (Circle or Square)
                        if (navBarSettings.SlideShapeType == NavBarSettings.ShapeType.Square)
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
                    }
                    
                    // Get this slide's subsection
                    string thisSlideSubsection = "";
                    try
                    {
                        PowerPoint.Slide thisSlide = presentation.Slides[slideIndexInPresentation];
                        thisSlideSubsection = thisSlide.Tags["Subsection"];
                    }
                    catch { }

                    // Apply colors only if NOT Number Only mode
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
                            shape.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(navBarSettings.SameSubsectionFillColor);
                            shape.Fill.Visible = Office.MsoTriState.msoTrue;
                            shape.Line.Visible = Office.MsoTriState.msoTrue;
                            shape.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(navBarSettings.SameSubsectionBorderColor);
                            shape.Line.Weight = 2.0f;
                        }
                        else
                        {
                            // OTHER SLIDES - USE CUSTOM COLOR
                            shape.Fill.Visible = Office.MsoTriState.msoFalse;
                            shape.Line.Visible = Office.MsoTriState.msoTrue;
                            shape.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(navBarSettings.OtherSlidesBorderColor);
                            shape.Line.Weight = 1.5f;
                        }
                    }
                    
                    shape.Tags.Add("NavBar", "True");

                    // ADD SLIDE NUMBERS if enabled OR if NumberOnly mode
                    if (navBarSettings.ShowSlideNumbers || navBarSettings.SlideShapeType == NavBarSettings.ShapeType.NumberOnly)
                    {
                        shape.TextFrame.TextRange.Text = slideIndexInPresentation.ToString();
                        shape.TextFrame.TextRange.Font.Size = navBarSettings.SlideShapeType == NavBarSettings.ShapeType.NumberOnly ? 10 : 8;
                        shape.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(navBarSettings.SlideNumberColor);
                        shape.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                        shape.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;
                        shape.TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;
                        shape.TextFrame.MarginLeft = 0;
                        shape.TextFrame.MarginRight = 0;
                        shape.TextFrame.MarginTop = 0;
                        shape.TextFrame.MarginBottom = 0;
                    }
                    
                    circleX += circleSize + circleSpacing;
                }

                currentX += Math.Max(150, (circleSize + circleSpacing) * slideCountInSection) + 30;
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
    }
}
        