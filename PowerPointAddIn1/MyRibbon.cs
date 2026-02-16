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
        private void MyRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            UpdateSectionInfo();
        }

        public void RefreshInfo()
        {
            UpdateSectionInfo();
        }

        private void btnAddSlide_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // Get the active PowerPoint application
                PowerPoint.Application app = Globals.ThisAddIn.Application;
                
                // Get the active presentation
                PowerPoint.Presentation presentation = app.ActivePresentation;
                
                if (presentation != null)
                {
                    // Get the slides collection
                    PowerPoint.Slides slides = presentation.Slides;
                    
                    // Add a new slide with Title and Content layout (layout index 2)
                    // Index = slides.Count + 1 to add at the end
                    PowerPoint.Slide newSlide = slides.Add(
                        slides.Count + 1, 
                        PowerPoint.PpSlideLayout.ppLayoutText);
                    
                    // Add title text
                    PowerPoint.Shape titleShape = newSlide.Shapes.Title;
                    titleShape.TextFrame.TextRange.Text = "New Slide Title";
                    
                    // Add body text to the content placeholder
                    // The second shape (index 2) is typically the content placeholder
                    if (newSlide.Shapes.Count > 1)
                    {
                        PowerPoint.Shape bodyShape = newSlide.Shapes[2];
                        bodyShape.TextFrame.TextRange.Text = "This is the body text of the new slide.\n\n" +
                            "• Bullet point 1\n" +
                            "• Bullet point 2\n" +
                            "• Bullet point 3";
                    }
                    
                    // Optional: Make the new slide the active slide
                    newSlide.Select();
                    
                    MessageBox.Show(
                        "New slide added successfully!", 
                        "Add Slide", 
                        MessageBoxButtons.OK, 
                        MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show(
                        "Please open a presentation first.", 
                        "No Presentation", 
                        MessageBoxButtons.OK, 
                        MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Error adding slide: " + ex.Message, 
                    "Error", 
                    MessageBoxButtons.OK, 
                    MessageBoxIcon.Error);
            }
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;
                PowerPoint.Selection selection = app.ActiveWindow.Selection;

                if (selection.Type == PowerPoint.PpSelectionType.ppSelectionText)
                {
                    PowerPoint.TextRange textRange = selection.TextRange;
                    textRange.InsertAfter("Hello World");
                }
                else
                {
                    MessageBox.Show(
                        "Please place your cursor in a text box first.",
                        "No Text Selected",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Error inserting text: " + ex.Message,
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void editBox1_TextChanged(object sender, RibbonControlEventArgs e)
        {

        }

        private void UpdateSectionInfo()
        {
            try
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;
                
                if (app.ActivePresentation == null)
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
            
            // Create black background bar
            PowerPoint.Shape navBackground = slide.Shapes.AddShape(
                Office.MsoAutoShapeType.msoShapeRectangle,
                0, 0, slideWidth, barHeight);
            navBackground.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            navBackground.Line.Visible = Office.MsoTriState.msoFalse;
            navBackground.Tags.Add("NavBar", "True");

            float startX = 20;
            float currentX = startX;
            float topY = 8;
            float circleSize = 12;
            float circleSpacing = 6;

            int currentSlideIndex = slide.SlideIndex;
            int currentSectionIndex = GetSectionIndexForSlide(sections, currentSlideIndex);

            // Define colors for subsections (more visible, solid colors)
            System.Drawing.Color[] subsectionColors = new System.Drawing.Color[]
            {
                System.Drawing.Color.SteelBlue,        // Steel Blue
                System.Drawing.Color.MediumSeaGreen,   // Medium Sea Green
                System.Drawing.Color.Goldenrod,        // Golden Rod
                System.Drawing.Color.IndianRed,        // Indian Red
                System.Drawing.Color.BlueViolet,       // Blue Violet
                System.Drawing.Color.DarkOrange,       // Dark Orange
                System.Drawing.Color.DarkSlateBlue,    // Dark Slate Blue
                System.Drawing.Color.RosyBrown         // Rosy Brown
            };

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
                sectionLabel.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
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

                        // Draw rounded rectangle with colored fill for subsection group
                        PowerPoint.Shape subsectionBox = slide.Shapes.AddShape(
                            Office.MsoAutoShapeType.msoShapeRoundedRectangle,
                            boxX, boxY, boxWidth, boxHeight);
                        
                        // Set fill color (cycling through colors) - VERY VISIBLE
                        System.Drawing.Color fillColor = subsectionColors[colorIndex % subsectionColors.Length];
                        subsectionBox.Fill.Visible = Office.MsoTriState.msoTrue;
                        subsectionBox.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(fillColor);
                        subsectionBox.Fill.Transparency = 0.3f; // Only 30% transparent = 70% visible!
                        
                        // Add thick visible border
                        subsectionBox.Line.Visible = Office.MsoTriState.msoTrue;
                        subsectionBox.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(fillColor);
                        subsectionBox.Line.Weight = 2.0f;
                        subsectionBox.Line.Transparency = 0.0f; // Solid border
                        
                        try
                        {
                            subsectionBox.Adjustments[1] = 0.25f; // Rounded corners
                        }
                        catch { }
                        
                        subsectionBox.Tags.Add("NavBar", "True");
                        subsectionBox.Tags.Add("SubsectionBox", group.SubsectionName);
                        subsectionBox.ZOrder(Office.MsoZOrderCmd.msoSendToBack);
                        
                        // DEBUG: Log that we created a box
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
                    
                    PowerPoint.Shape circle = slide.Shapes.AddShape(
                        Office.MsoAutoShapeType.msoShapeOval,
                        circleX, circleY, circleSize, circleSize);
                    
                    // Get this slide's subsection
                    string thisSlideSubsection = "";
                    try
                    {
                        PowerPoint.Slide thisSlide = presentation.Slides[slideIndexInPresentation];
                        thisSlideSubsection = thisSlide.Tags["Subsection"];
                    }
                    catch { }

                    if (slideIndexInPresentation == currentSlideIndex)
                    {
                        // CURRENT SLIDE: White filled
                        circle.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                        circle.Line.Visible = Office.MsoTriState.msoFalse;
                    }
                    else if (!string.IsNullOrEmpty(currentSlideSubsection) && 
                             !string.IsNullOrEmpty(thisSlideSubsection) && 
                             thisSlideSubsection == currentSlideSubsection)
                    {
                        // SAME SUBSECTION: Red border, black inside
                        circle.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
                        circle.Fill.Visible = Office.MsoTriState.msoTrue;
                        circle.Line.Visible = Office.MsoTriState.msoTrue;
                        circle.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                        circle.Line.Weight = 2.0f;
                    }
                    else
                    {
                        // OTHER SLIDES: White hollow
                        circle.Fill.Visible = Office.MsoTriState.msoFalse;
                        circle.Line.Visible = Office.MsoTriState.msoTrue;
                        circle.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                        circle.Line.Weight = 1.5f;
                    }
                    
                    circle.Tags.Add("NavBar", "True");
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
    }
}
