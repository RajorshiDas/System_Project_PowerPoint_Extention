using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PowerPointAddIn1
{
    /// <summary>
    /// Represents a presentation section for agenda generation.
    /// </summary>
    public class AgendaSectionInfo
    {
        public string Name { get; set; }
        public int FirstSlideIndex { get; set; }
        public int SlideCount { get; set; }
        public List<AgendaSubsectionInfo> Subsections { get; set; } = new List<AgendaSubsectionInfo>();
    }

    /// <summary>
    /// Represents a subsection (contiguous slides sharing the same Subsection tag).
    /// </summary>
    public class AgendaSubsectionInfo
    {
        public string Name { get; set; }
        public int FirstSlideIndex { get; set; }
        public int SlideCount { get; set; }
    }

    /// <summary>
    /// Generates, removes, and refreshes agenda slides in a PowerPoint presentation.
    /// Agenda slides are inserted after slide 1 and tagged so they can be identified later.
    /// </summary>
    public static class AgendaGenerator
    {
        // ── Tag constants ──────────────────────────────────────────────
        private const string AgendaTagName = "AgendaGenerated";
        private const string AgendaTagValue = "True";

        // ── Title constants ────────────────────────────────────────────
        private const string AgendaTitleText = "Agenda";
        private const float TitleLeft = 40f;
        private const float TitleTop = 24f;
        private const float TitleHeight = 50f;
        private const float TitleFontSize = 32f;

        // ── Content layout constants (points) ──────────────────────────
        private const float ContentLeft = 50f;
        private const float ContentTop = 90f;
        private const float ContentRightMargin = 50f;
        private const float BottomMargin = 40f;
        private const float PageNumWidth = 50f;

        // ── Font sizes ─────────────────────────────────────────────────
        private const float SectionFontSize = 16f;
        private const float SubsectionFontSize = 13f;

        // ── Spacing ────────────────────────────────────────────────────
        private const float SectionLineHeight = 28f;
        private const float SubsectionLineHeight = 22f;
        private const float SectionGapBefore = 10f;
        private const float SubsectionIndent = 28f;

        // ── Colors ─────────────────────────────────────────────────────
        private static readonly Color TextColor = Color.Black;

        // ────────────────────────────────────────────────────────────────
        // Internal layout item used for pagination
        // ────────────────────────────────────────────────────────────────
        private class AgendaItem
        {
            public bool IsSection;
            public string Text;
            public int SlideNumber;   // 0 = no page number shown
            public float Height;
        }

        // ================================================================
        //  PUBLIC API
        // ================================================================

        public static void AddAgenda(PowerPoint.Application app)
        {
            if (!ValidatePresentation(app, out PowerPoint.Presentation pres)) return;

            RemoveTaggedSlides(pres);

            List<AgendaSectionInfo> structure = GetPresentationStructure(pres);
            if (structure.Count == 0)
            {
                MessageBox.Show(
                    "No sections found in the presentation.\n" +
                    "Please create sections before generating an agenda.",
                    "No Sections", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int created = GenerateAgendaSlides(pres, structure);

            MessageBox.Show(
                $"{created} agenda slide(s) inserted after slide 1.",
                "Agenda Created", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public static void RemoveAgenda(PowerPoint.Application app)
        {
            if (!ValidatePresentation(app, out PowerPoint.Presentation pres)) return;

            int removed = RemoveTaggedSlides(pres);

            MessageBox.Show(
                removed > 0
                    ? $"{removed} agenda slide(s) removed."
                    : "No generated agenda slides were found.",
                removed > 0 ? "Agenda Removed" : "Nothing to Remove",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public static void RefreshAgenda(PowerPoint.Application app)
        {
            if (!ValidatePresentation(app, out PowerPoint.Presentation pres)) return;

            int removed = RemoveTaggedSlides(pres);

            List<AgendaSectionInfo> structure = GetPresentationStructure(pres);
            if (structure.Count == 0)
            {
                MessageBox.Show(
                    "No sections found in the presentation.\n" +
                    "Please create sections before generating an agenda.",
                    "No Sections", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int created = GenerateAgendaSlides(pres, structure);

            MessageBox.Show(
                $"Agenda refreshed \u2014 {removed} old slide(s) removed, {created} new slide(s) created.",
                "Agenda Refreshed", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // ================================================================
        //  VALIDATION
        // ================================================================

        private static bool ValidatePresentation(
            PowerPoint.Application app, out PowerPoint.Presentation pres)
        {
            pres = null;
            if (app.Presentations.Count == 0)
            {
                MessageBox.Show("Please open a presentation first.",
                    "No Presentation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            pres = app.ActivePresentation;
            if (pres.Slides.Count == 0)
            {
                MessageBox.Show("The presentation has no slides.",
                    "Empty Presentation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            return true;
        }

        // ================================================================
        //  STRUCTURE READING
        // ================================================================

        public static List<AgendaSectionInfo> GetPresentationStructure(
            PowerPoint.Presentation pres)
        {
            var result = new List<AgendaSectionInfo>();
            PowerPoint.SectionProperties sp = pres.SectionProperties;

            for (int i = 1; i <= sp.Count; i++)
            {
                int slideCount = sp.SlidesCount(i);
                if (slideCount == 0) continue;

                var sec = new AgendaSectionInfo
                {
                    Name = sp.Name(i),
                    FirstSlideIndex = sp.FirstSlide(i),
                    SlideCount = slideCount,
                    Subsections = CollectSubsections(pres, sp.FirstSlide(i), slideCount)
                };
                result.Add(sec);
            }

            return result;
        }

        private static List<AgendaSubsectionInfo> CollectSubsections(
            PowerPoint.Presentation pres, int firstSlide, int count)
        {
            var subs = new List<AgendaSubsectionInfo>();
            string current = null;
            int groupStart = 0;
            int groupCount = 0;

            for (int i = 0; i < count; i++)
            {
                string tag = "";
                try { tag = pres.Slides[firstSlide + i].Tags["Subsection"]; }
                catch { }

                if (!string.IsNullOrEmpty(tag))
                {
                    if (tag == current)
                    {
                        groupCount++;
                    }
                    else
                    {
                        FlushGroup(subs, current, firstSlide + groupStart, groupCount);
                        current = tag;
                        groupStart = i;
                        groupCount = 1;
                    }
                }
                else
                {
                    FlushGroup(subs, current, firstSlide + groupStart, groupCount);
                    current = null;
                    groupCount = 0;
                }
            }

            FlushGroup(subs, current, firstSlide + groupStart, groupCount);
            return subs;
        }

        private static void FlushGroup(List<AgendaSubsectionInfo> list,
            string name, int firstSlide, int count)
        {
            if (name != null && count > 0)
            {
                list.Add(new AgendaSubsectionInfo
                {
                    Name = name,
                    FirstSlideIndex = firstSlide,
                    SlideCount = count
                });
            }
        }

        // ================================================================
        //  AGENDA GENERATION (CORE)
        // ================================================================

        private static int GenerateAgendaSlides(
            PowerPoint.Presentation pres, List<AgendaSectionInfo> structure)
        {
            float slideHeight = pres.PageSetup.SlideHeight;
            float slideWidth = pres.PageSetup.SlideWidth;
            float available = slideHeight - ContentTop - BottomMargin;

            // Build flat item list — slide numbers reflect the presentation
            // structure as it exists without agenda slides, which is what the
            // user originally tagged and expects to see.
            List<AgendaItem> items = BuildAgendaItems(structure);
            List<List<AgendaItem>> pages = SplitIntoPages(items, available);

            // Create each agenda slide
            for (int p = 0; p < pages.Count; p++)
            {
                int insertAt = 2 + p; // after slide 1
                PowerPoint.Slide slide = pres.Slides.Add(
                    insertAt, PowerPoint.PpSlideLayout.ppLayoutBlank);

                slide.FollowMasterBackground = Office.MsoTriState.msoFalse;
                slide.Background.Fill.Solid();
                slide.Background.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.White);

                AddTitle(slide, slideWidth, p + 1, pages.Count);
                RenderItems(slide, pages[p], slideWidth, pres, pages.Count);
                TagAgendaSlide(slide);
            }

            return pages.Count;
        }

        // ================================================================
        //  ITEM BUILDING & PAGINATION
        // ================================================================

        private static List<AgendaItem> BuildAgendaItems(
            List<AgendaSectionInfo> structure)
        {
            var items = new List<AgendaItem>();

            for (int s = 0; s < structure.Count; s++)
            {
                AgendaSectionInfo sec = structure[s];
                float gap = s > 0 ? SectionGapBefore : 0f;

                if (sec.Subsections.Count > 0)
                {
                    // Section header without page number (subsections carry the numbers)
                    items.Add(new AgendaItem
                    {
                        IsSection = true,
                        Text = sec.Name,
                        SlideNumber = 0,
                        Height = SectionLineHeight + gap
                    });

                    foreach (AgendaSubsectionInfo sub in sec.Subsections)
                    {
                        items.Add(new AgendaItem
                        {
                            IsSection = false,
                            Text = sub.Name,
                            SlideNumber = sub.FirstSlideIndex,
                            Height = SubsectionLineHeight
                        });
                    }
                }
                else
                {
                    // Section without subsections — show with its starting page number
                    items.Add(new AgendaItem
                    {
                        IsSection = true,
                        Text = sec.Name,
                        SlideNumber = sec.FirstSlideIndex,
                        Height = SectionLineHeight + gap
                    });
                }
            }

            return items;
        }

        private static List<List<AgendaItem>> SplitIntoPages(
            List<AgendaItem> items, float availableHeight)
        {
            var pages = new List<List<AgendaItem>>();
            var page = new List<AgendaItem>();
            float used = 0f;

            foreach (AgendaItem item in items)
            {
                float h = item.Height;

                // Suppress extra top gap for the first item on a new page
                if (page.Count == 0 && item.IsSection)
                    h = SectionLineHeight;

                if (page.Count > 0 && used + h > availableHeight)
                {
                    pages.Add(page);
                    page = new List<AgendaItem>();
                    used = 0f;
                    h = item.IsSection ? SectionLineHeight : item.Height;
                }

                page.Add(item);
                used += h;
            }

            if (page.Count > 0) pages.Add(page);
            if (pages.Count == 0) pages.Add(new List<AgendaItem>());
            return pages;
        }

        // ================================================================
        //  SLIDE RENDERING
        // ================================================================

        private static void AddTitle(PowerPoint.Slide slide, float slideWidth,
            int pageNum, int totalPages)
        {
            string text = totalPages > 1
                ? $"{AgendaTitleText} ({pageNum}/{totalPages})"
                : AgendaTitleText;

            PowerPoint.Shape shape = slide.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal,
                TitleLeft, TitleTop, slideWidth - TitleLeft * 2, TitleHeight);
            shape.TextFrame.TextRange.Text = text;
            shape.TextFrame.TextRange.Font.Size = TitleFontSize;
            shape.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            shape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(TextColor);
            shape.TextFrame.TextRange.Font.Name = "Calibri";
            shape.Line.Visible = Office.MsoTriState.msoFalse;
            shape.Fill.Visible = Office.MsoTriState.msoFalse;
        }

        private static void RenderItems(PowerPoint.Slide slide,
            List<AgendaItem> items, float slideWidth,
            PowerPoint.Presentation pres, int agendaCount)
        {
            float contentWidth = slideWidth - ContentLeft - ContentRightMargin;
            float y = ContentTop;

            for (int idx = 0; idx < items.Count; idx++)
            {
                AgendaItem item = items[idx];
                float gap = (idx > 0 && item.IsSection) ? SectionGapBefore : 0f;
                y += gap;

                if (item.IsSection)
                {
                    RenderSectionHeader(slide, item, y, contentWidth, slideWidth, pres, agendaCount);
                    y += SectionLineHeight;
                }
                else
                {
                    RenderSubsectionEntry(slide, item, y, contentWidth, slideWidth, pres, agendaCount);
                    y += SubsectionLineHeight;
                }
            }
        }

        private static void RenderSectionHeader(PowerPoint.Slide slide,
            AgendaItem item, float y, float contentWidth, float slideWidth,
            PowerPoint.Presentation pres, int agendaCount)
        {
            // Section name (left-aligned, bold)
            PowerPoint.Shape nameShape = slide.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal,
                ContentLeft, y, contentWidth - PageNumWidth, SectionLineHeight);
            nameShape.TextFrame.TextRange.Text = item.Text;
            nameShape.TextFrame.TextRange.Font.Size = SectionFontSize;
            nameShape.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            nameShape.TextFrame.TextRange.Font.Color.RGB =
                ColorTranslator.ToOle(TextColor);
            nameShape.TextFrame.TextRange.Font.Name = "Calibri";
            nameShape.TextFrame.MarginLeft = 0f;
            nameShape.TextFrame.MarginTop = 3f;
            nameShape.TextFrame.MarginBottom = 3f;
            nameShape.TextFrame.WordWrap = Office.MsoTriState.msoFalse;
            nameShape.Line.Visible = Office.MsoTriState.msoFalse;
            nameShape.Fill.Visible = Office.MsoTriState.msoFalse;

            if (item.SlideNumber > 0)
            {
                ApplySlideLink(nameShape, pres, item.SlideNumber, agendaCount);

                // Page number
                PowerPoint.Shape numShape = slide.Shapes.AddTextbox(
                    Office.MsoTextOrientation.msoTextOrientationHorizontal,
                    slideWidth - ContentRightMargin - PageNumWidth, y,
                    PageNumWidth, SectionLineHeight);
                numShape.TextFrame.TextRange.Text = item.SlideNumber.ToString();
                numShape.TextFrame.TextRange.Font.Size = SectionFontSize;
                numShape.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                numShape.TextFrame.TextRange.Font.Color.RGB =
                    ColorTranslator.ToOle(TextColor);
                numShape.TextFrame.TextRange.Font.Name = "Calibri";
                numShape.TextFrame.TextRange.ParagraphFormat.Alignment =
                    PowerPoint.PpParagraphAlignment.ppAlignRight;
                numShape.TextFrame.MarginRight = 6f;
                numShape.TextFrame.MarginTop = 3f;
                numShape.TextFrame.MarginBottom = 3f;
                numShape.Line.Visible = Office.MsoTriState.msoFalse;
                numShape.Fill.Visible = Office.MsoTriState.msoFalse;

                ApplySlideLink(numShape, pres, item.SlideNumber, agendaCount);
            }
        }

        private static void RenderSubsectionEntry(PowerPoint.Slide slide,
            AgendaItem item, float y, float contentWidth, float slideWidth,
            PowerPoint.Presentation pres, int agendaCount)
        {
            float textWidth = contentWidth - SubsectionIndent - PageNumWidth;

            // Build display text: bullet + name + dot leader
            string bullet = "\u2022  ";
            string nameText = bullet + item.Text;
            string dots = BuildDotLeader(nameText, textWidth, SubsectionFontSize);
            string fullText = nameText + " " + dots;

            PowerPoint.Shape textShape = slide.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal,
                ContentLeft + SubsectionIndent, y, textWidth, SubsectionLineHeight);
            textShape.TextFrame.TextRange.Text = fullText;
            textShape.TextFrame.TextRange.Font.Size = SubsectionFontSize;
            textShape.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoFalse;
            textShape.TextFrame.TextRange.Font.Color.RGB =
                ColorTranslator.ToOle(TextColor);
            textShape.TextFrame.TextRange.Font.Name = "Calibri";
            textShape.TextFrame.MarginLeft = 0f;
            textShape.TextFrame.MarginRight = 0f;
            textShape.TextFrame.MarginTop = 1f;
            textShape.TextFrame.MarginBottom = 1f;
            textShape.TextFrame.WordWrap = Office.MsoTriState.msoFalse;
            textShape.Line.Visible = Office.MsoTriState.msoFalse;
            textShape.Fill.Visible = Office.MsoTriState.msoFalse;

            ApplySlideLink(textShape, pres, item.SlideNumber, agendaCount);

            // Right-aligned page number
            PowerPoint.Shape numShape = slide.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal,
                slideWidth - ContentRightMargin - PageNumWidth, y,
                PageNumWidth, SubsectionLineHeight);
            numShape.TextFrame.TextRange.Text = item.SlideNumber.ToString();
            numShape.TextFrame.TextRange.Font.Size = SubsectionFontSize;
            numShape.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoFalse;
            numShape.TextFrame.TextRange.Font.Color.RGB =
                ColorTranslator.ToOle(TextColor);
            numShape.TextFrame.TextRange.Font.Name = "Calibri";
            numShape.TextFrame.TextRange.ParagraphFormat.Alignment =
                PowerPoint.PpParagraphAlignment.ppAlignRight;
            numShape.TextFrame.MarginLeft = 0f;
            numShape.TextFrame.MarginRight = 4f;
            numShape.TextFrame.MarginTop = 1f;
            numShape.TextFrame.MarginBottom = 1f;
            numShape.Line.Visible = Office.MsoTriState.msoFalse;
            numShape.Fill.Visible = Office.MsoTriState.msoFalse;

            ApplySlideLink(numShape, pres, item.SlideNumber, agendaCount);
        }

        private static void ApplySlideLink(PowerPoint.Shape shape,
            PowerPoint.Presentation pres, int displayedSlideNum, int agendaCount)
        {
            if (displayedSlideNum <= 0) return;

            try
            {
                // The displayed number is the pre-agenda index.
                // After inserting agenda slides at position 2, every content
                // slide at index >= 2 has shifted by agendaCount.
                int actualIndex = displayedSlideNum >= 2
                    ? displayedSlideNum + agendaCount
                    : displayedSlideNum;

                if (actualIndex < 1 || actualIndex > pres.Slides.Count) return;

                PowerPoint.Slide target = pres.Slides[actualIndex];
                string subAddress = $"{target.SlideID},{target.SlideIndex},{target.Name}";

                var click = shape.ActionSettings[PowerPoint.PpMouseActivation.ppMouseClick];
                click.Action = PowerPoint.PpActionType.ppActionNone;
                click.Action = PowerPoint.PpActionType.ppActionHyperlink;
                click.Hyperlink.Address = "";
                click.Hyperlink.SubAddress = subAddress;

                var over = shape.ActionSettings[PowerPoint.PpMouseActivation.ppMouseOver];
                over.Action = PowerPoint.PpActionType.ppActionNone;
            }
            catch { }
        }

        private static string BuildDotLeader(
            string nameText, float boxWidth, float fontSize)
        {
            float charWidth = fontSize * 0.55f;
            int maxChars = (int)(boxWidth / charWidth);
            int dotsNeeded = maxChars - nameText.Length - 2;
            if (dotsNeeded < 3) dotsNeeded = 3;
            if (dotsNeeded > 50) dotsNeeded = 50;
            return new string('.', dotsNeeded);
        }

        // ================================================================
        //  TAGGING & REMOVAL
        // ================================================================

        private static void TagAgendaSlide(PowerPoint.Slide slide)
        {
            slide.Tags.Add(AgendaTagName, AgendaTagValue);
        }

        public static int RemoveTaggedSlides(PowerPoint.Presentation pres)
        {
            int removed = 0;
            // Iterate backwards to avoid index-shift issues when deleting
            for (int i = pres.Slides.Count; i >= 1; i--)
            {
                try
                {
                    if (pres.Slides[i].Tags[AgendaTagName] == AgendaTagValue)
                    {
                        pres.Slides[i].Delete();
                        removed++;
                    }
                }
                catch { }
            }
            return removed;
        }
    }
}
