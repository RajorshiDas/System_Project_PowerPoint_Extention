using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PowerPointAddIn1
{
    // ── Settings data object ─────────────────────────────────────────────────
    internal class SpotlightSettings
    {
        public float TransparencyPercent { get; set; } = 50f;  // 0 = opaque, 100 = clear
        public float SoftEdgesPoints     { get; set; } = 10f;  // 0 = none
        public Color OverlayColor        { get; set; } = Color.Black;
    }

    // ── Feature logic ────────────────────────────────────────────────────────
    internal static class SpotlightFeature
    {
        // Export at 2× slide-point dimensions for sharp crops
        private const int ExportScale = 2;

        // ── State ─────────────────────────────────────────────────────────────
        internal static int SourceSlideIndex { get; private set; } = -1;
        internal static readonly List<string> SpotlightShapeNames = new List<string>();
        internal static SpotlightSettings Settings { get; set; } = new SpotlightSettings();

        // ─────────────────────────────────────────────────────────────────────
        // 1. SELECT EFFECT AREAS
        //    Reads the current selection and registers shapes as spotlight areas.
        //    Shapes are prefixed SPOT_ so they can be identified later.
        // ─────────────────────────────────────────────────────────────────────
        internal static bool SelectAreas(PowerPoint.Application app)
        {
            if (app?.ActivePresentation == null)
            {
                Warn("Please open a presentation first.", "No Presentation");
                return false;
            }

            var win = app.ActiveWindow;
            if (win.Selection == null ||
                win.Selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                Warn("Select one or more shapes on the slide that define the spotlight areas,\n" +
                     "then click 'Select Effect Areas'.",
                     "No Shapes Selected");
                return false;
            }

            var sr = win.Selection.ShapeRange;
            if (sr.Count < 1) { Warn("No shapes selected.", "No Shapes"); return false; }

            SpotlightShapeNames.Clear();
            SourceSlideIndex = win.View.Slide.SlideIndex;

            for (int i = 1; i <= sr.Count; i++)
            {
                var sh = sr[i];
                if (!sh.Name.StartsWith("SPOT_", StringComparison.OrdinalIgnoreCase))
                    sh.Name = "SPOT_" + i + "_" + sh.Name;
                SpotlightShapeNames.Add(sh.Name);
            }

            MessageBox.Show(
                $"{SpotlightShapeNames.Count} spotlight area(s) stored from Slide {SourceSlideIndex}.\n\n" +
                "• Adjust look via 'Settings'.\n" +
                "• Click 'Create Spotlight' to generate the effect slide.",
                "Areas Selected",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);

            return true;
        }

        // ─────────────────────────────────────────────────────────────────────
        // 2. SETTINGS
        //    Opens the SpotlightSettingsForm and saves any changes.
        // ─────────────────────────────────────────────────────────────────────
        internal static void ShowSettings()
        {
            using (var dlg = new SpotlightSettingsForm(Settings))
            {
                if (dlg.ShowDialog() == DialogResult.OK)
                    Settings = dlg.Result;
            }
        }

        // ─────────────────────────────────────────────────────────────────────
        // 3. CREATE SPOTLIGHT SLIDE
        //
        //    Steps:
        //      1. Hide the SPOT_ marker shapes and export source slide to PNG.
        //      2. Add a new blank slide after the source.
        //      3. Fill the new slide with the exported image (full background).
        //      4. Overlay a dark semi-transparent rectangle on the whole slide.
        //      5. For each spotlight area: crop that region from the export,
        //         place it back at the exact same position, apply soft edges.
        //    Result: everything is darkened except the spotlight windows.
        // ─────────────────────────────────────────────────────────────────────
        internal static void CreateSpotlight(PowerPoint.Application app)
        {
            if (app?.ActivePresentation == null)
            {
                Warn("Please open a presentation first.", "No Presentation");
                return;
            }
            if (SourceSlideIndex < 1 || SpotlightShapeNames.Count == 0)
            {
                Warn("No spotlight areas stored.\n\nClick 'Select Effect Areas' first.",
                     "No Areas");
                return;
            }

            var pres  = app.ActivePresentation;
            if (SourceSlideIndex > pres.Slides.Count)
            {
                Warn("The source slide no longer exists.", "Error");
                return;
            }

            var   src    = pres.Slides[SourceSlideIndex];
            float slideW = pres.PageSetup.SlideWidth;
            float slideH = pres.PageSetup.SlideHeight;
            int   expW   = (int)(slideW * ExportScale);
            int   expH   = (int)(slideH * ExportScale);

            string tmpDir  = Path.GetTempPath();
            string bgFile  = Path.Combine(tmpDir, "spot_bg_"  + Guid.NewGuid().ToString("N") + ".png");

            // ── Temporarily hide marker shapes for a clean export ─────────────
            foreach (string n in SpotlightShapeNames)
            {
                var sh = FindShape(src, n);
                if (sh != null) sh.Visible = Office.MsoTriState.msoFalse;
            }

            try
            {
                src.Export(bgFile, "PNG", expW, expH);
            }
            finally
            {
                foreach (string n in SpotlightShapeNames)
                {
                    var sh = FindShape(src, n);
                    if (sh != null) sh.Visible = Office.MsoTriState.msoTrue;
                }
            }

            var cropFiles = new List<string>();
            try
            {
                // ── New blank slide after source ──────────────────────────────
                var ns = pres.Slides.Add(SourceSlideIndex + 1,
                                         PowerPoint.PpSlideLayout.ppLayoutBlank);

                // Remove any placeholder/layout shapes that sneak in
                for (int i = ns.Shapes.Count; i >= 1; i--)
                {
                    try { ns.Shapes[i].Delete(); } catch { }
                }

                // ── 1. Full-slide background (original slide image) ───────────
                ns.Shapes.AddPicture(bgFile,
                    Office.MsoTriState.msoFalse,
                    Office.MsoTriState.msoTrue,
                    0f, 0f, slideW, slideH);

                // ── 2. Dark overlay ───────────────────────────────────────────
                var overlay = ns.Shapes.AddShape(
                    Office.MsoAutoShapeType.msoShapeRectangle,
                    0f, 0f, slideW, slideH);
                overlay.Line.Visible    = Office.MsoTriState.msoFalse;
                overlay.Fill.ForeColor.RGB = ColorTranslator.ToOle(Settings.OverlayColor);
                // Transparency: 0 = fully opaque, 1 = fully clear
                overlay.Fill.Transparency = Settings.TransparencyPercent / 100f;

                // ── 3. Spotlight cutouts (any shape type) ────────────────────────
                //
                // For each marker shape:
                //   a) Crop the exported PNG to the shape's bounding box.
                //   b) Copy/Paste the shape itself onto the spotlight slide so its
                //      exact geometry (oval, freeform, polygon…) is preserved.
                //   c) Apply UserPicture — PowerPoint clips the image to the shape
                //      outline, making any shape work as a spotlight window.
                foreach (string shapeName in SpotlightShapeNames)
                {
                    var area = FindShape(src, shapeName);
                    if (area == null) continue;

                    // ── a) Bounding-box crop ───────────────────────────────────
                    int cx = Math.Max(0, (int)(area.Left  * ExportScale));
                    int cy = Math.Max(0, (int)(area.Top   * ExportScale));
                    int cw = Math.Max(1, Math.Min((int)(area.Width  * ExportScale), expW - cx));
                    int ch = Math.Max(1, Math.Min((int)(area.Height * ExportScale), expH - cy));

                    string cropFile = Path.Combine(tmpDir,
                        "spot_crop_" + Guid.NewGuid().ToString("N") + ".png");
                    cropFiles.Add(cropFile);

                    using (var bmp = new Bitmap(bgFile))
                    using (var cropped = bmp.Clone(new Rectangle(cx, cy, cw, ch), bmp.PixelFormat))
                    {
                        cropped.Save(cropFile, ImageFormat.Png);
                    }

                    // ── b) Copy marker shape → paste onto spotlight slide ──────
                    //      Preserves exact geometry (freeform nodes, oval, etc.)
                    area.Copy();
                    PowerPoint.ShapeRange pasted = ns.Shapes.Paste();
                    PowerPoint.Shape spotShape = pasted[1];

                    // Paste may apply an offset — restore original position/size
                    spotShape.Left   = area.Left;
                    spotShape.Top    = area.Top;
                    spotShape.Width  = area.Width;
                    spotShape.Height = area.Height;

                    // ── c) Picture fill clipped to shape outline ───────────────
                    //      UserPicture stretches the crop across the bounding box;
                    //      the shape outline clips it to the exact shape boundary.
                    spotShape.Fill.UserPicture(cropFile);
                    spotShape.Line.Visible = Office.MsoTriState.msoFalse;

                    if (Settings.SoftEdgesPoints > 0)
                    {
                        try { spotShape.SoftEdge.Radius = Settings.SoftEdgesPoints; }
                        catch { }
                    }
                }

                MessageBox.Show(
                    $"Spotlight slide created after Slide {SourceSlideIndex}.\n" +
                    $"{SpotlightShapeNames.Count} spotlight area(s) applied.\n\n" +
                    "Tip: adjust transparency / soft edges via 'Settings' and recreate.",
                    "Spotlight Created",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            finally
            {
                foreach (string f in cropFiles) TryDeleteTempFile(f);
                TryDeleteTempFile(bgFile);
            }
        }

        // ── Helpers ───────────────────────────────────────────────────────────
        private static PowerPoint.Shape FindShape(PowerPoint.Slide slide, string name)
        {
            for (int i = 1; i <= slide.Shapes.Count; i++)
                if (string.Equals(slide.Shapes[i].Name, name, StringComparison.OrdinalIgnoreCase))
                    return slide.Shapes[i];
            return null;
        }

        private static void Warn(string msg, string title) =>
            MessageBox.Show(msg, title, MessageBoxButtons.OK, MessageBoxIcon.Warning);

        private static void TryDeleteTempFile(string path)
        {
            if (string.IsNullOrEmpty(path)) return;

            for (int attempt = 0; attempt < 3; attempt++)
            {
                try
                {
                    if (!File.Exists(path)) return;
                    File.Delete(path);
                    return;
                }
                catch (IOException)
                {
                    Thread.Sleep(50);
                }
                catch (UnauthorizedAccessException)
                {
                    Thread.Sleep(50);
                }
                catch
                {
                    return;
                }
            }
        }
    }
}
