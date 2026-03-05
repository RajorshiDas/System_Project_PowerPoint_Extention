using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PowerPointAddIn1
{
    // ── Settings data object ─────────────────────────────────────────────────
    internal class BlurSettings
    {
        public int   Intensity       { get; set; } = 40;  // 0–100
        public float SoftEdgesPoints { get; set; } = 0f;  // 0 = none
    }

    // ── Feature logic ────────────────────────────────────────────────────────
    internal static class BlurFeature
    {
        // Export at 2× slide-point dimensions for sharp crops
        private const int ExportScale = 2;

        internal static BlurSettings Settings { get; set; } = new BlurSettings();

        // ─────────────────────────────────────────────────────────────────────
        // SETTINGS  –  opens the BlurSettingsForm dialog
        // ─────────────────────────────────────────────────────────────────────
        internal static void ShowSettings()
        {
            using (var dlg = new BlurSettingsForm(Settings))
            {
                if (dlg.ShowDialog() == DialogResult.OK)
                    Settings = dlg.Result;
            }
        }

        // ─────────────────────────────────────────────────────────────────────
        // BLUR SELECTED
        //   Creates a new slide after the source. The new slide shows the
        //   original content as a full-slide background with blurred crops
        //   placed over each selected shape area (clipped to its exact outline).
        //
        //   Steps:
        //     1. Hide marker shapes; export the slide to PNG; restore them.
        //     2. Add a new blank slide after the source.
        //     3. Place the unblurred export as the full-slide background.
        //     4. For each area: crop + blur the image, copy-paste the original
        //        shape onto the new slide, fill with the blurred crop.
        //        Apply soft edges if configured.
        // ─────────────────────────────────────────────────────────────────────
        internal static void BlurSelected(PowerPoint.Application app)
        {
            if (!TryGetSelection(app, out PowerPoint.Slide slide, out List<ShapeInfo> areas))
                return;

            var   pres   = app.ActivePresentation;
            float slideW = pres.PageSetup.SlideWidth;
            float slideH = pres.PageSetup.SlideHeight;
            int   expW   = (int)(slideW * ExportScale);
            int   expH   = (int)(slideH * ExportScale);

            string bgFile = TempFile("blur_bg", "png");
            var cropFiles = new List<string>();

            HideShapes(slide, areas);
            try   { slide.Export(bgFile, "PNG", expW, expH); }
            finally { ShowShapes(slide, areas); }

            try
            {
                // ── New blank slide after source ──────────────────────────────
                var ns = pres.Slides.Add(slide.SlideIndex + 1,
                    PowerPoint.PpSlideLayout.ppLayoutBlank);
                for (int i = ns.Shapes.Count; i >= 1; i--)
                    try { ns.Shapes[i].Delete(); } catch { }

                // ── 1. Full-slide background (original, unblurred) ────────────
                ns.Shapes.AddPicture(bgFile,
                    Office.MsoTriState.msoFalse,
                    Office.MsoTriState.msoTrue,
                    0f, 0f, slideW, slideH);

                // ── 2. Blurred crops for each selected area ───────────────────
                using (var fullBmp = new Bitmap(bgFile))
                {
                    foreach (var a in areas)
                    {
                        int cx = Math.Max(0, (int)(a.Left  * ExportScale));
                        int cy = Math.Max(0, (int)(a.Top   * ExportScale));
                        int cw = Math.Max(1, Math.Min((int)(a.Width  * ExportScale), expW - cx));
                        int ch = Math.Max(1, Math.Min((int)(a.Height * ExportScale), expH - cy));

                        string cropFile = TempFile("blur_sel", "png");
                        cropFiles.Add(cropFile);

                        using (var crop    = fullBmp.Clone(new Rectangle(cx, cy, cw, ch), fullBmp.PixelFormat))
                        using (var blurred = ApplyBlur(crop, Settings.Intensity))
                            blurred.Save(cropFile, ImageFormat.Png);

                        // Copy the original shape to preserve its exact geometry;
                        // paste onto the new slide, fill with the blurred crop.
                        var origShape = FindShape(slide, a.Name);
                        if (origShape != null)
                        {
                            origShape.Copy();
                            var pasted    = ns.Shapes.Paste();
                            var blurShape = pasted[1];
                            blurShape.Left   = a.Left;
                            blurShape.Top    = a.Top;
                            blurShape.Width  = a.Width;
                            blurShape.Height = a.Height;
                            blurShape.Fill.UserPicture(cropFile);
                            blurShape.Line.Visible = Office.MsoTriState.msoFalse;
                            if (Settings.SoftEdgesPoints > 0)
                                try { blurShape.SoftEdge.Radius = Settings.SoftEdgesPoints; } catch { }
                        }
                        else
                        {
                            // Fallback: plain rectangle picture
                            ns.Shapes.AddPicture(cropFile,
                                Office.MsoTriState.msoFalse,
                                Office.MsoTriState.msoTrue,
                                a.Left, a.Top, a.Width, a.Height);
                        }
                    }
                }

                MessageBox.Show(
                    $"Blur slide created after Slide {slide.SlideIndex}.\n" +
                    $"{areas.Count} area(s) blurred at intensity {Settings.Intensity}.\n\n" +
                    "Tip: adjust intensity / soft edges via 'Settings' and re-run.",
                    "Blur Selected", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error creating blur slide: " + ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                foreach (string f in cropFiles) TryDeleteFile(f);
                TryDeleteFile(bgFile);
            }
        }

        // ─────────────────────────────────────────────────────────────────────
        // BLUR REMAINDER
        //   The selected shape defines the area that stays in focus.
        //   Everything else on the slide is covered by a blurred overlay.
        //   Delegates to BlurAllExceptImpl (single or multiple shapes).
        // ─────────────────────────────────────────────────────────────────────
        internal static void BlurRemainder(PowerPoint.Application app)
        {
            BlurAllExceptImpl(app, "Blur Remainder");
        }

        // ─────────────────────────────────────────────────────────────────────
        // BLUR ALL EXCEPT SELECTED
        //   Blurs the entire slide except the selected shape(s).
        //   Supports multiple clear-area shapes simultaneously.
        // ─────────────────────────────────────────────────────────────────────
        internal static void BlurAllExcept(PowerPoint.Application app)
        {
            BlurAllExceptImpl(app, "Blur All Except Selected");
        }

        // ─────────────────────────────────────────────────────────────────────
        // BLUR ALL EXCEPT – shared implementation
        //
        //   Creates a new blank slide after the source. The new slide shows a
        //   blurred version of the entire slide as its background, with
        //   unblurred crops placed over each clear-area shape.
        //
        //   Steps:
        //     1. Hide marker shapes; export the slide to PNG; restore them.
        //     2. Create a blurred version of the full-slide export.
        //     3. Add a new blank slide after the source.
        //     4. Place the blurred export as the full-slide background.
        //     5. For each marker shape: crop the *unblurred* export, copy-paste
        //        the original shape onto the new slide, fill with the unblurred
        //        crop. Apply soft edges if configured.
        // ─────────────────────────────────────────────────────────────────────
        private static void BlurAllExceptImpl(PowerPoint.Application app, string title)
        {
            if (!TryGetSelection(app, out PowerPoint.Slide slide, out List<ShapeInfo> areas))
                return;

            var   pres     = app.ActivePresentation;
            float slideW   = pres.PageSetup.SlideWidth;
            float slideH   = pres.PageSetup.SlideHeight;
            int   expW     = (int)(slideW * ExportScale);
            int   expH     = (int)(slideH * ExportScale);

            string bgFile   = TempFile("blur_bg",   "png");
            string blurFile = TempFile("blur_full",  "png");
            var    cropFiles = new List<string>();

            HideShapes(slide, areas);
            try   { slide.Export(bgFile, "PNG", expW, expH); }
            finally { ShowShapes(slide, areas); }

            try
            {
                using (var fullBmp = new Bitmap(bgFile))
                {
                    // Save blurred full-slide image to disk
                    using (var blurredFull = ApplyBlur(fullBmp, Settings.Intensity))
                        blurredFull.Save(blurFile, ImageFormat.Png);

                    // ── New blank slide after source ──────────────────────────
                    var ns = pres.Slides.Add(slide.SlideIndex + 1,
                        PowerPoint.PpSlideLayout.ppLayoutBlank);
                    for (int i = ns.Shapes.Count; i >= 1; i--)
                        try { ns.Shapes[i].Delete(); } catch { }

                    // ── 1. Blurred full-slide background ──────────────────────
                    ns.Shapes.AddPicture(blurFile,
                        Office.MsoTriState.msoFalse,
                        Office.MsoTriState.msoTrue,
                        0f, 0f, slideW, slideH);

                    // ── 2. Unblurred crops for each clear-area shape ──────────
                    foreach (var a in areas)
                    {
                        int cx = Math.Max(0, (int)(a.Left  * ExportScale));
                        int cy = Math.Max(0, (int)(a.Top   * ExportScale));
                        int cw = Math.Max(1, Math.Min((int)(a.Width  * ExportScale), expW - cx));
                        int ch = Math.Max(1, Math.Min((int)(a.Height * ExportScale), expH - cy));

                        string cropFile = TempFile("blur_clear", "png");
                        cropFiles.Add(cropFile);

                        using (var crop = fullBmp.Clone(new Rectangle(cx, cy, cw, ch), fullBmp.PixelFormat))
                            crop.Save(cropFile, ImageFormat.Png);

                        // Copy-paste original shape → fill with unblurred crop.
                        var origShape = FindShape(slide, a.Name);
                        if (origShape != null)
                        {
                            origShape.Copy();
                            var pasted     = ns.Shapes.Paste();
                            var clearShape = pasted[1];
                            clearShape.Left   = a.Left;
                            clearShape.Top    = a.Top;
                            clearShape.Width  = a.Width;
                            clearShape.Height = a.Height;
                            clearShape.Fill.UserPicture(cropFile);
                            clearShape.Line.Visible = Office.MsoTriState.msoFalse;
                            if (Settings.SoftEdgesPoints > 0)
                                try { clearShape.SoftEdge.Radius = Settings.SoftEdgesPoints; } catch { }
                        }
                        else
                        {
                            ns.Shapes.AddPicture(cropFile,
                                Office.MsoTriState.msoFalse,
                                Office.MsoTriState.msoTrue,
                                a.Left, a.Top, a.Width, a.Height);
                        }
                    }
                }

                MessageBox.Show(
                    $"Blur slide created after Slide {slide.SlideIndex}.\n" +
                    $"{areas.Count} area(s) kept clear at intensity {Settings.Intensity}.\n\n" +
                    "Tip: adjust intensity / soft edges via 'Settings' and re-run.",
                    title, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error creating blur slide: " + ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                foreach (string f in cropFiles) TryDeleteFile(f);
                TryDeleteFile(bgFile);
                TryDeleteFile(blurFile);
            }
        }

        // ── Blur algorithm ────────────────────────────────────────────────────
        // Downscale with high-quality bicubic then upscale back to original size.
        // Multiple passes at the same scale give progressively softer results.
        // intensity 0 = no blur, 100 = maximum blur.
        private static Bitmap ApplyBlur(Bitmap source, int intensity)
        {
            if (intensity <= 0) return new Bitmap(source);

            // Scale factor: intensity 1 → ~0.99, intensity 100 → 0.03
            float scale = 1.0f - (intensity / 100.0f) * 0.97f;
            scale = Math.Max(0.03f, scale);

            int sw = Math.Max(1, (int)(source.Width  * scale));
            int sh = Math.Max(1, (int)(source.Height * scale));

            // Pass 1 – downscale
            var small = new Bitmap(sw, sh, PixelFormat.Format32bppArgb);
            using (var g = Graphics.FromImage(small))
            {
                g.InterpolationMode =
                    System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                g.DrawImage(source, 0, 0, sw, sh);
            }

            // Pass 2 – upscale back to original dimensions
            var result = new Bitmap(source.Width, source.Height, PixelFormat.Format32bppArgb);
            using (var g = Graphics.FromImage(result))
            {
                g.InterpolationMode =
                    System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                g.DrawImage(small, 0, 0, source.Width, source.Height);
            }

            small.Dispose();
            return result;
        }

        // ── Selection helpers ─────────────────────────────────────────────────
        private static bool TryGetSelection(PowerPoint.Application app,
            out PowerPoint.Slide slide, out List<ShapeInfo> areas)
        {
            slide = null;
            areas = null;

            if (app == null || app.Presentations.Count == 0)
            {
                Warn("Please open a presentation first.", "No Presentation");
                return false;
            }

            var win = app.ActiveWindow;
            if (win.Selection == null ||
                (win.Selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes &&
                 win.Selection.Type != PowerPoint.PpSelectionType.ppSelectionText))
            {
                Warn("Please select one or more shapes on the slide that define the\n" +
                     "area(s), then click the blur button.", "No Shapes Selected");
                return false;
            }

            var sr = win.Selection.ShapeRange;
            if (sr.Count < 1) { Warn("No shapes selected.", "No Shapes"); return false; }

            slide = win.View.Slide as PowerPoint.Slide;
            areas = new List<ShapeInfo>();
            for (int i = 1; i <= sr.Count; i++)
            {
                var sh = sr[i];
                areas.Add(new ShapeInfo(sh.Name, sh.Left, sh.Top, sh.Width, sh.Height));
            }
            return true;
        }

        private static void HideShapes(PowerPoint.Slide slide, List<ShapeInfo> areas)
        {
            foreach (var a in areas)
            {
                var sh = FindShape(slide, a.Name);
                if (sh != null) sh.Visible = Office.MsoTriState.msoFalse;
            }
        }

        private static void ShowShapes(PowerPoint.Slide slide, List<ShapeInfo> areas)
        {
            foreach (var a in areas)
            {
                var sh = FindShape(slide, a.Name);
                if (sh != null) sh.Visible = Office.MsoTriState.msoTrue;
            }
        }

        private static void DeleteShapes(PowerPoint.Slide slide, List<ShapeInfo> areas)
        {
            foreach (var a in areas)
            {
                var sh = FindShape(slide, a.Name);
                if (sh != null) try { sh.Delete(); } catch { }
            }
        }

        private static PowerPoint.Shape FindShape(PowerPoint.Slide slide, string name)
        {
            for (int i = 1; i <= slide.Shapes.Count; i++)
                if (string.Equals(slide.Shapes[i].Name, name, StringComparison.OrdinalIgnoreCase))
                    return slide.Shapes[i];
            return null;
        }

        private static string TempFile(string prefix, string ext)
            => Path.Combine(Path.GetTempPath(),
                prefix + "_" + Guid.NewGuid().ToString("N") + "." + ext);

        private static void TryDeleteFile(string path)
        {
            try { if (File.Exists(path)) File.Delete(path); } catch { }
        }

        private static void Warn(string msg, string title)
            => MessageBox.Show(msg, title, MessageBoxButtons.OK, MessageBoxIcon.Warning);

        // ── Value object capturing shape geometry at selection time ───────────
        private class ShapeInfo
        {
            public string Name   { get; }
            public float  Left   { get; }
            public float  Top    { get; }
            public float  Width  { get; }
            public float  Height { get; }

            public ShapeInfo(string name, float left, float top, float width, float height)
            {
                Name   = name;
                Left   = left;
                Top    = top;
                Width  = width;
                Height = height;
            }
        }
    }
}
