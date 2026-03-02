using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1
{
    public static class ZoomAnimTour
    {
        private const float Margin = 0.92f;
        private const float Duration = 0.65f;
        private const float MinSize = 8f;

        public static void Build(PowerPoint.Application app)
        {
            var win = app?.ActiveWindow;
            if (win == null) throw new InvalidOperationException("No active PowerPoint window.");

            if (win.View == null || win.View.Type != PowerPoint.PpViewType.ppViewNormal)
                throw new InvalidOperationException("Go to Normal view and select shapes.");

            var sourceSlide = win.View.Slide;
            if (sourceSlide == null) throw new InvalidOperationException("No active slide.");

            var pres = app.ActivePresentation;
            if (pres == null) throw new InvalidOperationException("No active presentation.");

            if (win.Selection == null || win.Selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                throw new InvalidOperationException("Select zoom area shapes first (or click Select Zoom Areas).");

            var sr = win.Selection.ShapeRange;
            if (sr == null || sr.Count < 1)
                throw new InvalidOperationException("Select at least 1 zoom area shape.");

            // Keep only shapes with real area
            var areas = new List<PowerPoint.Shape>();
            for (int i = 1; i <= sr.Count; i++)
            {
                var sh = sr[i];
                if (sh.Width >= MinSize && sh.Height >= MinSize)
                    areas.Add(sh);
            }

            if (areas.Count == 0)
                throw new InvalidOperationException("Selected shapes are invalid (too small / lines). Select shapes with area.");

            areas = areas.OrderBy(s => s.Top).ThenBy(s => s.Left).ToList();

            float slideW = pres.PageSetup.SlideWidth;
            float slideH = pres.PageSetup.SlideHeight;

            // Export current slide as image (so we zoom EVERYTHING)
            string tmpFile = Path.Combine(Path.GetTempPath(), $"zoomtour_{Guid.NewGuid():N}.png");
            sourceSlide.Export(tmpFile, "PNG", (int)(slideW * 2), (int)(slideH * 2));

            // New blank slide after current
            var animSlide = pres.Slides.Add(sourceSlide.SlideIndex + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);

            // Add the exported image covering the slide
            var pic = animSlide.Shapes.AddPicture(
                tmpFile,
                Office.MsoTriState.msoFalse,
                Office.MsoTriState.msoTrue,
                0, 0, slideW, slideH);

            // Click controlled
            animSlide.SlideShowTransition.AdvanceOnClick = Office.MsoTriState.msoTrue;
            animSlide.SlideShowTransition.AdvanceOnTime = Office.MsoTriState.msoFalse;

            var seq = animSlide.TimeLine.MainSequence;

            foreach (var a in areas)
            {
                float w = a.Width, h = a.Height, left = a.Left, top = a.Top;

                float scale = Math.Min(slideW / w, slideH / h) * Margin;

                float cx = left + w / 2f;
                float cy = top + h / 2f;

                // scaling about top-left of picture (0,0)
                float scaledCX = cx * scale;
                float scaledCY = cy * scale;

                float slideCX = slideW / 2f;
                float slideCY = slideH / 2f;

                float dx = slideCX - scaledCX;
                float dy = slideCY - scaledCY;

                // CLICK: zoom IN (grow + move together)
                var growIn = AddGrowShrink(seq, pic, scale * 100f, PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                growIn.Timing.Duration = Duration;

                var moveIn = AddMotionBy(seq, pic, dx, dy, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                moveIn.Timing.Duration = Duration;

                // CLICK: zoom OUT
                var growOut = AddGrowShrink(seq, pic, 100f / scale, PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                growOut.Timing.Duration = Duration;

                var moveOut = AddMotionBy(seq, pic, -dx, -dy, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                moveOut.Timing.Duration = Duration;
            }

            try { File.Delete(tmpFile); } catch { }

            animSlide.Select();
        }

        // ---- Compatibility wrapper for AddEffect across different interop signatures ----
        private static PowerPoint.Effect AddEffectCompat(
            PowerPoint.Sequence seq,
            PowerPoint.Shape shape,
            PowerPoint.MsoAnimEffect effect,
            PowerPoint.MsoAnimTriggerType trigger)
        {
            dynamic dseq = seq;
            object levelNone = PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone;

            // Try common signatures (different versions swap order / add index)
            try { return (PowerPoint.Effect)dseq.AddEffect(shape, effect, levelNone, trigger); } catch { }
            try { return (PowerPoint.Effect)dseq.AddEffect(shape, effect, trigger, levelNone); } catch { }

            try { return (PowerPoint.Effect)dseq.AddEffect(shape, effect, levelNone, trigger, seq.Count + 1); } catch { }
            try { return (PowerPoint.Effect)dseq.AddEffect(shape, effect, trigger, seq.Count + 1); } catch { }

            // Last resort: minimal args
            try { return (PowerPoint.Effect)dseq.AddEffect(shape, effect); } catch { }

            throw new InvalidOperationException("PowerPoint AddEffect signature not compatible on this version.");
        }

        private static PowerPoint.Effect AddGrowShrink(
            PowerPoint.Sequence seq,
            PowerPoint.Shape shape,
            float percent,
            PowerPoint.MsoAnimTriggerType trigger)
        {
            var eff = AddEffectCompat(seq, shape, PowerPoint.MsoAnimEffect.msoAnimEffectGrowShrink, trigger);

            // Amount vs Size differs by build
            try { eff.EffectParameters.Amount = percent; }
            catch { try { eff.EffectParameters.Size = percent; } catch { } }

            return eff;
        }

        private static PowerPoint.Effect AddMotionBy(
            PowerPoint.Sequence seq,
            PowerPoint.Shape shape,
            float dx,
            float dy,
            PowerPoint.MsoAnimTriggerType trigger)
        {
            // Use a widely available motion path effect (Right)
            var eff = AddEffectCompat(seq, shape, PowerPoint.MsoAnimEffect.msoAnimEffectPathRight, trigger);

            // Find or add Motion behavior
            PowerPoint.AnimationBehavior motion = null;
            try
            {
                for (int i = 1; i <= eff.Behaviors.Count; i++)
                {
                    var b = eff.Behaviors[i];
                    if (b.Type == PowerPoint.MsoAnimType.msoAnimTypeMotion)
                    {
                        motion = b;
                        break;
                    }
                }
            }
            catch { }

            try
            {
                if (motion == null)
                    motion = eff.Behaviors.Add(PowerPoint.MsoAnimType.msoAnimTypeMotion);

                motion.MotionEffect.ByX = dx;
                motion.MotionEffect.ByY = dy;
            }
            catch
            {
                // If MotionEffect.ByX/ByY isn't supported on your PP build,
                // tell me what error you get and I’ll switch to a different motion method.
            }

            return eff;
        }
    }
}