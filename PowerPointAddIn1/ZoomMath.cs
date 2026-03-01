using System;
using System.Collections.Generic;
using System.IO;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PowerPointAddIn1
{
    public struct RectF
    {
        public float L, T, W, H;

        public RectF(float left, float top, float width, float height)
        {
            L = left;
            T = top;
            W = width;
            H = height;
        }

        public float R => L + W;
        public float B => T + H;
        public float CX => L + W / 2f;
        public float CY => T + H / 2f;
    }

    public struct CameraPose
    {
        public float Scale;
        public float X;
        public float Y;

        public CameraPose(float scale, float x, float y)
        {
            Scale = scale;
            X = x;
            Y = y;
        }
    }

    public class ZoomLabSettings
    {
        public bool SeparateSlides { get; set; }
        public float ZoomInSeconds { get; set; }
        public float PanSeconds { get; set; }
        public float HoldSeconds { get; set; }
        public float ZoomOutSeconds { get; set; }
        public float PaddingPoints { get; set; }
        public float MaxScale { get; set; }

        public ZoomLabSettings()
        {
            SeparateSlides = false;
            ZoomInSeconds = 0.6f;
            PanSeconds = 0.55f;
            HoldSeconds = 0.2f;
            ZoomOutSeconds = 0.55f;
            PaddingPoints = 12f;
            MaxScale = 5f;
        }
    }

    public class AnimationBuilder
    {
        public static PowerPoint.Effect AddZoom(PowerPoint.Slide slide, PowerPoint.Shape picture, 
            float targetScale, float duration, PowerPoint.MsoAnimTriggerType trigger)
        {
            PowerPoint.Sequence seq = slide.TimeLine.MainSequence;

            PowerPoint.Effect fx = seq.AddEffect(
                picture,
                PowerPoint.MsoAnimEffect.msoAnimEffectGrowShrink,
                PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone,
                trigger);

            fx.Timing.Duration = duration;

            float pct = targetScale * 100f;
            PowerPoint.AnimationBehavior beh = fx.Behaviors[1];
            beh.ScaleEffect.ByX = pct;
            beh.ScaleEffect.ByY = pct;

            return fx;
        }

        public static PowerPoint.Effect AddPan(PowerPoint.Slide slide, PowerPoint.Shape picture, 
            float fromX, float fromY, float toX, float toY, float duration, PowerPoint.MsoAnimTriggerType trigger)
        {
            PowerPoint.Sequence seq = slide.TimeLine.MainSequence;

            float dx = toX - fromX;
            float dy = toY - fromY;

            // (MsoAnimEffect)47 = msoAnimEffectPathCustom — not exposed in older interop assemblies
            PowerPoint.Effect fx = seq.AddEffect(
                picture,
                (PowerPoint.MsoAnimEffect)47,
                PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone,
                trigger);

            fx.Timing.Duration = duration;

            PowerPoint.AnimationBehavior beh = fx.Behaviors[1];
            beh.MotionEffect.Path = "M 0 0 L " + dx.ToString("F2") + " " + dy.ToString("F2");

            return fx;
        }
    }

    public static class ZoomMath
    {
        public static RectF InflateAndClamp(RectF rect, float pad, float slideW, float slideH)
        {
            float newL = Math.Max(0, rect.L - pad);
            float newT = Math.Max(0, rect.T - pad);
            float newR = Math.Min(slideW, rect.R + pad);
            float newB = Math.Min(slideH, rect.B + pad);

            return new RectF(newL, newT, newR - newL, newB - newT);
        }

        public static CameraPose PoseForRect(RectF rect, float slideW, float slideH)
        {
            float scale = Math.Min(slideW / rect.W, slideH / rect.H);
            float x = slideW / 2f - scale * rect.CX;
            float y = slideH / 2f - scale * rect.CY;

            return new CameraPose(scale, x, y);
        }

        public static void AddOrderBadges(PowerPoint.Slide slide, List<RectF> rects)
        {
            for (int i = 0; i < rects.Count; i++)
            {
                AddOneBadge(slide, rects[i], i + 1);
            }
        }

        public static PowerPoint.Shape AddOneBadge(PowerPoint.Slide slide, RectF rect, int index)
        {
            PowerPoint.Presentation presentation = slide.Parent as PowerPoint.Presentation;
            float slideWidth = presentation.PageSetup.SlideWidth;
            float slideHeight = presentation.PageSetup.SlideHeight;

            float badgeSize = 24f;
            float offset = 8f;

            float badgeX = rect.L - offset;
            float badgeY = rect.T - offset;

            badgeX = Math.Max(0, Math.Min(badgeX, slideWidth - badgeSize));
            badgeY = Math.Max(0, Math.Min(badgeY, slideHeight - badgeSize));

            PowerPoint.Shape badge = slide.Shapes.AddShape(
                Office.MsoAutoShapeType.msoShapeOval,
                badgeX, badgeY, badgeSize, badgeSize);

            badge.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(255, 68, 68, 68));
            badge.Line.Visible = Office.MsoTriState.msoFalse;

            badge.TextFrame.TextRange.Text = index.ToString();
            badge.TextFrame.TextRange.Font.Size = 12f;
            badge.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
            badge.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            badge.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;
            badge.TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;

            return badge;
        }

        public static void BuildSingleAnimationSlide(PowerPoint.Application app, List<RectF> rectangles)
        {
            if (app.ActivePresentation == null)
            {
                throw new InvalidOperationException("Please open a presentation first.");
            }

            if (rectangles == null || rectangles.Count == 0)
            {
                throw new InvalidOperationException("No zoom areas provided.");
            }

            PowerPoint.Presentation presentation = app.ActivePresentation;
            PowerPoint.Slide sourceSlide = app.ActiveWindow.View.Slide;

            float slideWidth = presentation.PageSetup.SlideWidth;
            float slideHeight = presentation.PageSetup.SlideHeight;

            // Export current slide to PNG
            string tempPath = Path.Combine(Path.GetTempPath(), "slide_export_" + Guid.NewGuid().ToString() + ".png");
            sourceSlide.Export(tempPath, "PNG", (int)slideWidth, (int)slideHeight);

            // Create new blank slide
            PowerPoint.Slide newSlide = presentation.Slides.Add(sourceSlide.SlideIndex + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);

            // Insert exported image as full-slide picture
            PowerPoint.Shape pictureShape = newSlide.Shapes.AddPicture(
                tempPath,
                Office.MsoTriState.msoFalse,
                Office.MsoTriState.msoTrue,
                0, 0, slideWidth, slideHeight);

            // Initial state - full slide view
            CameraPose initialPose = new CameraPose(1.0f, 0, 0);
            int animationSequence = 1;

            // Add animations for each rectangle
            for (int i = 0; i < rectangles.Count; i++)
            {
                RectF rect = InflateAndClamp(rectangles[i], 10, slideWidth, slideHeight);
                CameraPose targetPose = PoseForRect(rect, slideWidth, slideHeight);

                if (i == 0)
                {
                    // First rectangle: zoom + pan on click
                    AddZoomAndPan(newSlide, pictureShape, initialPose, targetPose, 
                        PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick, ref animationSequence, 0.5f);
                }
                else
                {
                    // Subsequent rectangles: zoom + pan after previous, with pan starting with previous
                    CameraPose previousPose = PoseForRect(
                        InflateAndClamp(rectangles[i - 1], 10, slideWidth, slideHeight), 
                        slideWidth, slideHeight);
                    
                    AddZoomAndPan(newSlide, pictureShape, previousPose, targetPose, 
                        PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious, ref animationSequence, 0.3f);
                }
            }

            // Final: zoom out + pan back on next click
            RectF lastRect = InflateAndClamp(rectangles[rectangles.Count - 1], 10, slideWidth, slideHeight);
            CameraPose lastPose = PoseForRect(lastRect, slideWidth, slideHeight);
            AddZoomAndPan(newSlide, pictureShape, lastPose, initialPose, 
                PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick, ref animationSequence, 0.5f);

            // Clean up temp file
            File.Delete(tempPath);

            // Select the new slide
            newSlide.Select();
        }

        private static void AddZoomAndPan(PowerPoint.Slide slide, PowerPoint.Shape shape, 
            CameraPose fromPose, CameraPose toPose, PowerPoint.MsoAnimTriggerType trigger, 
            ref int sequence, float holdDelay)
        {
            PowerPoint.Sequence seq = slide.TimeLine.MainSequence;

            // --- Zoom (Grow/Shrink) ---
            float scalePct = (toPose.Scale / fromPose.Scale) * 100f;

            PowerPoint.Effect zoomFx = seq.AddEffect(
                shape,
                PowerPoint.MsoAnimEffect.msoAnimEffectGrowShrink,
                PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone,
                trigger);

            zoomFx.Timing.Duration = 1.0f;
            PowerPoint.AnimationBehavior zoomBeh = zoomFx.Behaviors[1];
            zoomBeh.ScaleEffect.ByX = scalePct;
            zoomBeh.ScaleEffect.ByY = scalePct;

            // --- Pan (Custom Motion Path) ---
            float panX = toPose.X - fromPose.X;
            float panY = toPose.Y - fromPose.Y;

            // (MsoAnimEffect)47 = msoAnimEffectPathCustom
            PowerPoint.Effect panFx = seq.AddEffect(
                shape,
                (PowerPoint.MsoAnimEffect)47,
                PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone,
                PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);

            panFx.Timing.Duration = 1.0f;
            PowerPoint.AnimationBehavior panBeh = panFx.Behaviors[1];
            panBeh.MotionEffect.Path = "M 0 0 L " + panX.ToString("F2") + " " + panY.ToString("F2");

            sequence++;
        }
    }
}
