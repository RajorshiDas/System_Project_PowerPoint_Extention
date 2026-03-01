using System;
using System.Collections.Generic;
using System.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PowerPointAddIn1
{
    public class ZoomLab
    {
        public static void ZoomToArea(PowerPoint.Presentation presentation, PowerPoint.Slide slide, List<PowerPoint.Shape> rectangles)
        {
            if (rectangles == null || rectangles.Count == 0)
            {
                throw new ArgumentException("No rectangles selected");
            }

            // Sort rectangles by their order (left to right, top to bottom)
            var sortedRects = rectangles.OrderBy(r => r.Top).ThenBy(r => r.Left).ToList();

            // Get slide dimensions
            float slideWidth = presentation.PageSetup.SlideWidth;
            float slideHeight = presentation.PageSetup.SlideHeight;

            // Create animation slide after current slide
            PowerPoint.Slide animSlide = presentation.Slides.Add(slide.SlideIndex + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            
            // Copy all content from original slide
            foreach (PowerPoint.Shape shape in slide.Shapes)
            {
                shape.Copy();
                animSlide.Shapes.Paste();
            }

            // Delete the rectangles from the animation slide
            for (int i = animSlide.Shapes.Count; i >= 1; i--)
            {
                PowerPoint.Shape animShape = animSlide.Shapes[i];
                foreach (var rect in sortedRects)
                {
                    if (Math.Abs(animShape.Left - rect.Left) < 1 && 
                        Math.Abs(animShape.Top - rect.Top) < 1 &&
                        Math.Abs(animShape.Width - rect.Width) < 1)
                    {
                        animShape.Delete();
                        break;
                    }
                }
            }

            // Add zoom animations for each rectangle
            for (int i = 0; i < sortedRects.Count; i++)
            {
                var rect = sortedRects[i];
                AddZoomToRectangle(animSlide, rect, slideWidth, slideHeight, i);
            }

            // Add final zoom back animation
            AddZoomBackAnimation(animSlide, sortedRects.Count);
        }

        private static void AddZoomToRectangle(PowerPoint.Slide slide, PowerPoint.Shape rectangle, 
            float slideWidth, float slideHeight, int sequenceIndex)
        {
            // Calculate zoom factor
            float zoomX = slideWidth / rectangle.Width;
            float zoomY = slideHeight / rectangle.Height;
            float zoom = Math.Min(zoomX, zoomY) * 100; // Convert to percentage

            // Calculate the offset needed to center the rectangle
            float centerX = rectangle.Left + (rectangle.Width / 2);
            float centerY = rectangle.Top + (rectangle.Height / 2);
            float slideCenterX = slideWidth / 2;
            float slideCenterY = slideHeight / 2;
            float offsetX = slideCenterX - centerX;
            float offsetY = slideCenterY - centerY;

            // Get animation sequence
            var timeline = slide.TimeLine;
            var mainSeq = timeline.MainSequence;

            // Animate all shapes on the slide
            foreach (PowerPoint.Shape shape in slide.Shapes)
            {
                // Add emphasis grow/shrink effect
                var effect = mainSeq.AddEffect(shape, PowerPoint.MsoAnimEffect.msoAnimEffectGrowShrink,
                    PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone,
                    sequenceIndex == 0 ? PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick : 
                                        PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                
                effect.Timing.Duration = 1.0f;
                effect.EffectParameters.Size = zoom;
                
                // Add motion path for panning
                var pathEffect = mainSeq.AddEffect(shape, PowerPoint.MsoAnimEffect.msoAnimEffectPathDown,
                    PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone,
                    PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                pathEffect.Timing.Duration = 1.0f;
                
                // Set custom path
                try
                {
                    var motionEffect = pathEffect.Behaviors[1] as PowerPoint.AnimationBehavior;
                    if (motionEffect != null && motionEffect.MotionEffect != null)
                    {
                        motionEffect.MotionEffect.Path = $"M 0 0 L {offsetX} {offsetY}";
                    }
                }
                catch { }
            }
        }

        private static void AddZoomBackAnimation(PowerPoint.Slide slide, int afterSequence)
        {
            var timeline = slide.TimeLine;
            var mainSeq = timeline.MainSequence;

            foreach (PowerPoint.Shape shape in slide.Shapes)
            {
                // Zoom back to original
                var effect = mainSeq.AddEffect(shape, PowerPoint.MsoAnimEffect.msoAnimEffectGrowShrink,
                    PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone,
                    PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                
                effect.Timing.Duration = 1.0f;
                effect.EffectParameters.Size = 100; // Back to 100%
            }
        }

        public static void DrillDown(PowerPoint.Presentation presentation, PowerPoint.Slide bigPictureSlide, 
            PowerPoint.Shape rectangle)
        {
            // Get next slide as drill down target
            if (bigPictureSlide.SlideIndex >= presentation.Slides.Count)
            {
                throw new InvalidOperationException("No slide available to drill down into. Add a slide after this one.");
            }

            PowerPoint.Slide drillDownSlide = presentation.Slides[bigPictureSlide.SlideIndex + 1];
            
            // Export drill down slide as image
            string tempPath = System.IO.Path.GetTempPath();
            string imagePath = System.IO.Path.Combine(tempPath, $"drilldown_{Guid.NewGuid()}.png");
            
            drillDownSlide.Export(imagePath, "PNG");

            // Replace rectangle with image
            float left = rectangle.Left;
            float top = rectangle.Top;
            float width = rectangle.Width;
            float height = rectangle.Height;
            
            rectangle.Delete();
            
            var picture = bigPictureSlide.Shapes.AddPicture(imagePath, 
                Office.MsoTriState.msoFalse, 
                Office.MsoTriState.msoTrue, 
                left, top, width, height);
            picture.Tags.Add("DrillDown", "True");
            picture.Tags.Add("TargetSlide", drillDownSlide.SlideIndex.ToString());

            // Create animation slide with zoom effect
            CreateDrillDownAnimation(presentation, bigPictureSlide, drillDownSlide);

            try
            {
                System.IO.File.Delete(imagePath);
            }
            catch { }
        }

        private static void CreateDrillDownAnimation(PowerPoint.Presentation presentation, 
            PowerPoint.Slide sourceSlide, PowerPoint.Slide targetSlide)
        {
            // Create animation slide between source and target
            PowerPoint.Slide animSlide = presentation.Slides.Add(sourceSlide.SlideIndex + 1, 
                PowerPoint.PpSlideLayout.ppLayoutBlank);
            
            // Copy source slide content
            foreach (PowerPoint.Shape shape in sourceSlide.Shapes)
            {
                shape.Copy();
                animSlide.Shapes.Paste();
            }

            // Find the drill down picture
            PowerPoint.Shape pictureShape = null;
            foreach (PowerPoint.Shape shape in animSlide.Shapes)
            {
                if (shape.Tags["DrillDown"] == "True")
                {
                    pictureShape = shape;
                    break;
                }
            }

            if (pictureShape != null)
            {
                // Add zoom animation to the picture
                var timeline = animSlide.TimeLine;
                var mainSeq = timeline.MainSequence;
                
                // Zoom in effect - picture grows to fill slide
                var effect = mainSeq.AddEffect(pictureShape, 
                    PowerPoint.MsoAnimEffect.msoAnimEffectGrowShrink,
                    PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone,
                    PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                effect.Timing.Duration = 1.5f;
                
                float slideWidth = presentation.PageSetup.SlideWidth;
                float slideHeight = presentation.PageSetup.SlideHeight;
                float scaleX = (slideWidth / pictureShape.Width) * 100;
                float scaleY = (slideHeight / pictureShape.Height) * 100;
                float scale = Math.Max(scaleX, scaleY);
                
                effect.EffectParameters.Size = scale;
            }
        }

        public static void StepBack(PowerPoint.Presentation presentation, PowerPoint.Slide currentSlide)
        {
            // Find the previous "big picture" slide
            if (currentSlide.SlideIndex <= 1)
            {
                throw new InvalidOperationException("No previous slide to step back to.");
            }

            PowerPoint.Slide previousSlide = presentation.Slides[currentSlide.SlideIndex - 1];

            // Create zoom out animation slide
            PowerPoint.Slide animSlide = presentation.Slides.Add(currentSlide.SlideIndex + 1, 
                PowerPoint.PpSlideLayout.ppLayoutBlank);
            
            // Copy current slide
            foreach (PowerPoint.Shape shape in currentSlide.Shapes)
            {
                shape.Copy();
                animSlide.Shapes.Paste();
            }

            // Add zoom out effect to all shapes
            var timeline = animSlide.TimeLine;
            var mainSeq = timeline.MainSequence;
            
            foreach (PowerPoint.Shape shape in animSlide.Shapes)
            {
                var effect = mainSeq.AddEffect(shape, 
                    PowerPoint.MsoAnimEffect.msoAnimEffectGrowShrink,
                    PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone,
                    PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
                effect.Timing.Duration = 1.5f;
                effect.EffectParameters.Size = 50; // Shrink to 50%
            }

            // Export previous slide as background
            string tempPath = System.IO.Path.GetTempPath();
            string imagePath = System.IO.Path.Combine(tempPath, $"stepback_{Guid.NewGuid()}.png");
            previousSlide.Export(imagePath, "PNG");

            // Add previous slide as background image
            var bgPicture = animSlide.Shapes.AddPicture(imagePath,
                Office.MsoTriState.msoFalse,
                Office.MsoTriState.msoTrue,
                0, 0, presentation.PageSetup.SlideWidth, presentation.PageSetup.SlideHeight);
            bgPicture.ZOrder(Office.MsoZOrderCmd.msoSendToBack);

            try
            {
                System.IO.File.Delete(imagePath);
            }
            catch { }
        }
    }
}
