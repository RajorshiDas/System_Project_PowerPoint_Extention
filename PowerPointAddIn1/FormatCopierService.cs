using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1
{
    public class FormatCopierService
    {
        private bool _hasSource;
        private int _sourceSlideIndex = -1;
        private string _sourceShapeName;
        private float _sourceWidth;
        private float _sourceHeight;
        private float _sourceLeft;
        private float _sourceTop;

        public string CaptureSourceShape(PowerPoint.Application app)
        {
            try
            {
                if (app == null || app.Presentations.Count == 0)
                    return "Please open a presentation first.";

                var slide = GetActiveSlide(app);
                if (slide == null)
                    return "Please open a slide in Normal view.";

                var selectedShapes = GetSelectedShapes(app);
                if (selectedShapes.Count == 0)
                    return "Please select a shape first.";

                var source = selectedShapes[0];
                _sourceSlideIndex = slide.SlideIndex;
                _sourceShapeName = source.Name;
                _sourceWidth = source.Width;
                _sourceHeight = source.Height;
                _sourceLeft = source.Left;
                _sourceTop = source.Top;
                _hasSource = true;

                return "Source shape captured successfully.";
            }
            catch (COMException)
            {
                return "Unable to capture source shape from current selection.";
            }
            catch (Exception)
            {
                return "Unexpected error while capturing source shape.";
            }
        }

        public string ApplyCopy(PowerPoint.Application app, bool copySize, bool copyPosition, bool copyAnimation)
        {
            try
            {
                if (!_hasSource)
                    return "No source shape captured. Click Capture Selected Shape first.";

                if (app == null || app.Presentations.Count == 0)
                    return "Please open a presentation first.";

                if (!copySize && !copyPosition && !copyAnimation)
                    return "Select at least one copy option.";

                var presentation = app.ActivePresentation;
                if (_sourceSlideIndex < 1 || _sourceSlideIndex > presentation.Slides.Count)
                    return "Source slide no longer exists.";

                var sourceSlide = presentation.Slides[_sourceSlideIndex];
                var sourceShape = FindShapeOnSlide(sourceSlide, _sourceShapeName);
                if (sourceShape == null)
                    return "Source shape no longer exists. Capture it again.";

                var targetShapes = GetSelectedShapes(app);
                if (targetShapes.Count == 0)
                    return "Please select one or more target shapes.";

                int appliedCount = 0;
                foreach (var target in targetShapes)
                {
                    if (target == null) continue;
                    if (IsSameShape(target, sourceShape)) continue;

                    try
                    {
                        if (copySize)
                        {
                            target.Width = _sourceWidth;
                            target.Height = _sourceHeight;
                        }

                        if (copyPosition)
                        {
                            target.Left = _sourceLeft;
                            target.Top = _sourceTop;
                        }

                        if (copyAnimation)
                        {
                            CopyAnimations(sourceShape, target);
                        }

                        appliedCount++;
                    }
                    catch
                    {
                        // Skip target shape and continue.
                    }
                }

                return appliedCount > 0
                    ? string.Format("Copied to {0} shape(s).", appliedCount)
                    : "No target shapes were updated.";
            }
            catch (COMException)
            {
                return "Unable to apply copy. Please verify current selection.";
            }
            catch (Exception)
            {
                return "Unexpected error while applying copy.";
            }
        }

        public string ClearSource()
        {
            _hasSource = false;
            _sourceSlideIndex = -1;
            _sourceShapeName = null;
            _sourceWidth = 0f;
            _sourceHeight = 0f;
            _sourceLeft = 0f;
            _sourceTop = 0f;

            return "Source shape cleared.";
        }

        public bool HasSource()
        {
            return _hasSource;
        }

        private static PowerPoint.Slide GetActiveSlide(PowerPoint.Application app)
        {
            try
            {
                if (app == null || app.ActiveWindow == null || app.ActiveWindow.View == null)
                    return null;

                return app.ActiveWindow.View.Slide as PowerPoint.Slide;
            }
            catch
            {
                return null;
            }
        }

        private static List<PowerPoint.Shape> GetSelectedShapes(PowerPoint.Application app)
        {
            var result = new List<PowerPoint.Shape>();

            try
            {
                if (app == null || app.ActiveWindow == null || app.ActiveWindow.Selection == null)
                    return result;

                var selection = app.ActiveWindow.Selection;
                if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes &&
                    selection.Type != PowerPoint.PpSelectionType.ppSelectionText)
                {
                    return result;
                }

                var range = selection.ShapeRange;
                if (range == null) return result;

                for (int i = 1; i <= range.Count; i++)
                {
                    if (range[i] != null)
                        result.Add(range[i]);
                }
            }
            catch
            {
                // Return empty list on errors.
            }

            return result;
        }

        private static PowerPoint.Shape FindShapeOnSlide(PowerPoint.Slide slide, string shapeName)
        {
            try
            {
                if (slide == null || string.IsNullOrEmpty(shapeName)) return null;

                for (int i = 1; i <= slide.Shapes.Count; i++)
                {
                    var shape = slide.Shapes[i];
                    if (string.Equals(shape.Name, shapeName, StringComparison.OrdinalIgnoreCase))
                        return shape;
                }
            }
            catch
            {
                // Ignore lookup issues.
            }

            return null;
        }

        private static bool IsSameShape(PowerPoint.Shape a, PowerPoint.Shape b)
        {
            if (a == null || b == null) return false;

            try
            {
                return a.Id == b.Id;
            }
            catch
            {
                return false;
            }
        }

        private static void CopyAnimations(PowerPoint.Shape sourceShape, PowerPoint.Shape targetShape)
        {
            try
            {
                var sourceSlide = sourceShape.Parent as PowerPoint.Slide;
                var targetSlide = targetShape.Parent as PowerPoint.Slide;
                if (sourceSlide == null || targetSlide == null) return;

                var sourceSequence = sourceSlide.TimeLine.MainSequence;
                var targetSequence = targetSlide.TimeLine.MainSequence;
                if (sourceSequence == null || targetSequence == null) return;

                for (int i = targetSequence.Count; i >= 1; i--)
                {
                    try
                    {
                        var effect = targetSequence[i];
                        if (effect != null && effect.Shape != null && effect.Shape.Id == targetShape.Id)
                            effect.Delete();
                    }
                    catch
                    {
                        // Ignore effect delete failure.
                    }
                }

                for (int i = 1; i <= sourceSequence.Count; i++)
                {
                    try
                    {
                        var effect = sourceSequence[i];
                        if (effect == null || effect.Shape == null || effect.Shape.Id != sourceShape.Id) continue;

                        targetSequence.AddEffect(
                            targetShape,
                            effect.EffectType,
                            PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone,
                            effect.Timing.TriggerType,
                            targetSequence.Count + 1);
                    }
                    catch
                    {
                        // Ignore unsupported effect copy.
                    }
                }
            }
            catch
            {
                // Ignore animation copy failures.
            }
        }
    }
}
