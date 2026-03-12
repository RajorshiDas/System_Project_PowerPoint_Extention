using System;
using System.Collections.Generic;
using System.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1
{
    public static class PositionLabService
    {
        private static readonly List<string> _capturedShapeNames = new List<string>();
        private static int _capturedSlideIndex = -1;

        public static bool HasCapturedSelection
        {
            get { return _capturedShapeNames.Count > 0; }
        }

        public static int CapturedCount
        {
            get { return _capturedShapeNames.Count; }
        }

        public static int CaptureSelection(PowerPoint.Application app)
        {
            _capturedShapeNames.Clear();
            _capturedSlideIndex = -1;

            try
            {
                PowerPoint.Selection sel = app.ActiveWindow.Selection;
                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    PowerPoint.Slide slide = app.ActiveWindow.View.Slide as PowerPoint.Slide;
                    if (slide != null)
                        _capturedSlideIndex = slide.SlideIndex;

                    foreach (PowerPoint.Shape shape in sel.ShapeRange)
                        _capturedShapeNames.Add(shape.Name);
                }
            }
            catch
            {
                // No valid selection
            }
            return _capturedShapeNames.Count;
        }

        public static void ClearSelection()
        {
            _capturedShapeNames.Clear();
            _capturedSlideIndex = -1;
        }

        public static void AlignLeft(PowerPoint.Application app)
        {
            var shapes = GetSelectedShapes(app);
            if (shapes.Count < 1) return;

            float minLeft = shapes.Min(s => s.Left);
            foreach (var shape in shapes)
                shape.Left = minLeft;
        }

        public static void AlignRight(PowerPoint.Application app)
        {
            var shapes = GetSelectedShapes(app);
            if (shapes.Count < 1) return;

            float maxRight = shapes.Max(s => s.Left + s.Width);
            foreach (var shape in shapes)
                shape.Left = maxRight - shape.Width;
        }

        public static void AlignTop(PowerPoint.Application app)
        {
            var shapes = GetSelectedShapes(app);
            if (shapes.Count < 1) return;

            float minTop = shapes.Min(s => s.Top);
            foreach (var shape in shapes)
                shape.Top = minTop;
        }

        public static void AlignBottom(PowerPoint.Application app)
        {
            var shapes = GetSelectedShapes(app);
            if (shapes.Count < 1) return;

            float maxBottom = shapes.Max(s => s.Top + s.Height);
            foreach (var shape in shapes)
                shape.Top = maxBottom - shape.Height;
        }

        public static void AlignCenter(PowerPoint.Application app)
        {
            var shapes = GetSelectedShapes(app);
            if (shapes.Count < 1) return;

            float minLeft = shapes.Min(s => s.Left);
            float maxRight = shapes.Max(s => s.Left + s.Width);
            float centerX = (minLeft + maxRight) / 2f;

            float minTop = shapes.Min(s => s.Top);
            float maxBottom = shapes.Max(s => s.Top + s.Height);
            float centerY = (minTop + maxBottom) / 2f;

            foreach (var shape in shapes)
            {
                shape.Left = centerX - shape.Width / 2f;
                shape.Top = centerY - shape.Height / 2f;
            }
        }

        public static void DistributeHorizontal(PowerPoint.Application app)
        {
            var shapes = GetSelectedShapes(app);
            if (shapes.Count < 3) return;

            var sorted = shapes.OrderBy(s => s.Left).ToList();
            float totalWidth = sorted.Sum(s => s.Width);
            float minLeft = sorted.First().Left;
            float maxRight = sorted.Last().Left + sorted.Last().Width;
            float spacing = (maxRight - minLeft - totalWidth) / (sorted.Count - 1);

            float currentLeft = minLeft;
            foreach (var shape in sorted)
            {
                shape.Left = currentLeft;
                currentLeft += shape.Width + spacing;
            }
        }

        public static void DistributeVertical(PowerPoint.Application app)
        {
            var shapes = GetSelectedShapes(app);
            if (shapes.Count < 3) return;

            var sorted = shapes.OrderBy(s => s.Top).ToList();
            float totalHeight = sorted.Sum(s => s.Height);
            float minTop = sorted.First().Top;
            float maxBottom = sorted.Last().Top + sorted.Last().Height;
            float spacing = (maxBottom - minTop - totalHeight) / (sorted.Count - 1);

            float currentTop = minTop;
            foreach (var shape in sorted)
            {
                shape.Top = currentTop;
                currentTop += shape.Height + spacing;
            }
        }

        public static void Swap(PowerPoint.Application app)
        {
            var shapes = GetSelectedShapes(app);
            if (shapes.Count != 2) return;

            var a = shapes[0];
            var b = shapes[1];

            float tempLeft = a.Left;
            float tempTop = a.Top;

            a.Left = b.Left;
            a.Top = b.Top;

            b.Left = tempLeft;
            b.Top = tempTop;
        }

        public static void AlignRadially(PowerPoint.Application app)
        {
            var shapes = GetSelectedShapes(app);
            if (shapes.Count < 2) return;

            // Find the bounding-box center of all selected shapes
            float minLeft = shapes.Min(s => s.Left);
            float maxRight = shapes.Max(s => s.Left + s.Width);
            float minTop = shapes.Min(s => s.Top);
            float maxBottom = shapes.Max(s => s.Top + s.Height);
            float centerX = (minLeft + maxRight) / 2f;
            float centerY = (minTop + maxBottom) / 2f;

            // Use half the smaller bounding-box dimension as the radius
            float radius = Math.Min(maxRight - minLeft, maxBottom - minTop) / 2f;
            if (radius < 20f) radius = 80f;

            double angleStep = 2 * Math.PI / shapes.Count;

            for (int i = 0; i < shapes.Count; i++)
            {
                double angle = -Math.PI / 2 + i * angleStep; // start at top
                shapes[i].Left = centerX + (float)(radius * Math.Cos(angle)) - shapes[i].Width / 2f;
                shapes[i].Top  = centerY + (float)(radius * Math.Sin(angle)) - shapes[i].Height / 2f;
            }
        }

        private static List<PowerPoint.Shape> GetSelectedShapes(PowerPoint.Application app)
        {
            // If shapes were captured, resolve them from the slide
            if (_capturedShapeNames.Count > 0)
            {
                return ResolveCapturedShapes(app);
            }

            // Otherwise use the live PowerPoint selection
            var result = new List<PowerPoint.Shape>();
            try
            {
                PowerPoint.Selection sel = app.ActiveWindow.Selection;
                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    foreach (PowerPoint.Shape shape in sel.ShapeRange)
                        result.Add(shape);
                }
            }
            catch
            {
                // No valid selection
            }
            return result;
        }

        private static List<PowerPoint.Shape> ResolveCapturedShapes(PowerPoint.Application app)
        {
            var result = new List<PowerPoint.Shape>();
            try
            {
                PowerPoint.Slide slide = app.ActiveWindow.View.Slide as PowerPoint.Slide;
                if (slide == null) return result;

                // Only apply on the same slide where shapes were captured
                if (_capturedSlideIndex > 0 && slide.SlideIndex != _capturedSlideIndex)
                    return result;

                var nameSet = new HashSet<string>(_capturedShapeNames);
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    if (nameSet.Contains(shape.Name))
                        result.Add(shape);
                }
            }
            catch
            {
                // Slide or shapes no longer valid
            }
            return result;
        }
    }
}
