using System;
using System.Collections.Generic;
using System.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

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
            if (shapes.Count < 2) return;

            // Shape closest to the left edge of the slide is the anchor
            float minLeft = shapes.Min(s => s.Left);
            foreach (var shape in shapes)
                shape.Left = minLeft;
        }

        public static void AlignRight(PowerPoint.Application app)
        {
            var shapes = GetSelectedShapes(app);
            if (shapes.Count < 2) return;

            // Shape closest to the right edge of the slide is the anchor
            float maxRight = shapes.Max(s => s.Left + s.Width);
            foreach (var shape in shapes)
                shape.Left = maxRight - shape.Width;
        }

        public static void AlignTop(PowerPoint.Application app)
        {
            var shapes = GetSelectedShapes(app);
            if (shapes.Count < 2) return;

            // Shape closest to the top edge of the slide is the anchor
            float minTop = shapes.Min(s => s.Top);
            foreach (var shape in shapes)
                shape.Top = minTop;
        }

        public static void AlignBottom(PowerPoint.Application app)
        {
            var shapes = GetSelectedShapes(app);
            if (shapes.Count < 2) return;

            // Shape closest to the bottom edge of the slide is the anchor
            float maxBottom = shapes.Max(s => s.Top + s.Height);
            foreach (var shape in shapes)
                shape.Top = maxBottom - shape.Height;
        }

        /// <summary>
        /// Aligns all other shapes to the horizontal center (Y midpoint) of the
        /// first / reference shape, leaving their X positions unchanged.
        /// </summary>
        public static void AlignHorizontal(PowerPoint.Application app)
        {
            var shapes = GetSelectedShapes(app);
            if (shapes.Count < 2) return;

            // shapes[0] is the reference (first selected / first captured)
            float referenceCenterY = shapes[0].Top + shapes[0].Height / 2f;
            for (int i = 1; i < shapes.Count; i++)
                shapes[i].Top = referenceCenterY - shapes[i].Height / 2f;
        }

        /// <summary>
        /// Aligns all other shapes to the vertical center (X midpoint) of the
        /// first / reference shape, leaving their Y positions unchanged.
        /// </summary>
        public static void AlignVertical(PowerPoint.Application app)
        {
            var shapes = GetSelectedShapes(app);
            if (shapes.Count < 2) return;

            // shapes[0] is the reference (first selected / first captured)
            float referenceCenterX = shapes[0].Left + shapes[0].Width / 2f;
            for (int i = 1; i < shapes.Count; i++)
                shapes[i].Left = referenceCenterX - shapes[i].Width / 2f;
        }

        /// <summary>
        /// Combines AlignHorizontal and AlignVertical — moves all other shapes
        /// so that their center coincides with the center of the reference shape.
        /// </summary>
        public static void AlignCenter(PowerPoint.Application app)
        {
            var shapes = GetSelectedShapes(app);
            if (shapes.Count < 2) return;

            float referenceCenterX = shapes[0].Left + shapes[0].Width / 2f;
            float referenceCenterY = shapes[0].Top + shapes[0].Height / 2f;
            for (int i = 1; i < shapes.Count; i++)
            {
                shapes[i].Left = referenceCenterX - shapes[i].Width / 2f;
                shapes[i].Top = referenceCenterY - shapes[i].Height / 2f;
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

        /// <summary>
        /// Aligns shapes[2..n] so that each sits at the same radial distance from
        /// the reference object (shapes[0]) as the distance-setter (shapes[1]).
        ///
        /// Selection order:
        ///   shapes[0]  – reference object  (the origin / centre)
        ///   shapes[1]  – distance-setter   (defines the radius; its position is NOT changed)
        ///   shapes[2+] – objects to align  (each keeps its current angle from the origin
        ///                                   but is moved to exactly <radius> away from it)
        ///
        /// Requires at least 3 selected shapes.
        /// If a shape to be aligned sits exactly on the origin its current angle cannot
        /// be determined; it is placed directly above the origin as a safe fallback.
        /// </summary>
        public static void AlignRadially(PowerPoint.Application app)
        {
            var shapes = GetSelectedShapes(app);
            if (shapes.Count < 3) return;

            // Centre of the reference object (origin)
            float originX = shapes[0].Left + shapes[0].Width / 2f;
            float originY = shapes[0].Top + shapes[0].Height / 2f;

            // Centre of the distance-setter → defines the radius
            float setterCenterX = shapes[1].Left + shapes[1].Width / 2f;
            float setterCenterY = shapes[1].Top + shapes[1].Height / 2f;

            double dx = setterCenterX - originX;
            double dy = setterCenterY - originY;
            double radius = Math.Sqrt(dx * dx + dy * dy);

            // If the distance-setter is sitting on the origin, use a safe default
            if (radius < 1.0) radius = 80.0;

            // Move each shape-to-align (index 2 and beyond)
            for (int i = 2; i < shapes.Count; i++)
            {
                float shapeCenterX = shapes[i].Left + shapes[i].Width / 2f;
                float shapeCenterY = shapes[i].Top + shapes[i].Height / 2f;

                double adx = shapeCenterX - originX;
                double ady = shapeCenterY - originY;
                double currentDist = Math.Sqrt(adx * adx + ady * ady);

                double angle;
                if (currentDist < 1.0)
                {
                    // Shape is on the origin — place it directly above as a fallback
                    angle = -Math.PI / 2.0;
                }
                else
                {
                    angle = Math.Atan2(ady, adx);
                }

                // Position the shape so its centre is exactly <radius> from the origin
                float newCenterX = originX + (float)(radius * Math.Cos(angle));
                float newCenterY = originY + (float)(radius * Math.Sin(angle));

                shapes[i].Left = newCenterX - shapes[i].Width / 2f;
                shapes[i].Top = newCenterY - shapes[i].Height / 2f;
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

                // Build a lookup from the slide shapes
                var lookup = new Dictionary<string, PowerPoint.Shape>();
                foreach (PowerPoint.Shape shape in slide.Shapes)
                    lookup[shape.Name] = shape;

                // Return shapes in the original selection order
                foreach (string name in _capturedShapeNames)
                {
                    if (lookup.TryGetValue(name, out PowerPoint.Shape found))
                        result.Add(found);
                }
            }
            catch
            {
                // Slide or shapes no longer valid
            }
            return result;
        }

        public static IReadOnlyList<string> GetCapturedNames()
        {
            return _capturedShapeNames;
        }
    }
}