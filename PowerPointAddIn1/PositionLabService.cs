using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PowerPointAddIn1
{
    public static class PositionLabService
    {
        private const string FeatureName = "Positions Lab";

        private static readonly List<string> _capturedShapeNames = new List<string>();
        private static int _capturedSlideIndex = -1;

        private struct ShapePosition
        {
            public float Left;
            public float Top;

            public ShapePosition(float left, float top)
            {
                Left = left;
                Top = top;
            }
        }

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
                if (app == null || app.ActiveWindow == null)
                {
                    Warn("PowerPoint window is not available.", "No Active Window");
                    return 0;
                }

                PowerPoint.Selection sel = app.ActiveWindow.Selection;
                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    PowerPoint.Slide slide = app.ActiveWindow.View.Slide as PowerPoint.Slide;
                    if (slide != null)
                        _capturedSlideIndex = slide.SlideIndex;

                    foreach (PowerPoint.Shape shape in sel.ShapeRange)
                        _capturedShapeNames.Add(shape.Name);
                }
                else
                {
                    Warn("Please select one or more shapes first.", "No Shape Selection");
                }
            }
            catch (COMException ex)
            {
                Log("CaptureSelection COM error", ex);
                Warn("Unable to capture the current selection. Please try again.", "Selection Error");
            }
            catch (InvalidOperationException ex)
            {
                Log("CaptureSelection invalid operation", ex);
                Warn("Selection is not available in the current view.", "Selection Error");
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
            ExecuteAtomicAlignment(shapes,
                selected =>
                {
                    float minLeft = selected.Min(s => s.Left);
                    foreach (var shape in selected)
                        shape.Left = minLeft;
                },
                "Align Left");
        }

        public static void AlignRight(PowerPoint.Application app)
        {
            var shapes = GetSelectedShapes(app);
            if (shapes.Count < 2) return;

            // Shape closest to the right edge of the slide is the anchor
            ExecuteAtomicAlignment(shapes,
                selected =>
                {
                    float maxRight = selected.Max(s => s.Left + s.Width);
                    foreach (var shape in selected)
                        shape.Left = maxRight - shape.Width;
                },
                "Align Right");
        }

        public static void AlignTop(PowerPoint.Application app)
        {
            var shapes = GetSelectedShapes(app);
            if (shapes.Count < 2) return;

            // Shape closest to the top edge of the slide is the anchor
            ExecuteAtomicAlignment(shapes,
                selected =>
                {
                    float minTop = selected.Min(s => s.Top);
                    foreach (var shape in selected)
                        shape.Top = minTop;
                },
                "Align Top");
        }

        public static void AlignBottom(PowerPoint.Application app)
        {
            var shapes = GetSelectedShapes(app);
            if (shapes.Count < 2) return;

            // Shape closest to the bottom edge of the slide is the anchor
            ExecuteAtomicAlignment(shapes,
                selected =>
                {
                    float maxBottom = selected.Max(s => s.Top + s.Height);
                    foreach (var shape in selected)
                        shape.Top = maxBottom - shape.Height;
                },
                "Align Bottom");
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
            ExecuteAtomicAlignment(shapes,
                selected =>
                {
                    float referenceCenterY = selected[0].Top + selected[0].Height / 2f;
                    for (int i = 1; i < selected.Count; i++)
                        selected[i].Top = referenceCenterY - selected[i].Height / 2f;
                },
                "Align Horizontal");
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
            ExecuteAtomicAlignment(shapes,
                selected =>
                {
                    float referenceCenterX = selected[0].Left + selected[0].Width / 2f;
                    for (int i = 1; i < selected.Count; i++)
                        selected[i].Left = referenceCenterX - selected[i].Width / 2f;
                },
                "Align Vertical");
        }

        /// <summary>
        /// Combines AlignHorizontal and AlignVertical — moves all other shapes
        /// so that their center coincides with the center of the reference shape.
        /// </summary>
        public static void AlignCenter(PowerPoint.Application app)
        {
            var shapes = GetSelectedShapes(app);
            if (shapes.Count < 2) return;

            ExecuteAtomicAlignment(shapes,
                selected =>
                {
                    float referenceCenterX = selected[0].Left + selected[0].Width / 2f;
                    float referenceCenterY = selected[0].Top + selected[0].Height / 2f;
                    for (int i = 1; i < selected.Count; i++)
                    {
                        selected[i].Left = referenceCenterX - selected[i].Width / 2f;
                        selected[i].Top = referenceCenterY - selected[i].Height / 2f;
                    }
                },
                "Align Center");
        }

        public static void Swap(PowerPoint.Application app)
        {
            var shapes = GetSelectedShapes(app);
            if (shapes.Count != 2) return;

            ExecuteAtomicAlignment(shapes,
                selected =>
                {
                    var a = selected[0];
                    var b = selected[1];

                    float tempLeft = a.Left;
                    float tempTop = a.Top;

                    a.Left = b.Left;
                    a.Top = b.Top;

                    b.Left = tempLeft;
                    b.Top = tempTop;
                },
                "Swap");
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

            ExecuteAtomicAlignment(shapes,
                selected =>
                {
                    // Centre of the reference object (origin)
                    float originX = selected[0].Left + selected[0].Width / 2f;
                    float originY = selected[0].Top + selected[0].Height / 2f;

                    // Centre of the distance-setter → defines the radius
                    float setterCenterX = selected[1].Left + selected[1].Width / 2f;
                    float setterCenterY = selected[1].Top + selected[1].Height / 2f;

                    double dx = setterCenterX - originX;
                    double dy = setterCenterY - originY;
                    double radius = Math.Sqrt(dx * dx + dy * dy);

                    // If the distance-setter is sitting on the origin, use a safe default
                    if (radius < 1.0) radius = 80.0;

                    // Move each shape-to-align (index 2 and beyond)
                    for (int i = 2; i < selected.Count; i++)
                    {
                        float shapeCenterX = selected[i].Left + selected[i].Width / 2f;
                        float shapeCenterY = selected[i].Top + selected[i].Height / 2f;

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

                        selected[i].Left = newCenterX - selected[i].Width / 2f;
                        selected[i].Top = newCenterY - selected[i].Height / 2f;
                    }
                },
                "Align Radially");
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
                if (app == null || app.ActiveWindow == null)
                {
                    Warn("PowerPoint window is not available.", "No Active Window");
                    return result;
                }

                PowerPoint.Selection sel = app.ActiveWindow.Selection;
                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    foreach (PowerPoint.Shape shape in sel.ShapeRange)
                        result.Add(shape);
                }
            }
            catch (COMException ex)
            {
                Log("GetSelectedShapes COM error", ex);
                Warn("Unable to read selected shapes.", "Selection Error");
            }
            catch (InvalidOperationException ex)
            {
                Log("GetSelectedShapes invalid operation", ex);
                Warn("Selection is not available in the current view.", "Selection Error");
            }
            return result;
        }

        private static List<PowerPoint.Shape> ResolveCapturedShapes(PowerPoint.Application app)
        {
            var result = new List<PowerPoint.Shape>();
            try
            {
                if (app == null || app.ActiveWindow == null)
                {
                    Warn("PowerPoint window is not available.", "No Active Window");
                    return result;
                }

                PowerPoint.Slide slide = app.ActiveWindow.View.Slide as PowerPoint.Slide;
                if (slide == null) return result;

                // Only apply on the same slide where shapes were captured
                if (_capturedSlideIndex > 0 && slide.SlideIndex != _capturedSlideIndex)
                {
                    Warn("Captured shapes belong to a different slide. Switch back or clear selection.",
                        "Captured Selection Mismatch");
                    return result;
                }

                // Build a lookup from the slide shapes
                var lookup = new Dictionary<string, PowerPoint.Shape>();
                foreach (PowerPoint.Shape shape in slide.Shapes)
                    lookup[shape.Name] = shape;

                // Return shapes in the original selection order
                foreach (string name in _capturedShapeNames)
                {
                    if (lookup.TryGetValue(name, out PowerPoint.Shape found))
                    {
                        result.Add(found);
                    }
                    else
                    {
                        Log("ResolveCapturedShapes missing shape", null, name);
                    }
                }

                if (result.Count == 0 && _capturedShapeNames.Count > 0)
                {
                    Warn("Captured shapes could not be resolved. They may have been deleted or renamed.",
                        "Captured Selection Missing");
                }
            }
            catch (COMException ex)
            {
                Log("ResolveCapturedShapes COM error", ex);
                Warn("Unable to resolve captured shapes.", "Selection Error");
            }
            catch (InvalidOperationException ex)
            {
                Log("ResolveCapturedShapes invalid operation", ex);
                Warn("Slide context is not valid in the current view.", "Selection Error");
            }
            return result;
        }

        public static IReadOnlyList<string> GetCapturedNames()
        {
            return _capturedShapeNames;
        }

        private static void ExecuteAtomicAlignment(
            List<PowerPoint.Shape> shapes,
            Action<List<PowerPoint.Shape>> applyChanges,
            string operationName)
        {
            var snapshot = new Dictionary<int, ShapePosition>();

            try
            {
                foreach (var shape in shapes)
                {
                    snapshot[shape.Id] = new ShapePosition(shape.Left, shape.Top);
                }

                applyChanges(shapes);
            }
            catch (COMException ex)
            {
                Log(operationName + " COM error", ex);
                TryRestorePositions(shapes, snapshot, operationName);
                Warn(operationName + " could not be completed. Original positions were restored.",
                    FeatureName);
            }
            catch (InvalidOperationException ex)
            {
                Log(operationName + " invalid operation", ex);
                TryRestorePositions(shapes, snapshot, operationName);
                Warn(operationName + " is not available in the current context.", FeatureName);
            }
        }

        private static void TryRestorePositions(
            List<PowerPoint.Shape> shapes,
            Dictionary<int, ShapePosition> snapshot,
            string operationName)
        {
            foreach (var shape in shapes)
            {
                try
                {
                    if (snapshot.TryGetValue(shape.Id, out var original))
                    {
                        shape.Left = original.Left;
                        shape.Top = original.Top;
                    }
                }
                catch (COMException ex)
                {
                    Log(operationName + " rollback COM error", ex, shape.Name);
                    // Best-effort rollback only. If a shape is deleted/invalid during rollback,
                    // we continue restoring remaining shapes to minimize inconsistent state.
                }
            }
        }

        private static void Warn(string message, string title)
        {
            MessageBox.Show(message, title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private static void Log(string message, Exception ex, string context = null)
        {
            string extra = string.IsNullOrEmpty(context) ? string.Empty : (" | Context: " + context);
            string error = ex == null ? string.Empty : (" | " + ex.GetType().Name + ": " + ex.Message);
            System.Diagnostics.Debug.WriteLine("[" + FeatureName + "] " + message + extra + error);
        }
    }
}