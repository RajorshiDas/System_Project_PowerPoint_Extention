using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1
{
    public enum ReferenceMode
    {
        FirstSelected,
        OutermostObject
    }

    public class ResizeLabService
    {
        public List<PowerPoint.Shape> GetSelectedShapes(PowerPoint.Application app)
        {
            var shapes = new List<PowerPoint.Shape>();

            if (app == null)
            {
                MessageBox.Show("PowerPoint application is not available.", "Resize Lab",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return shapes;
            }

            try
            {
                var window = app.ActiveWindow;
                if (window == null || window.Selection == null)
                {
                    MessageBox.Show("Please select one or more shapes first.", "Resize Lab",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return shapes;
                }

                var selection = window.Selection;
                if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    MessageBox.Show("Please select one or more shapes first.", "Resize Lab",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return shapes;
                }

                var shapeRange = selection.ShapeRange;
                if (shapeRange == null || shapeRange.Count == 0)
                {
                    MessageBox.Show("Please select one or more shapes first.", "Resize Lab",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return shapes;
                }

                for (int i = 1; i <= shapeRange.Count; i++)
                {
                    var shape = shapeRange[i];
                    if (shape != null)
                    {
                        shapes.Add(shape);
                    }
                }

                if (shapes.Count == 0)
                {
                    MessageBox.Show("Please select one or more shapes first.", "Resize Lab",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (COMException)
            {
                MessageBox.Show("Unable to read the current selection. Please select one or more shapes and try again.",
                    "Resize Lab", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception)
            {
                MessageBox.Show("An unexpected error occurred while reading selected shapes.",
                    "Resize Lab", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return shapes;
        }

        public PowerPoint.Slide GetActiveSlide(PowerPoint.Application app)
        {
            if (app == null)
            {
                return null;
            }

            try
            {
                var window = app.ActiveWindow;
                if (window == null || window.View == null)
                {
                    return null;
                }

                return window.View.Slide as PowerPoint.Slide;
            }
            catch (COMException)
            {
                return null;
            }
            catch
            {
                return null;
            }
        }

        public SizeF GetSlideSize(PowerPoint.Application app)
        {
            if (app == null)
            {
                return SizeF.Empty;
            }

            try
            {
                var presentation = app.ActivePresentation;
                if (presentation == null || presentation.PageSetup == null)
                {
                    return SizeF.Empty;
                }

                return new SizeF(presentation.PageSetup.SlideWidth, presentation.PageSetup.SlideHeight);
            }
            catch (COMException)
            {
                return SizeF.Empty;
            }
            catch
            {
                return SizeF.Empty;
            }
        }

        public PowerPoint.Shape GetReferenceShapeForLeft(IList<PowerPoint.Shape> shapes, ReferenceMode referenceMode)
        {
            if (!HasShapes(shapes)) return null;

            if (referenceMode == ReferenceMode.FirstSelected)
            {
                return GetFirstResizableShape(shapes);
            }

            PowerPoint.Shape reference = null;
            foreach (var shape in shapes)
            {
                if (!IsResizableShape(shape)) continue;
                if (reference == null || shape.Left < reference.Left)
                    reference = shape;
            }
            return reference;
        }

        public PowerPoint.Shape GetReferenceShapeForRight(IList<PowerPoint.Shape> shapes, ReferenceMode referenceMode)
        {
            if (!HasShapes(shapes)) return null;

            if (referenceMode == ReferenceMode.FirstSelected)
            {
                return GetFirstResizableShape(shapes);
            }

            PowerPoint.Shape reference = null;
            float referenceRight = float.MinValue;

            foreach (var shape in shapes)
            {
                if (!IsResizableShape(shape)) continue;

                float right = GetRightEdge(shape);
                if (reference == null || right > referenceRight)
                {
                    reference = shape;
                    referenceRight = right;
                }
            }

            return reference;
        }

        public PowerPoint.Shape GetReferenceShapeForTop(IList<PowerPoint.Shape> shapes, ReferenceMode referenceMode)
        {
            if (!HasShapes(shapes)) return null;

            if (referenceMode == ReferenceMode.FirstSelected)
            {
                return GetFirstResizableShape(shapes);
            }

            PowerPoint.Shape reference = null;
            foreach (var shape in shapes)
            {
                if (!IsResizableShape(shape)) continue;
                if (reference == null || shape.Top < reference.Top)
                    reference = shape;
            }
            return reference;
        }

        public PowerPoint.Shape GetReferenceShapeForBottom(IList<PowerPoint.Shape> shapes, ReferenceMode referenceMode)
        {
            if (!HasShapes(shapes)) return null;

            if (referenceMode == ReferenceMode.FirstSelected)
            {
                return GetFirstResizableShape(shapes);
            }

            PowerPoint.Shape reference = null;
            float referenceBottom = float.MinValue;

            foreach (var shape in shapes)
            {
                if (!IsResizableShape(shape)) continue;

                float bottom = GetBottomEdge(shape);
                if (reference == null || bottom > referenceBottom)
                {
                    reference = shape;
                    referenceBottom = bottom;
                }
            }

            return reference;
        }

        public PowerPoint.Shape GetReferenceShapeForSize(IList<PowerPoint.Shape> shapes, ReferenceMode referenceMode)
        {
            if (!HasShapes(shapes)) return null;

            if (referenceMode == ReferenceMode.FirstSelected)
            {
                return GetFirstResizableShape(shapes);
            }

            PowerPoint.Shape reference = null;
            float referenceArea = float.MinValue;

            foreach (var shape in shapes)
            {
                if (!IsResizableShape(shape)) continue;

                float currentArea = shape.Width * shape.Height;
                if (reference == null || currentArea > referenceArea)
                {
                    reference = shape;
                    referenceArea = currentArea;
                }
            }

            return reference;
        }

        public void StretchToLeft(List<PowerPoint.Shape> shapes, ReferenceMode mode)
        {
            if (!HasAtLeastTwoShapes(shapes)) return;

            var reference = GetReferenceShapeForLeft(shapes, mode);
            if (!IsResizableShape(reference)) return;

            float referenceLeft;
            try
            {
                referenceLeft = reference.Left;
            }
            catch
            {
                return;
            }

            foreach (var shape in shapes)
            {
                if (!IsResizableShape(shape) || IsReferenceShape(shape, reference)) continue;

                try
                {
                    float oldRight = GetRightEdge(shape);
                    float newWidth = oldRight - referenceLeft;
                    if (newWidth <= 0f) continue;

                    shape.Left = referenceLeft;
                    shape.Width = newWidth;
                }
                catch (COMException)
                {
                    // Skip COM-failed shape.
                }
                catch
                {
                    // Skip non-fatal shape error.
                }
            }
        }

        public void StretchToRight(List<PowerPoint.Shape> shapes, ReferenceMode mode)
        {
            if (!HasAtLeastTwoShapes(shapes)) return;

            var reference = GetReferenceShapeForRight(shapes, mode);
            if (!IsResizableShape(reference)) return;

            float referenceRight;
            try
            {
                referenceRight = reference.Left + reference.Width;
            }
            catch
            {
                return;
            }

            foreach (var shape in shapes)
            {
                if (!IsResizableShape(shape) || IsReferenceShape(shape, reference)) continue;

                try
                {
                    float newWidth = referenceRight - shape.Left;
                    if (newWidth <= 0f) continue;

                    shape.Width = newWidth;
                }
                catch (COMException)
                {
                    // Skip COM-failed shape.
                }
                catch
                {
                    // Skip non-fatal shape error.
                }
            }
        }

        public void StretchToTop(List<PowerPoint.Shape> shapes, ReferenceMode mode)
        {
            if (!HasAtLeastTwoShapes(shapes)) return;

            var reference = GetReferenceShapeForTop(shapes, mode);
            if (!IsResizableShape(reference)) return;

            float referenceTop;
            try
            {
                referenceTop = reference.Top;
            }
            catch
            {
                return;
            }

            foreach (var shape in shapes)
            {
                if (!IsResizableShape(shape) || IsReferenceShape(shape, reference)) continue;

                try
                {
                    float oldBottom = GetBottomEdge(shape);
                    float newHeight = oldBottom - referenceTop;
                    if (newHeight <= 0f) continue;

                    shape.Top = referenceTop;
                    shape.Height = newHeight;
                }
                catch (COMException)
                {
                    // Skip COM-failed shape.
                }
                catch
                {
                    // Skip non-fatal shape error.
                }
            }
        }

        public void StretchToBottom(List<PowerPoint.Shape> shapes, ReferenceMode mode)
        {
            if (!HasAtLeastTwoShapes(shapes)) return;

            var reference = GetReferenceShapeForBottom(shapes, mode);
            if (!IsResizableShape(reference)) return;

            float referenceBottom;
            try
            {
                referenceBottom = reference.Top + reference.Height;
            }
            catch
            {
                return;
            }

            foreach (var shape in shapes)
            {
                if (!IsResizableShape(shape) || IsReferenceShape(shape, reference)) continue;

                try
                {
                    float newHeight = referenceBottom - shape.Top;
                    if (newHeight <= 0f) continue;

                    shape.Height = newHeight;
                }
                catch (COMException)
                {
                    // Skip COM-failed shape.
                }
                catch
                {
                    // Skip non-fatal shape error.
                }
            }
        }

        public void MatchWidth(List<PowerPoint.Shape> shapes, ReferenceMode mode)
        {
            if (!HasAtLeastTwoShapes(shapes)) return;

            var reference = GetReferenceShapeForSize(shapes, mode);
            if (!IsResizableShape(reference)) return;

            float referenceWidth;
            try
            {
                referenceWidth = reference.Width;
                if (referenceWidth <= 0f) return;
            }
            catch
            {
                return;
            }

            foreach (var shape in shapes)
            {
                if (!IsResizableShape(shape) || IsReferenceShape(shape, reference)) continue;

                try
                {
                    if (shape.Width <= 0f) continue;
                    shape.Width = referenceWidth;
                }
                catch (COMException)
                {
                    // Skip COM-failed shape.
                }
                catch
                {
                    // Skip non-fatal shape error.
                }
            }
        }

        public void MatchHeight(List<PowerPoint.Shape> shapes, ReferenceMode mode)
        {
            if (!HasAtLeastTwoShapes(shapes)) return;

            var reference = GetReferenceShapeForSize(shapes, mode);
            if (!IsResizableShape(reference)) return;

            float referenceHeight;
            try
            {
                referenceHeight = reference.Height;
                if (referenceHeight <= 0f) return;
            }
            catch
            {
                return;
            }

            foreach (var shape in shapes)
            {
                if (!IsResizableShape(shape) || IsReferenceShape(shape, reference)) continue;

                try
                {
                    if (shape.Height <= 0f) continue;
                    shape.Height = referenceHeight;
                }
                catch (COMException)
                {
                    // Skip COM-failed shape.
                }
                catch
                {
                    // Skip non-fatal shape error.
                }
            }
        }

        public void MatchBoth(List<PowerPoint.Shape> shapes, ReferenceMode mode)
        {
            if (!HasAtLeastTwoShapes(shapes)) return;

            var reference = GetReferenceShapeForSize(shapes, mode);
            if (!IsResizableShape(reference)) return;

            float referenceWidth;
            float referenceHeight;
            try
            {
                referenceWidth = reference.Width;
                referenceHeight = reference.Height;
                if (referenceWidth <= 0f || referenceHeight <= 0f) return;
            }
            catch
            {
                return;
            }

            foreach (var shape in shapes)
            {
                if (!IsResizableShape(shape) || IsReferenceShape(shape, reference)) continue;

                try
                {
                    if (shape.Width <= 0f || shape.Height <= 0f) continue;
                    shape.Width = referenceWidth;
                    shape.Height = referenceHeight;
                }
                catch (COMException)
                {
                    // Skip COM-failed shape.
                }
                catch
                {
                    // Skip non-fatal shape error.
                }
            }
        }

        public void FitToSlideWidth(List<PowerPoint.Shape> shapes)
        {
            if (!HasShapes(shapes)) return;

            var slideSize = GetSlideSize(Globals.ThisAddIn.Application);
            if (slideSize == SizeF.Empty || slideSize.Width <= 0f) return;

            foreach (var shape in shapes)
            {
                if (!IsResizableShape(shape)) continue;

                try
                {
                    shape.Left = 0f;
                    shape.Width = slideSize.Width;
                }
                catch (COMException)
                {
                    // Skip COM-failed shape.
                }
                catch
                {
                    // Skip non-fatal shape error.
                }
            }
        }

        public void FitToSlideHeight(List<PowerPoint.Shape> shapes)
        {
            if (!HasShapes(shapes)) return;

            var slideSize = GetSlideSize(Globals.ThisAddIn.Application);
            if (slideSize == SizeF.Empty || slideSize.Height <= 0f) return;

            foreach (var shape in shapes)
            {
                if (!IsResizableShape(shape)) continue;

                try
                {
                    shape.Top = 0f;
                    shape.Height = slideSize.Height;
                }
                catch (COMException)
                {
                    // Skip COM-failed shape.
                }
                catch
                {
                    // Skip non-fatal shape error.
                }
            }
        }

        public void FitToSlideBoth(List<PowerPoint.Shape> shapes)
        {
            if (!HasShapes(shapes)) return;

            var slideSize = GetSlideSize(Globals.ThisAddIn.Application);
            if (slideSize == SizeF.Empty || slideSize.Width <= 0f || slideSize.Height <= 0f) return;

            foreach (var shape in shapes)
            {
                if (!IsResizableShape(shape)) continue;

                try
                {
                    shape.Left = 0f;
                    shape.Top = 0f;
                    shape.Width = slideSize.Width;
                    shape.Height = slideSize.Height;
                }
                catch (COMException)
                {
                    // Skip COM-failed shape.
                }
                catch
                {
                    // Skip non-fatal shape error.
                }
            }
        }

        private static bool HasAtLeastTwoShapes(IList<PowerPoint.Shape> shapes)
        {
            if (!HasShapes(shapes)) return false;

            int count = 0;
            foreach (var shape in shapes)
            {
                if (!IsResizableShape(shape)) continue;

                count++;
                if (count >= 2) return true;
            }

            return false;
        }

        private static PowerPoint.Shape GetFirstResizableShape(IList<PowerPoint.Shape> shapes)
        {
            if (!HasShapes(shapes)) return null;

            foreach (var shape in shapes)
            {
                if (IsResizableShape(shape))
                    return shape;
            }

            return null;
        }

        private static bool IsResizableShape(PowerPoint.Shape shape)
        {
            if (shape == null) return false;

            try
            {
                var type = shape.Type;
                if (type == Office.MsoShapeType.msoGroup ||
                    type == Office.MsoShapeType.msoLine)
                {
                    return false;
                }

                if (shape.Connector == Office.MsoTriState.msoTrue)
                {
                    return false;
                }

                return shape.Width > 0f && shape.Height > 0f;
            }
            catch (COMException)
            {
                return false;
            }
            catch
            {
                return false;
            }
        }

        private static float GetRightEdge(PowerPoint.Shape shape)
        {
            return shape.Left + shape.Width;
        }

        private static float GetBottomEdge(PowerPoint.Shape shape)
        {
            return shape.Top + shape.Height;
        }

        private static bool HasShapes(IList<PowerPoint.Shape> shapes)
        {
            return shapes != null && shapes.Count > 0;
        }

        private static bool IsReferenceShape(PowerPoint.Shape shape, PowerPoint.Shape reference)
        {
            if (shape == null || reference == null) return false;

            try
            {
                return shape.Id == reference.Id;
            }
            catch
            {
                return false;
            }
        }
    }
}
