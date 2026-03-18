using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1
{
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
    }
}
