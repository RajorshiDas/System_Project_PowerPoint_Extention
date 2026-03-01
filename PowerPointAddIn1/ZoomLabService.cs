using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1
{
    public static class ZoomLabService
    {
        private static ZoomLabSettings _settings = new ZoomLabSettings();

        public static ZoomLabSettings Settings
        {
            get { return _settings; }
            set { _settings = value; }
        }

        public static void ZoomToArea(PowerPoint.Application app)
        {
            try
            {
                List<RectF> rectangles = GetSelectedRectangles(app);
                if (rectangles.Count == 0)
                {
                    return;
                }

                ZoomMath.BuildSingleAnimationSlide(app, rectangles);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message, "Zoom to Area", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static List<RectF> GetSelectedRectangles(PowerPoint.Application app)
        {
            List<RectF> result = new List<RectF>();

            if (app == null || app.ActiveWindow == null)
                return result;

            PowerPoint.Selection sel = app.ActiveWindow.Selection;
            if (sel == null)
                return result;

            PowerPoint.ShapeRange range = null;

            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                try
                {
                    range = sel.ShapeRange;
                }
                catch
                {
                    MessageBox.Show("Could not read ShapeRange. Try selecting the shape again.", "Zoom to Area");
                    return result;
                }
            }
            else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionText)
            {
                // User clicked inside a text box — get the parent shape
                try
                {
                    PowerPoint.TextRange textRange = sel.TextRange;
                    PowerPoint.Shape parentShape = textRange.Parent.Parent;
                    parentShape.Select();
                    range = app.ActiveWindow.Selection.ShapeRange;
                }
                catch
                {
                    MessageBox.Show(
                        "Could not get the shape from text selection.\r\n" +
                        "Try clicking the border of the shape instead of inside the text.",
                        "Zoom to Area");
                    return result;
                }
            }
            else
            {
                MessageBox.Show(
                    "Selection is not shapes.\r\n" +
                    "Selection.Type = " + sel.Type.ToString() + "\r\n\r\n" +
                    "Tip: Draw rectangle shapes on the areas you want to zoom into,\r\n" +
                    "then select them and click Zoom to Area.",
                    "Zoom to Area");
                return result;
            }

            if (range == null)
                return result;

            // Expand selection (handles groups)
            List<PowerPoint.Shape> shapes = new List<PowerPoint.Shape>();
            for (int i = 1; i <= range.Count; i++)
            {
                CollectShapesRecursive(range[i], shapes);
            }

            // Filter shapes with valid bounding boxes
            foreach (PowerPoint.Shape sh in shapes)
            {
                if (IsValidZoomShape(sh))
                {
                    result.Add(new RectF(sh.Left, sh.Top, sh.Width, sh.Height));
                }
            }

            // If still none, show debug info
            if (result.Count == 0)
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("No valid shapes detected.");
                sb.AppendLine("Selected shapes info:");
                foreach (PowerPoint.Shape sh in shapes)
                {
                    sb.AppendLine("- Name=" + sh.Name + ", Type=" + sh.Type.ToString() + ", AutoShapeType=" + TryGetAutoType(sh));
                }

                MessageBox.Show(sb.ToString(), "Zoom to Area Debug");
            }

            return result;
        }

        private static void CollectShapesRecursive(PowerPoint.Shape sh, List<PowerPoint.Shape> output)
        {
            if (sh == null)
                return;

            // If it is a group, add its children
            if (sh.Type == MsoShapeType.msoGroup)
            {
                try
                {
                    PowerPoint.GroupShapes groupItems = sh.GroupItems;
                    for (int i = 1; i <= groupItems.Count; i++)
                    {
                        CollectShapesRecursive(groupItems[i], output);
                    }
                }
                catch
                {
                    // If group access fails, just add the group shape itself
                    output.Add(sh);
                }
                return;
            }

            output.Add(sh);
        }

        private static bool IsValidZoomShape(PowerPoint.Shape sh)
        {
            // Accept any shape with a valid bounding box as a zoom target
            try
            {
                return sh.Width > 0 && sh.Height > 0;
            }
            catch
            {
                return false;
            }
        }

        private static string TryGetAutoType(PowerPoint.Shape sh)
        {
            try
            {
                if (sh.Type == MsoShapeType.msoAutoShape)
                {
                    return sh.AutoShapeType.ToString();
                }
                else
                {
                    return "(n/a)";
                }
            }
            catch
            {
                return "(error)";
            }
        }
    }
}
