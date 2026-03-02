using System;
using System.Collections.Generic;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PowerPointAddIn1
{
    public static class ZoomSelection
    {
        public static void SelectZoomAreas(PowerPoint.Application app)
        {
            var win = app?.ActiveWindow;
            if (win == null) throw new InvalidOperationException("No active PowerPoint window.");

            var slide = win.View?.Slide;
            if (slide == null) throw new InvalidOperationException("Open a slide in Normal view.");

            var found = new List<PowerPoint.Shape>();

            // 1) Prefer shapes named ZOOM_*
            for (int i = 1; i <= slide.Shapes.Count; i++)
            {
                var sh = slide.Shapes[i];
                if (!string.IsNullOrEmpty(sh.Name) &&
                    sh.Name.StartsWith("ZOOM_", StringComparison.OrdinalIgnoreCase) &&
                    sh.Width > 2 && sh.Height > 2)
                {
                    found.Add(sh);
                }
            }

            // 2) Fallback: AutoShape / Freeform / Group (exclude placeholders/pictures/text)
            if (found.Count == 0)
            {
                for (int i = 1; i <= slide.Shapes.Count; i++)
                {
                    var sh = slide.Shapes[i];

                    int type = (int)sh.Type; // IMPORTANT: your interop returns int

                    if (type == (int)Office.MsoShapeType.msoPlaceholder) continue;
                    if (type == (int)Office.MsoShapeType.msoPicture) continue;

                    bool hasText = false;
                    try
                    {
                        if (sh.HasTextFrame == Office.MsoTriState.msoTrue &&
                            sh.TextFrame.HasText == Office.MsoTriState.msoTrue)
                            hasText = true;
                    }
                    catch { }

                    if (hasText) continue;

                    bool candidate =
                        type == (int)Office.MsoShapeType.msoAutoShape ||
                        type == (int)Office.MsoShapeType.msoFreeform ||
                        type == (int)Office.MsoShapeType.msoGroup;

                    if (!candidate) continue;
                    if (sh.Width <= 2 || sh.Height <= 2) continue;

                    found.Add(sh);
                }
            }

            if (found.Count == 0)
                throw new InvalidOperationException(
                    "No zoom shapes found.\n\n" +
                    "Fix: rename zoom shapes to ZOOM_1, ZOOM_2... OR draw shapes with real area.");

            slide.Select();
            found[0].Select(Office.MsoTriState.msoTrue);
            for (int i = 1; i < found.Count; i++)
                found[i].Select(Office.MsoTriState.msoTrue);
        }
    }
}