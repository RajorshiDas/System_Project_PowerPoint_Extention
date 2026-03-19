using System.Text;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1
{
    public class SlideTextService
    {
        public PowerPoint.Slide GetActiveSlide(PowerPoint.Application app)
        {
            try
            {
                if (app == null)
                    return null;

                if (app.ActiveWindow == null)
                    return null;

                if (app.ActiveWindow.View == null)
                    return null;

                return app.ActiveWindow.View.Slide;
            }
            catch
            {
                return null;
            }
        }

        public string ExtractTextFromSlide(PowerPoint.Slide slide)
        {
            if (slide == null)
                return string.Empty;

            StringBuilder sb = new StringBuilder();

            foreach (PowerPoint.Shape shape in slide.Shapes)
            {
                try
                {
                    if (shape.HasTextFrame == MsoTriState.msoTrue &&
                        shape.TextFrame.HasText == MsoTriState.msoTrue)
                    {
                        string text = shape.TextFrame.TextRange.Text;

                        if (!string.IsNullOrWhiteSpace(text))
                        {
                            sb.AppendLine(text.Trim());
                        }
                    }
                }
                catch
                {
                }
            }

            return sb.ToString().Trim();
        }
    }
}