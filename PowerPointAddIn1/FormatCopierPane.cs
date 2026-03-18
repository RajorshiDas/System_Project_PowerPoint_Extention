using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1
{
    public class FormatCopierPane : UserControl
    {
        private Label _sourceShapeLabel;
        private CheckBox _copySizeCheckBox;
        private CheckBox _copyPositionCheckBox;
        private CheckBox _copyAnimationCheckBox;
        private Label _statusLabel;

        private int _sourceSlideIndex = -1;
        private string _sourceShapeName;

        public FormatCopierPane()
        {
            InitializeControls();
        }

        private void InitializeControls()
        {
            BackColor = Color.White;
            Dock = DockStyle.Fill;
            AutoScroll = true;

            int y = 10;
            const int left = 12;
            const int width = 230;
            const int buttonHeight = 30;

            var titleLabel = new Label
            {
                Text = "Format Copier",
                AutoSize = true,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = Color.DarkSlateGray,
                Location = new Point(left, y)
            };
            Controls.Add(titleLabel);
            y += 28;

            var sourceHeader = new Label
            {
                Text = "SOURCE SHAPE",
                AutoSize = true,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.DarkSlateGray,
                Location = new Point(left, y)
            };
            Controls.Add(sourceHeader);
            y += 20;

            _sourceShapeLabel = new Label
            {
                Text = "Source: (none)",
                AutoSize = true,
                Font = new Font("Segoe UI", 8),
                ForeColor = Color.DimGray,
                Location = new Point(left, y)
            };
            Controls.Add(_sourceShapeLabel);
            y += 20;

            y = AddButton("Capture Selected Shape", y, left, width, buttonHeight, CaptureSelectedShape_Click);
            y += 6;

            var optionsHeader = new Label
            {
                Text = "COPY OPTIONS",
                AutoSize = true,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.DarkSlateGray,
                Location = new Point(left, y)
            };
            Controls.Add(optionsHeader);
            y += 20;

            _copySizeCheckBox = new CheckBox
            {
                Text = "Copy Size",
                Checked = true,
                AutoSize = true,
                Font = new Font("Segoe UI", 9),
                Location = new Point(left, y)
            };
            Controls.Add(_copySizeCheckBox);
            y += 24;

            _copyPositionCheckBox = new CheckBox
            {
                Text = "Copy Position",
                Checked = true,
                AutoSize = true,
                Font = new Font("Segoe UI", 9),
                Location = new Point(left, y)
            };
            Controls.Add(_copyPositionCheckBox);
            y += 24;

            _copyAnimationCheckBox = new CheckBox
            {
                Text = "Copy Animation",
                AutoSize = true,
                Font = new Font("Segoe UI", 9),
                Location = new Point(left, y)
            };
            Controls.Add(_copyAnimationCheckBox);
            y += 28;

            y = AddButton("Apply Copy", y, left, width, buttonHeight, ApplyCopy_Click);
            y = AddButton("Clear Source", y, left, width, buttonHeight, ClearSource_Click);

            _statusLabel = new Label
            {
                Text = "Ready",
                AutoSize = true,
                Font = new Font("Segoe UI", 8),
                ForeColor = Color.DimGray,
                Location = new Point(left, y + 4)
            };
            Controls.Add(_statusLabel);
        }

        private int AddButton(string text, int y, int left, int width, int height, EventHandler onClick)
        {
            var button = new Button
            {
                Text = text,
                Location = new Point(left, y),
                Size = new Size(width, height),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(230, 230, 250),
                Font = new Font("Segoe UI", 9)
            };
            button.Click += onClick;
            Controls.Add(button);
            return y + height + 5;
        }

        private void CaptureSelectedShape_Click(object sender, EventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var selectedShapes = GetSelectedShapes(app);

                if (selectedShapes.Count == 0)
                {
                    MessageBox.Show("Please select a source shape first.", "Format Copier",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var source = selectedShapes[0];
                var slide = app.ActiveWindow.View.Slide as PowerPoint.Slide;
                if (slide == null)
                {
                    MessageBox.Show("Unable to determine active slide.", "Format Copier",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                _sourceSlideIndex = slide.SlideIndex;
                _sourceShapeName = source.Name;
                _sourceShapeLabel.Text = "Source: " + _sourceShapeName;
                _statusLabel.Text = "Source captured.";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Capture failed: " + ex.Message, "Format Copier",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ApplyCopy_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(_sourceShapeName) || _sourceSlideIndex < 1)
                {
                    MessageBox.Show("No source shape captured.", "Format Copier",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var app = Globals.ThisAddIn.Application;
                if (app.Presentations.Count == 0)
                {
                    MessageBox.Show("Please open a presentation first.", "Format Copier",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var presentation = app.ActivePresentation;
                if (_sourceSlideIndex > presentation.Slides.Count)
                {
                    MessageBox.Show("Source slide no longer exists.", "Format Copier",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var sourceSlide = presentation.Slides[_sourceSlideIndex];
                var sourceShape = FindShapeOnSlide(sourceSlide, _sourceShapeName);
                if (sourceShape == null)
                {
                    MessageBox.Show("Source shape no longer exists.", "Format Copier",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var targets = GetSelectedShapes(app);
                if (targets.Count == 0)
                {
                    MessageBox.Show("Please select one or more target shapes.", "Format Copier",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                bool copySize = _copySizeCheckBox.Checked;
                bool copyPosition = _copyPositionCheckBox.Checked;
                bool copyAnimation = _copyAnimationCheckBox.Checked;

                if (!copySize && !copyPosition && !copyAnimation)
                {
                    MessageBox.Show("Select at least one copy option.", "Format Copier",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                int copied = 0;
                foreach (var target in targets)
                {
                    if (target == null) continue;
                    if (IsSameShape(target, sourceShape)) continue;

                    try
                    {
                        if (copySize)
                        {
                            target.Width = sourceShape.Width;
                            target.Height = sourceShape.Height;
                        }

                        if (copyPosition)
                        {
                            target.Left = sourceShape.Left;
                            target.Top = sourceShape.Top;
                        }

                        if (copyAnimation)
                        {
                            CopyAnimations(sourceShape, target);
                        }

                        copied++;
                    }
                    catch (COMException)
                    {
                        // Skip COM-failed target shape.
                    }
                }

                if (copied == 0)
                {
                    MessageBox.Show("No target shapes were updated.", "Format Copier",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    _statusLabel.Text = "No changes applied.";
                    return;
                }

                _statusLabel.Text = string.Format("Applied to {0} shape(s).", copied);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Apply failed: " + ex.Message, "Format Copier",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ClearSource_Click(object sender, EventArgs e)
        {
            _sourceSlideIndex = -1;
            _sourceShapeName = null;
            _sourceShapeLabel.Text = "Source: (none)";
            _statusLabel.Text = "Source cleared.";
        }

        private static List<PowerPoint.Shape> GetSelectedShapes(PowerPoint.Application app)
        {
            var shapes = new List<PowerPoint.Shape>();
            if (app == null) return shapes;

            try
            {
                var window = app.ActiveWindow;
                if (window == null || window.Selection == null) return shapes;

                var selection = window.Selection;
                if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes &&
                    selection.Type != PowerPoint.PpSelectionType.ppSelectionText)
                {
                    return shapes;
                }

                var range = selection.ShapeRange;
                if (range == null) return shapes;

                for (int i = 1; i <= range.Count; i++)
                {
                    if (range[i] != null)
                        shapes.Add(range[i]);
                }
            }
            catch
            {
                // Return empty list on selection errors.
            }

            return shapes;
        }

        private static PowerPoint.Shape FindShapeOnSlide(PowerPoint.Slide slide, string name)
        {
            if (slide == null || string.IsNullOrEmpty(name)) return null;

            try
            {
                for (int i = 1; i <= slide.Shapes.Count; i++)
                {
                    if (string.Equals(slide.Shapes[i].Name, name, StringComparison.OrdinalIgnoreCase))
                        return slide.Shapes[i];
                }
            }
            catch
            {
                // Ignore lookup failures.
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
            if (sourceShape == null || targetShape == null) return;

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
                        // Ignore delete failure for individual effect.
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
                        // Ignore unsupported effect copy scenario.
                    }
                }
            }
            catch
            {
                // Keep operation resilient.
            }
        }
    }
}
