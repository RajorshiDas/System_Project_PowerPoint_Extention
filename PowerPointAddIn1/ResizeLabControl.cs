using System;
using System.Drawing;
using System.Windows.Forms;

namespace PowerPointAddIn1
{
    public class ResizeLabControl : UserControl
    {
        private readonly ResizeLabService _service = new ResizeLabService();
        private ReferenceMode _referenceMode = ReferenceMode.FirstSelected;
        private Label _modeLabel;

        public ResizeLabControl()
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
            const int width = 220;
            const int height = 30;

            var title = new Label
            {
                Text = "RESIZE LAB",
                AutoSize = true,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = Color.DarkSlateGray,
                Location = new Point(left, y)
            };
            Controls.Add(title);
            y += 24;

            _modeLabel = new Label
            {
                Text = "Reference: First selected object",
                AutoSize = true,
                Font = new Font("Segoe UI", 8),
                ForeColor = Color.DimGray,
                Location = new Point(left, y)
            };
            Controls.Add(_modeLabel);
            y += 24;

            y = AddButton("Settings", y, left, width, height, SettingsButton_Click);
            y += 6;

            y = AddHeader("STRETCH", y, left);
            y = AddButton("Stretch Left", y, left, width, height, (s, e) => ExecuteWithSelection(2, shapes => _service.StretchToLeft(shapes, _referenceMode)));
            y = AddButton("Stretch Right", y, left, width, height, (s, e) => ExecuteWithSelection(2, shapes => _service.StretchToRight(shapes, _referenceMode)));
            y = AddButton("Stretch Top", y, left, width, height, (s, e) => ExecuteWithSelection(2, shapes => _service.StretchToTop(shapes, _referenceMode)));
            y = AddButton("Stretch Bottom", y, left, width, height, (s, e) => ExecuteWithSelection(2, shapes => _service.StretchToBottom(shapes, _referenceMode)));
            y += 6;

            y = AddHeader("MATCH", y, left);
            y = AddButton("Match Width", y, left, width, height, (s, e) => ExecuteWithSelection(2, shapes => _service.MatchWidth(shapes, _referenceMode)));
            y = AddButton("Match Height", y, left, width, height, (s, e) => ExecuteWithSelection(2, shapes => _service.MatchHeight(shapes, _referenceMode)));
            y = AddButton("Match Both", y, left, width, height, (s, e) => ExecuteWithSelection(2, shapes => _service.MatchBoth(shapes, _referenceMode)));
            y += 6;

            y = AddHeader("FIT TO SLIDE", y, left);
            y = AddButton("Fit Width", y, left, width, height, (s, e) => ExecuteWithSelection(1, _service.FitToSlideWidth));
            y = AddButton("Fit Height", y, left, width, height, (s, e) => ExecuteWithSelection(1, _service.FitToSlideHeight));
            AddButton("Fit Both", y, left, width, height, (s, e) => ExecuteWithSelection(1, _service.FitToSlideBoth));
        }

        private int AddHeader(string text, int y, int left)
        {
            var header = new Label
            {
                Text = text,
                AutoSize = true,
                Font = new Font("Segoe UI", 9, FontStyle.Bold),
                ForeColor = Color.DarkSlateGray,
                Location = new Point(left, y)
            };
            Controls.Add(header);
            return y + 20;
        }

        private int AddButton(string text, int y, int left, int width, int height, EventHandler onClick)
        {
            var btn = new Button
            {
                Text = text,
                Location = new Point(left, y),
                Size = new Size(width, height),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(230, 230, 250),
                Font = new Font("Segoe UI", 9)
            };
            btn.Click += onClick;
            Controls.Add(btn);
            return y + height + 5;
        }

        private void SettingsButton_Click(object sender, EventArgs e)
        {
            using (var form = new ResizeLabSettingsForm(_referenceMode))
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    _referenceMode = form.SelectedReferenceMode;
                    _modeLabel.Text = _referenceMode == ReferenceMode.OutermostObject
                        ? "Reference: Outermost object"
                        : "Reference: First selected object";
                }
            }
        }

        private void ExecuteWithSelection(int minimumCount, Action<System.Collections.Generic.List<Microsoft.Office.Interop.PowerPoint.Shape>> action)
        {
            var shapes = _service.GetSelectedShapes(Globals.ThisAddIn.Application);
            if (shapes == null || shapes.Count < minimumCount)
            {
                string msg = minimumCount <= 1
                    ? "Please select at least one shape."
                    : "Please select at least two shapes.";

                MessageBox.Show(msg, "Resize Lab", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                action(shapes);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Resize Lab error: " + ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
