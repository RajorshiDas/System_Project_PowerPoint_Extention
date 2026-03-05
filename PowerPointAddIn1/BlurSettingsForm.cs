using System;
using System.Drawing;
using System.Windows.Forms;

namespace PowerPointAddIn1
{
    /// <summary>
    /// Settings dialog for the Blur effect:
    ///   • Blur Intensity (0–100 %)
    ///   • Soft Edges preset
    /// </summary>
    internal class BlurSettingsForm : Form
    {
        private static readonly float[] SoftEdgeValues =
            { 0f, 1f, 2.5f, 5f, 10f, 25f, 50f };

        private Panel         pnlHeader;
        private Label         lblTitle;
        private Label         lblIntensity;
        private NumericUpDown nudIntensity;
        private Label         lblPct;
        private TrackBar      trkIntensity;
        private Label         lblPreviewMin;
        private Label         lblPreviewMax;
        private Label         lblSoftEdges;
        private ComboBox      cboSoftEdges;
        private Button        btnOK;
        private Button        btnCancel;

        public BlurSettings Result { get; private set; }

        public BlurSettingsForm(BlurSettings current)
        {
            Result = new BlurSettings
            {
                Intensity       = current.Intensity,
                SoftEdgesPoints = current.SoftEdgesPoints
            };
            BuildLayout();
            Populate(current);
        }

        // ── Form construction ─────────────────────────────────────────────────
        private void BuildLayout()
        {
            Text            = "Blur Settings";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox     = false;
            MinimizeBox     = false;
            StartPosition   = FormStartPosition.CenterParent;
            ClientSize      = new Size(380, 262);
            BackColor       = Color.White;
            Font            = new Font("Segoe UI", 9f);

            // ── Blue header ───────────────────────────────────────────────────
            pnlHeader = new Panel
            {
                Location  = new Point(0, 0),
                Size      = new Size(380, 38),
                BackColor = Color.FromArgb(0, 114, 198)
            };
            lblTitle = new Label
            {
                Text      = "BLUR SETTINGS",
                ForeColor = Color.White,
                Font      = new Font("Segoe UI", 10f, FontStyle.Bold),
                Location  = new Point(12, 0),
                Size      = new Size(356, 38),
                TextAlign = ContentAlignment.MiddleLeft
            };
            pnlHeader.Controls.Add(lblTitle);
            Controls.Add(pnlHeader);

            // ── Row 1: Intensity label + numeric spinner ──────────────────────
            lblIntensity = new Label
            {
                Text      = "Blur Intensity",
                Location  = new Point(20, 62),
                AutoSize  = true,
                TextAlign = ContentAlignment.MiddleLeft
            };
            nudIntensity = new NumericUpDown
            {
                Location  = new Point(256, 59),
                Size      = new Size(72, 24),
                Minimum   = 0,
                Maximum   = 100,
                TextAlign = HorizontalAlignment.Center
            };
            lblPct = new Label
            {
                Text     = "%",
                Location = new Point(333, 62),
                AutoSize = true
            };

            // ── Row 2: TrackBar ───────────────────────────────────────────────
            trkIntensity = new TrackBar
            {
                Location      = new Point(20,  92),
                Size          = new Size(310, 40),
                Minimum       = 0,
                Maximum       = 100,
                TickFrequency = 10,
                SmallChange   = 1,
                LargeChange   = 10
            };
            lblPreviewMin = new Label
            {
                Text      = "None",
                Location  = new Point(20, 132),
                AutoSize  = true,
                ForeColor = Color.Gray
            };
            lblPreviewMax = new Label
            {
                Text      = "Max",
                Location  = new Point(306, 132),
                AutoSize  = true,
                ForeColor = Color.Gray
            };

            // Keep spinner and trackbar in sync
            nudIntensity.ValueChanged += (s, e) =>
            {
                if (trkIntensity.Value != (int)nudIntensity.Value)
                    trkIntensity.Value = (int)nudIntensity.Value;
            };
            trkIntensity.ValueChanged += (s, e) =>
            {
                if (nudIntensity.Value != trkIntensity.Value)
                    nudIntensity.Value = trkIntensity.Value;
            };

            // ── Row 3: Soft Edges ─────────────────────────────────────────────
            lblSoftEdges = new Label
            {
                Text     = "Soft Edges",
                Location = new Point(20, 166),
                AutoSize = true
            };
            cboSoftEdges = new ComboBox
            {
                Location      = new Point(200, 163),
                Size          = new Size(136, 24),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cboSoftEdges.Items.AddRange(new object[]
            {
                "None",
                "1 Point",
                "2.5 Points",
                "5 Points",
                "10 Points",
                "25 Points",
                "50 Points"
            });

            Controls.Add(lblIntensity);
            Controls.Add(nudIntensity);
            Controls.Add(lblPct);
            Controls.Add(trkIntensity);
            Controls.Add(lblPreviewMin);
            Controls.Add(lblPreviewMax);
            Controls.Add(lblSoftEdges);
            Controls.Add(cboSoftEdges);

            // ── Buttons ───────────────────────────────────────────────────────
            btnOK = new Button
            {
                Text     = "OK",
                Location = new Point(195, 222),
                Size     = new Size(80, 28)
            };
            btnOK.Click += (s, e) =>
            {
                SaveResult();
                DialogResult = DialogResult.OK;
                Close();
            };
            btnCancel = new Button
            {
                Text         = "Cancel",
                DialogResult = DialogResult.Cancel,
                Location     = new Point(285, 222),
                Size         = new Size(80, 28)
            };

            Controls.Add(btnOK);
            Controls.Add(btnCancel);
            AcceptButton = btnOK;
            CancelButton = btnCancel;
        }

        // ── Populate controls from settings ──────────────────────────────────
        private void Populate(BlurSettings current)
        {
            nudIntensity.Value = Math.Max(0, Math.Min(100, current.Intensity));
            trkIntensity.Value = (int)nudIntensity.Value;

            int idx = 0;
            for (int i = 0; i < SoftEdgeValues.Length; i++)
            {
                if (Math.Abs(SoftEdgeValues[i] - current.SoftEdgesPoints) < 0.01f)
                {
                    idx = i;
                    break;
                }
            }
            cboSoftEdges.SelectedIndex = idx;
        }

        // ── Collect control values into Result ────────────────────────────────
        private void SaveResult()
        {
            Result = new BlurSettings
            {
                Intensity       = (int)nudIntensity.Value,
                SoftEdgesPoints = SoftEdgeValues[Math.Max(0, cboSoftEdges.SelectedIndex)]
            };
        }
    }
}
