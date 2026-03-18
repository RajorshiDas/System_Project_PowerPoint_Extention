using System;
using System.Drawing;
using System.Windows.Forms;

namespace PowerPointAddIn1
{
    /// <summary>
    /// Settings dialog that matches the Spotlight Settings design:
    ///   • Spotlight Transparency (%)
    ///   • Soft Edges (preset points)
    ///   • Spotlight (overlay) Color
    /// </summary>
    internal class SpotlightSettingsForm : Form
    {
        // Preset soft-edge radii matching PowerPoint's built-in dropdown
        private static readonly float[] SoftEdgeValues =
            { 0f, 1f, 2.5f, 5f, 10f, 25f, 50f };

        private Panel          pnlHeader;
        private Label          lblTitle;
        private Label          lblTransparency;
        private NumericUpDown  nudTransparency;
        private Label          lblPct;
        private Label          lblSoftEdges;
        private ComboBox       cboSoftEdges;
        private Label          lblColor;
        private Panel          pnlColorSwatch;
        private Button         btnOK;
        private Button         btnCancel;

        public SpotlightSettings Result { get; private set; }

        public SpotlightSettingsForm(SpotlightSettings current)
        {
            Result = Clone(current);
            BuildLayout();
            Populate(current);
        }

        // ── Form construction ─────────────────────────────────────────────────
        private void BuildLayout()
        {
            Text            = "Spotlight Settings";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox     = false;
            MinimizeBox     = false;
            StartPosition   = FormStartPosition.CenterParent;
            ClientSize      = new Size(380, 218);
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
                Text      = "SPOTLIGHT SETTINGS",
                ForeColor = Color.White,
                Font      = new Font("Segoe UI", 10f, FontStyle.Bold),
                Location  = new Point(12, 0),
                Size      = new Size(356, 38),
                TextAlign = ContentAlignment.MiddleLeft
            };
            pnlHeader.Controls.Add(lblTitle);

            // ── Row 1: Transparency ───────────────────────────────────────────
            lblTransparency = new Label
            {
                Text      = "Spotlight Transparency",
                Location  = new Point(20, 62),
                AutoSize  = true,
                TextAlign = ContentAlignment.MiddleLeft
            };
            nudTransparency = new NumericUpDown
            {
                Location    = new Point(260, 59),
                Size        = new Size(58, 24),
                Minimum     = 0,
                Maximum     = 100,
                DecimalPlaces = 0,
                TextAlign   = HorizontalAlignment.Center
            };
            lblPct = new Label
            {
                Text     = "%",
                Location = new Point(322, 62),
                AutoSize = true
            };

            // ── Row 2: Soft Edges ─────────────────────────────────────────────
            lblSoftEdges = new Label
            {
                Text      = "Soft Edges",
                Location  = new Point(20, 106),
                AutoSize  = true,
                TextAlign = ContentAlignment.MiddleLeft
            };
            cboSoftEdges = new ComboBox
            {
                Location      = new Point(220, 103),
                Size          = new Size(140, 24),
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

            // ── Row 3: Overlay color ──────────────────────────────────────────
            lblColor = new Label
            {
                Text      = "Spotlight Color",
                Location  = new Point(20, 150),
                AutoSize  = true,
                TextAlign = ContentAlignment.MiddleLeft
            };
            pnlColorSwatch = new Panel
            {
                Location    = new Point(260, 147),
                Size        = new Size(60, 24),
                BackColor   = Color.Black,
                BorderStyle = BorderStyle.FixedSingle,
                Cursor      = Cursors.Hand
            };
            var tooltip = new ToolTip();
            tooltip.SetToolTip(pnlColorSwatch, "Click to choose overlay colour");
            pnlColorSwatch.Click += (s, e) =>
            {
                using (var cd = new ColorDialog { Color = pnlColorSwatch.BackColor, FullOpen = true })
                    if (cd.ShowDialog() == DialogResult.OK)
                        pnlColorSwatch.BackColor = cd.Color;
            };

            // ── Buttons ───────────────────────────────────────────────────────
            btnOK = new Button
            {
                Text     = "OK",
                Location = new Point(195, 180),
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
                Text         = "CANCEL",
                Location     = new Point(285, 180),
                Size         = new Size(80, 28),
                DialogResult = DialogResult.Cancel
            };

            AcceptButton = btnOK;
            CancelButton = btnCancel;

            Controls.AddRange(new Control[]
            {
                pnlHeader,
                lblTransparency, nudTransparency, lblPct,
                lblSoftEdges,    cboSoftEdges,
                lblColor,        pnlColorSwatch,
                btnOK,           btnCancel
            });
        }

        // ── Populate controls from settings ───────────────────────────────────
        private void Populate(SpotlightSettings s)
        {
            nudTransparency.Value = (decimal)Math.Max(0, Math.Min(100, s.TransparencyPercent));

            int idx = 0;
            for (int i = 0; i < SoftEdgeValues.Length; i++)
            {
                if (Math.Abs(SoftEdgeValues[i] - s.SoftEdgesPoints) < 0.01f)
                {
                    idx = i;
                    break;
                }
            }
            cboSoftEdges.SelectedIndex = idx;
            pnlColorSwatch.BackColor   = s.OverlayColor;
        }

        // ── Collect control values into Result ────────────────────────────────
        private void SaveResult()
        {
            Result = new SpotlightSettings
            {
                TransparencyPercent = (float)nudTransparency.Value,
                SoftEdgesPoints     = SoftEdgeValues[Math.Max(0, cboSoftEdges.SelectedIndex)],
                OverlayColor        = pnlColorSwatch.BackColor
            };
        }

        private static SpotlightSettings Clone(SpotlightSettings s) =>
            new SpotlightSettings
            {
                TransparencyPercent = s.TransparencyPercent,
                SoftEdgesPoints     = s.SoftEdgesPoints,
                OverlayColor        = s.OverlayColor
            };

        private void InitializeComponent()
        {
            this.SuspendLayout();
            // 
            // SpotlightSettingsForm
            // 
            this.ClientSize = new System.Drawing.Size(282, 253);
            this.Name = "SpotlightSettingsForm";
            this.Load += new System.EventHandler(this.SpotlightSettingsForm_Load);
            this.ResumeLayout(false);

        }

        private void SpotlightSettingsForm_Load(object sender, EventArgs e)
        {

        }
    }
}
