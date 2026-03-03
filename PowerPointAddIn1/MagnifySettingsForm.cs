using System;
using System.Drawing;
using System.Windows.Forms;

namespace PowerPointAddIn1
{
    /// <summary>
    /// Settings dialog for the Magnify Glass effect:
    ///   • Magnification Factor
    ///   • Border Color
    ///   • Border Thickness
    ///   • Soft Edges preset
    /// </summary>
    internal class MagnifySettingsForm : Form
    {
        private static readonly float[] SoftEdgeValues =
            { 0f, 1f, 2.5f, 5f, 10f, 25f, 50f };

        private Panel         pnlHeader;
        private Label         lblTitle;
        private Label         lblFactor;
        private NumericUpDown nudFactor;
        private Label         lblBorderColor;
        private Panel         pnlBorderColor;
        private Label         lblBorderThickness;
        private NumericUpDown nudBorderThickness;
        private Label         lblSoftEdges;
        private ComboBox      cboSoftEdges;
        private Button        btnOK;
        private Button        btnCancel;

        public MagnifySettings Result { get; private set; }

        public MagnifySettingsForm(MagnifySettings current)
        {
            Result = Clone(current);
            BuildLayout();
            Populate(current);
        }

        // ── Form construction ─────────────────────────────────────────────────
        private void BuildLayout()
        {
            Text            = "Magnify Settings";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox     = false;
            MinimizeBox     = false;
            StartPosition   = FormStartPosition.CenterParent;
            ClientSize      = new Size(380, 272);
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
                Text      = "MAGNIFY SETTINGS",
                ForeColor = Color.White,
                Font      = new Font("Segoe UI", 10f, FontStyle.Bold),
                Location  = new Point(12, 0),
                Size      = new Size(356, 38),
                TextAlign = ContentAlignment.MiddleLeft
            };
            pnlHeader.Controls.Add(lblTitle);

            // ── Row 1: Magnification Factor ───────────────────────────────────
            lblFactor = new Label
            {
                Text      = "Magnification Factor",
                Location  = new Point(20, 62),
                AutoSize  = true,
                TextAlign = ContentAlignment.MiddleLeft
            };
            nudFactor = new NumericUpDown
            {
                Location      = new Point(256, 59),
                Size          = new Size(80, 24),
                Minimum       = (decimal)1.1,
                Maximum       = 10m,
                DecimalPlaces = 1,
                Increment     = (decimal)0.5,
                TextAlign     = HorizontalAlignment.Center
            };

            // ── Row 2: Border Color ───────────────────────────────────────────
            lblBorderColor = new Label
            {
                Text      = "Border Color",
                Location  = new Point(20, 106),
                AutoSize  = true,
                TextAlign = ContentAlignment.MiddleLeft
            };
            pnlBorderColor = new Panel
            {
                Location    = new Point(256, 103),
                Size        = new Size(80, 24),
                BackColor   = Color.Black,
                BorderStyle = BorderStyle.FixedSingle,
                Cursor      = Cursors.Hand
            };
            new ToolTip().SetToolTip(pnlBorderColor, "Click to choose border colour");
            pnlBorderColor.Click += (s, e) =>
            {
                using (var cd = new ColorDialog { Color = pnlBorderColor.BackColor, FullOpen = true })
                    if (cd.ShowDialog() == DialogResult.OK)
                        pnlBorderColor.BackColor = cd.Color;
            };

            // ── Row 3: Border Thickness ───────────────────────────────────────
            lblBorderThickness = new Label
            {
                Text      = "Border Thickness (pt)",
                Location  = new Point(20, 150),
                AutoSize  = true,
                TextAlign = ContentAlignment.MiddleLeft
            };
            nudBorderThickness = new NumericUpDown
            {
                Location      = new Point(256, 147),
                Size          = new Size(80, 24),
                Minimum       = 0m,
                Maximum       = 20m,
                DecimalPlaces = 1,
                Increment     = (decimal)0.5,
                TextAlign     = HorizontalAlignment.Center
            };

            // ── Row 4: Soft Edges ─────────────────────────────────────────────
            lblSoftEdges = new Label
            {
                Text     = "Soft Edges",
                Location = new Point(20, 194),
                AutoSize = true
            };
            cboSoftEdges = new ComboBox
            {
                Location      = new Point(200, 191),
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

            // ── Buttons ───────────────────────────────────────────────────────
            btnOK = new Button
            {
                Text     = "OK",
                Location = new Point(195, 232),
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
                Location     = new Point(285, 232),
                Size         = new Size(80, 28),
                DialogResult = DialogResult.Cancel
            };

            AcceptButton = btnOK;
            CancelButton = btnCancel;

            Controls.AddRange(new Control[]
            {
                pnlHeader,
                lblFactor,          nudFactor,
                lblBorderColor,     pnlBorderColor,
                lblBorderThickness, nudBorderThickness,
                lblSoftEdges,       cboSoftEdges,
                btnOK,              btnCancel
            });
        }

        // ── Populate controls from settings ───────────────────────────────────
        private void Populate(MagnifySettings s)
        {
            nudFactor.Value          = (decimal)Math.Max(1.1f, Math.Min(10f, s.MagnificationFactor));
            pnlBorderColor.BackColor = s.BorderColor;
            nudBorderThickness.Value = (decimal)Math.Max(0f, Math.Min(20f, s.BorderThickness));

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
        }

        // ── Collect control values into Result ────────────────────────────────
        private void SaveResult()
        {
            Result = new MagnifySettings
            {
                MagnificationFactor = (float)nudFactor.Value,
                BorderColor         = pnlBorderColor.BackColor,
                BorderThickness     = (float)nudBorderThickness.Value,
                SoftEdgesPoints     = SoftEdgeValues[Math.Max(0, cboSoftEdges.SelectedIndex)]
            };
        }

        private static MagnifySettings Clone(MagnifySettings s) =>
            new MagnifySettings
            {
                MagnificationFactor = s.MagnificationFactor,
                BorderColor         = s.BorderColor,
                BorderThickness     = s.BorderThickness,
                SoftEdgesPoints     = s.SoftEdgesPoints
            };
    }
}
