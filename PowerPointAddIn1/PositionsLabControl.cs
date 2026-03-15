using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace PowerPointAddIn1
{
    public class PositionsLabControl : UserControl
    {
        public PositionsLabControl()
        {
            InitializeControls();
        }

        private void InitializeControls()
        {
            this.BackColor = Color.White;
            this.Dock = DockStyle.Fill;
            this.AutoScroll = true;

            int y = 10;
            int btnWidth = 200;
            int btnHeight = 30;
            int leftMargin = 15;
            int spacing = 5;

            // --- Lock / Clear Selection ---
            _lblSelectionStatus = new Label
            {
                Text = "Selection: Live",
                Font = new Font("Segoe UI", 8),
                ForeColor = Color.Gray,
                Location = new Point(leftMargin, y),
                AutoSize = true
            };
            this.Controls.Add(_lblSelectionStatus);
            y += _lblSelectionStatus.Height + spacing;

            y = AddButton("Lock Selection", CreateIcon(IconKind.SelectAll), y, leftMargin, 95, btnHeight, BtnLockSelection_Click);
            int clearBtnLeft = leftMargin + 100;
            y -= (btnHeight + 5); // stay on the same row
            y = AddButton("Clear", CreateIcon(IconKind.ClearSelection), y, clearBtnLeft, 95, btnHeight, BtnClearSelection_Click);

            y += 10;

            // --- ALIGN section header ---
            var lblAlign = new Label
            {
                Text = "ALIGN",
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = Color.DarkSlateGray,
                Location = new Point(leftMargin, y),
                AutoSize = true
            };
            this.Controls.Add(lblAlign);
            y += lblAlign.Height + spacing;

            y = AddButton("Align Left", CreateIcon(IconKind.AlignLeft), y, leftMargin, btnWidth, btnHeight, BtnAlignLeft_Click);
            y = AddButton("Align Right", CreateIcon(IconKind.AlignRight), y, leftMargin, btnWidth, btnHeight, BtnAlignRight_Click);
            y = AddButton("Align Top", CreateIcon(IconKind.AlignTop), y, leftMargin, btnWidth, btnHeight, BtnAlignTop_Click);
            y = AddButton("Align Bottom", CreateIcon(IconKind.AlignBottom), y, leftMargin, btnWidth, btnHeight, BtnAlignBottom_Click);
            y = AddButton("Align Center", CreateIcon(IconKind.AlignCenter), y, leftMargin, btnWidth, btnHeight, BtnAlignCenter_Click);
            y = AddButton("Align Radially", CreateIcon(IconKind.AlignRadially), y, leftMargin, btnWidth, btnHeight, BtnAlignRadially_Click);

            y += 10;

            // --- REFERENCE ALIGN section header ---
            var lblRefAlign = new Label
            {
                Text = "REFERENCE ALIGN",
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = Color.DarkSlateGray,
                Location = new Point(leftMargin, y),
                AutoSize = true
            };
            this.Controls.Add(lblRefAlign);
            y += lblRefAlign.Height + spacing;

            y = AddButton("Align Horizontal", CreateIcon(IconKind.AlignHorizontal), y, leftMargin, btnWidth, btnHeight, BtnAlignHorizontal_Click);
            y = AddButton("Align Vertical", CreateIcon(IconKind.AlignVertical), y, leftMargin, btnWidth, btnHeight, BtnAlignVertical_Click);

            y += 10;

            y = AddButton("Swap", CreateIcon(IconKind.Swap), y, leftMargin, btnWidth, btnHeight, BtnSwap_Click);
        }

        private int AddButton(string text, Image icon, int y, int left, int width, int height, EventHandler onClick)
        {
            var btn = new Button
            {
                Text = text,
                Image = icon,
                ImageAlign = ContentAlignment.MiddleLeft,
                TextAlign = ContentAlignment.MiddleLeft,
                TextImageRelation = TextImageRelation.ImageBeforeText,
                Padding = new Padding(4, 0, 0, 0),
                Location = new Point(left, y),
                Size = new Size(width, height),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(230, 230, 250),
                Font = new Font("Segoe UI", 9)
            };
            btn.Click += onClick;
            this.Controls.Add(btn);
            return y + height + 5;
        }

        private Label _lblSelectionStatus;

        private enum IconKind
        {
            AlignLeft, AlignRight, AlignTop, AlignBottom, AlignCenter,
            AlignRadially, AlignHorizontal, AlignVertical,
            Swap, SelectAll, ClearSelection
        }

        private static Image CreateIcon(IconKind kind)
        {
            int s = 16;
            var bmp = new Bitmap(s, s);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);
                var pen = new Pen(Color.FromArgb(60, 60, 60), 1.5f);
                var brush = new SolidBrush(Color.FromArgb(80, 80, 80));

                switch (kind)
                {
                    case IconKind.AlignLeft:
                        g.DrawLine(pen, 2, 1, 2, 14);           // guide line
                        g.FillRectangle(brush, 4, 3, 10, 3);    // bar 1
                        g.FillRectangle(brush, 4, 9, 6, 3);     // bar 2
                        break;

                    case IconKind.AlignRight:
                        g.DrawLine(pen, 13, 1, 13, 14);
                        g.FillRectangle(brush, 2, 3, 10, 3);
                        g.FillRectangle(brush, 6, 9, 6, 3);
                        break;

                    case IconKind.AlignTop:
                        g.DrawLine(pen, 1, 2, 14, 2);
                        g.FillRectangle(brush, 3, 4, 3, 10);
                        g.FillRectangle(brush, 9, 4, 3, 6);
                        break;

                    case IconKind.AlignBottom:
                        g.DrawLine(pen, 1, 13, 14, 13);
                        g.FillRectangle(brush, 3, 2, 3, 10);
                        g.FillRectangle(brush, 9, 6, 3, 6);
                        break;

                    case IconKind.AlignCenter:
                        g.DrawLine(pen, 8, 1, 8, 14);           // vertical center
                        g.FillRectangle(brush, 3, 3, 10, 3);
                        g.FillRectangle(brush, 5, 9, 6, 3);
                        break;

                    case IconKind.AlignRadially:
                        // Origin / reference dot at centre
                        g.FillEllipse(brush, 6, 6, 4, 4);
                        // Radius circle (dashed guide)
                        var radDash = new Pen(Color.FromArgb(120, 120, 120), 1f) { DashStyle = DashStyle.Dot };
                        g.DrawEllipse(radDash, 2, 2, 12, 12);
                        radDash.Dispose();
                        // Distance-setter: filled dot on the radius (top)
                        g.FillEllipse(brush, 6, 0, 4, 4);
                        // Shapes to align: hollow dots on the same radius (right & left)
                        g.DrawEllipse(pen, 12, 6, 3, 3);
                        g.DrawEllipse(pen, 1, 6, 3, 3);
                        break;

                    // AlignHorizontal: reference shape on left with centre line; other shapes snapping to it
                    case IconKind.AlignHorizontal:
                        // horizontal center guide line
                        g.DrawLine(pen, 1, 8, 14, 8);
                        // reference shape (left, straddles the line)
                        g.FillRectangle(brush, 1, 5, 5, 6);
                        // other shape snapping to the same Y centre
                        g.DrawRectangle(pen, 8, 5, 5, 6);
                        break;

                    // AlignVertical: reference shape on top with centre line; other shapes snapping to it
                    case IconKind.AlignVertical:
                        // vertical center guide line
                        g.DrawLine(pen, 8, 1, 8, 14);
                        // reference shape (top, straddles the line)
                        g.FillRectangle(brush, 5, 1, 6, 5);
                        // other shape snapping to the same X centre
                        g.DrawRectangle(pen, 5, 8, 6, 5);
                        break;

                    case IconKind.Swap:
                        g.DrawLine(pen, 2, 5, 13, 5);           // top arrow line
                        g.DrawLine(pen, 10, 2, 13, 5);          // top arrowhead
                        g.DrawLine(pen, 10, 8, 13, 5);
                        g.DrawLine(pen, 13, 11, 2, 11);         // bottom arrow line
                        g.DrawLine(pen, 5, 8, 2, 11);           // bottom arrowhead
                        g.DrawLine(pen, 5, 14, 2, 11);
                        break;

                    case IconKind.SelectAll:
                        g.DrawRectangle(pen, 1, 1, 13, 13);     // selection rect
                        var dash = new Pen(Color.FromArgb(80, 80, 80), 1f) { DashStyle = DashStyle.Dot };
                        g.DrawRectangle(dash, 1, 1, 13, 13);
                        g.FillRectangle(brush, 4, 4, 4, 4);     // shape 1
                        g.FillRectangle(brush, 9, 9, 4, 4);     // shape 2
                        dash.Dispose();
                        break;

                    case IconKind.ClearSelection:
                        var redPen = new Pen(Color.FromArgb(180, 40, 40), 2f);
                        g.DrawLine(redPen, 3, 3, 12, 12);       // X mark
                        g.DrawLine(redPen, 12, 3, 3, 12);
                        redPen.Dispose();
                        break;
                }

                pen.Dispose();
                brush.Dispose();
            }
            return bmp;
        }

        private void BtnAlignLeft_Click(object sender, EventArgs e)
        {
            ExecuteAction(PositionLabService.AlignLeft);
        }

        private void BtnAlignRight_Click(object sender, EventArgs e)
        {
            ExecuteAction(PositionLabService.AlignRight);
        }

        private void BtnAlignTop_Click(object sender, EventArgs e)
        {
            ExecuteAction(PositionLabService.AlignTop);
        }

        private void BtnAlignBottom_Click(object sender, EventArgs e)
        {
            ExecuteAction(PositionLabService.AlignBottom);
        }

        private void BtnAlignCenter_Click(object sender, EventArgs e)
        {
            ExecuteAction(PositionLabService.AlignCenter);
        }

        private void BtnAlignHorizontal_Click(object sender, EventArgs e)
        {
            ExecuteAction(PositionLabService.AlignHorizontal);
        }

        private void BtnAlignVertical_Click(object sender, EventArgs e)
        {
            ExecuteAction(PositionLabService.AlignVertical);
        }

        private void BtnSwap_Click(object sender, EventArgs e)
        {
            ExecuteAction(PositionLabService.Swap);
        }

        private void BtnAlignRadially_Click(object sender, EventArgs e)
        {
            ExecuteAction(PositionLabService.AlignRadially);
        }

        private void BtnLockSelection_Click(object sender, EventArgs e)
        {
            try
            {
                int count = PositionLabService.CaptureSelection(Globals.ThisAddIn.Application);
                if (count > 0)
                {
                    var names = PositionLabService.GetCapturedNames();
                    string order = string.Join(", ", names);
                    _lblSelectionStatus.Text = $"Locked ({count}): {order}";
                    _lblSelectionStatus.ForeColor = Color.FromArgb(0, 120, 60);
                }
                else
                {
                    _lblSelectionStatus.Text = "Selection: Live (none captured)";
                    _lblSelectionStatus.ForeColor = Color.Gray;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lock Selection error: " + ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnClearSelection_Click(object sender, EventArgs e)
        {
            PositionLabService.ClearSelection();
            _lblSelectionStatus.Text = "Selection: Live";
            _lblSelectionStatus.ForeColor = Color.Gray;
        }

        private void ExecuteAction(Action<Microsoft.Office.Interop.PowerPoint.Application> action)
        {
            try
            {
                action(Globals.ThisAddIn.Application);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Positions Lab error: " + ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            // 
            // PositionsLabControl
            // 
            this.Name = "PositionsLabControl";
            this.Size = new System.Drawing.Size(164, 161);
            this.ResumeLayout(false);

        }
    }
}