using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Forms;
using QRCoder;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1
{
    public class QRCodeControl : UserControl
    {
        private TextBox txtContent;
        private NumericUpDown nudSize;
        private Button btnGenerate;
        private Button btnInsert;
        private PictureBox picPreview;
        private Label lblContent;
        private Label lblSize;
        private Bitmap _qrBitmap;

        public QRCodeControl()
        {
            InitializeControls();
        }

        private void InitializeControls()
        {
            this.BackColor = Color.White;
            this.Dock = DockStyle.Fill;
            int pad = 10;
            int y = pad;

            // --- Content label ---
            lblContent = new Label
            {
                Text = "Text or URL:",
                Location = new Point(pad, y),
                AutoSize = true,
                Font = new Font("Segoe UI", 9f, FontStyle.Bold)
            };
            this.Controls.Add(lblContent);
            y += lblContent.Height + 4;

            // --- Content textbox ---
            txtContent = new TextBox
            {
                Location = new Point(pad, y),
                Width = 220,
                Multiline = true,
                Height = 60,
                ScrollBars = ScrollBars.Vertical,
                Font = new Font("Segoe UI", 9f)
            };
            this.Controls.Add(txtContent);
            y += txtContent.Height + 10;

            // --- Size label ---
            lblSize = new Label
            {
                Text = "Pixels per module:",
                Location = new Point(pad, y),
                AutoSize = true,
                Font = new Font("Segoe UI", 9f, FontStyle.Bold)
            };
            this.Controls.Add(lblSize);
            y += lblSize.Height + 4;

            // --- Size picker ---
            nudSize = new NumericUpDown
            {
                Location = new Point(pad, y),
                Width = 80,
                Minimum = 5,
                Maximum = 40,
                Value = 20,
                Font = new Font("Segoe UI", 9f)
            };
            this.Controls.Add(nudSize);
            y += nudSize.Height + 12;

            // --- Generate button ---
            btnGenerate = new Button
            {
                Text = "Generate QR Code",
                Location = new Point(pad, y),
                Width = 220,
                Height = 30,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(0, 120, 215),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9f, FontStyle.Bold)
            };
            btnGenerate.Click += BtnGenerate_Click;
            this.Controls.Add(btnGenerate);
            y += btnGenerate.Height + 12;

            // --- Preview ---
            picPreview = new PictureBox
            {
                Location = new Point(pad, y),
                Size = new Size(220, 220),
                BorderStyle = BorderStyle.FixedSingle,
                SizeMode = PictureBoxSizeMode.Zoom,
                BackColor = Color.WhiteSmoke
            };
            this.Controls.Add(picPreview);
            y += picPreview.Height + 12;

            // --- Insert button ---
            btnInsert = new Button
            {
                Text = "Insert into Slide",
                Location = new Point(pad, y),
                Width = 220,
                Height = 30,
                Enabled = false,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(16, 124, 16),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9f, FontStyle.Bold)
            };
            btnInsert.Click += BtnInsert_Click;
            this.Controls.Add(btnInsert);
        }

        private void BtnGenerate_Click(object sender, EventArgs e)
        {
            string content = txtContent.Text.Trim();
            if (string.IsNullOrEmpty(content))
            {
                MessageBox.Show("Please enter text or a URL.", "No Content",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                using (var qrGenerator = new QRCodeGenerator())
                {
                    QRCodeData qrData = qrGenerator.CreateQrCode(content, QRCodeGenerator.ECCLevel.Q);
                    using (var qrCode = new QRCode(qrData))
                    {
                        _qrBitmap?.Dispose();
                        _qrBitmap = qrCode.GetGraphic((int)nudSize.Value);
                    }
                }

                picPreview.Image = _qrBitmap;
                btnInsert.Enabled = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error generating QR code: " + ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnInsert_Click(object sender, EventArgs e)
        {
            if (_qrBitmap == null)
            {
                MessageBox.Show("Generate a QR code first.", "No QR Code",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                PowerPoint.Application app = Globals.ThisAddIn.Application;

                if (app.Presentations.Count == 0)
                {
                    MessageBox.Show("Please open a presentation first.", "No Presentation",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                PowerPoint.Slide slide = app.ActiveWindow.View.Slide;

                // Save bitmap to a temp file
                string tempPath = Path.Combine(Path.GetTempPath(), "qrcode_temp.png");
                _qrBitmap.Save(tempPath, ImageFormat.Png);

                // Insert as a picture centred on the slide
                float slideWidth = app.ActivePresentation.PageSetup.SlideWidth;
                float slideHeight = app.ActivePresentation.PageSetup.SlideHeight;
                float imgSize = 150f; // points

                PowerPoint.Shape pic = slide.Shapes.AddPicture(
                    tempPath,
                    Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoCTrue,
                    (slideWidth - imgSize) / 2f,
                    (slideHeight - imgSize) / 2f,
                    imgSize,
                    imgSize);

                pic.Name = "QRCode_" + Guid.NewGuid().ToString("N").Substring(0, 8);

                try { File.Delete(tempPath); } catch { }

                MessageBox.Show("QR code inserted into the current slide.", "Success",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error inserting QR code: " + ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _qrBitmap?.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}