using System;
using System.Drawing;
using System.Windows.Forms;

namespace PowerPointAddIn1
{
    public class NavBarColorSettings : Form
    {
        private NavBarSettings settings;
        public NavBarSettings Settings => settings;

        public NavBarColorSettings(NavBarSettings currentSettings)
        {
            settings = new NavBarSettings
            {
                BackgroundColor = currentSettings.BackgroundColor,
                SectionNameColor = currentSettings.SectionNameColor,
                CurrentSlideColor = currentSettings.CurrentSlideColor,
                SameSubsectionBorderColor = currentSettings.SameSubsectionBorderColor,
                SameSubsectionFillColor = currentSettings.SameSubsectionFillColor,
                OtherSlidesBorderColor = currentSettings.OtherSlidesBorderColor,
                SubsectionBoxTransparency = currentSettings.SubsectionBoxTransparency,
                SlideShapeType = currentSettings.SlideShapeType,
                ShowSlideNumbers = currentSettings.ShowSlideNumbers,
                SlideNumberColor = currentSettings.SlideNumberColor
            };

            InitializeComponents();
        }

        private void InitializeComponents()
        {
            this.Text = "Customize Navigation Bar";
            this.Size = new Size(450, 620);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;

            int y = 20;
            
            // === SHAPE OPTIONS ===
            Label lblShapeHeader = new Label 
            { 
                Text = "SHAPE OPTIONS", 
                Location = new Point(20, y), 
                Size = new Size(400, 20),
                Font = new Font(this.Font, FontStyle.Bold)
            };
            this.Controls.Add(lblShapeHeader);
            y += 30;

            // Shape Type: Circle or Square
            Label lblShape = new Label { Text = "Shape Type:", Location = new Point(20, y + 5), Size = new Size(180, 20) };
            ComboBox cmbShape = new ComboBox
            {
                Location = new Point(210, y),
                Size = new Size(180, 30),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cmbShape.Items.AddRange(new object[] { "Circle", "Square" });
            cmbShape.SelectedIndex = settings.SlideShapeType == NavBarSettings.ShapeType.Circle ? 0 : 1;
            cmbShape.SelectedIndexChanged += (s, e) => {
                settings.SlideShapeType = cmbShape.SelectedIndex == 0 ? 
                    NavBarSettings.ShapeType.Circle : NavBarSettings.ShapeType.Square;
            };
            this.Controls.AddRange(new Control[] { lblShape, cmbShape });
            y += 45;

            // Show Slide Numbers
            CheckBox chkShowNumbers = new CheckBox
            {
                Text = "Show Slide Numbers",
                Location = new Point(20, y),
                Size = new Size(200, 30),
                Checked = settings.ShowSlideNumbers
            };
            chkShowNumbers.CheckedChanged += (s, e) => settings.ShowSlideNumbers = chkShowNumbers.Checked;
            this.Controls.Add(chkShowNumbers);
            y += 45;

            // === COLOR OPTIONS ===
            Label lblColorHeader = new Label 
            { 
                Text = "COLOR OPTIONS", 
                Location = new Point(20, y), 
                Size = new Size(400, 20),
                Font = new Font(this.Font, FontStyle.Bold)
            };
            this.Controls.Add(lblColorHeader);
            y += 30;

            AddColorOption("Background Color:", settings.BackgroundColor, y, (c) => settings.BackgroundColor = c); y += 45;
            AddColorOption("Section Names:", settings.SectionNameColor, y, (c) => settings.SectionNameColor = c); y += 45;
            AddColorOption("Current Slide:", settings.CurrentSlideColor, y, (c) => settings.CurrentSlideColor = c); y += 45;
            AddColorOption("Slide Numbers:", settings.SlideNumberColor, y, (c) => settings.SlideNumberColor = c); y += 45;
            AddColorOption("Same Subsection Border:", settings.SameSubsectionBorderColor, y, (c) => settings.SameSubsectionBorderColor = c); y += 45;
            AddColorOption("Same Subsection Fill:", settings.SameSubsectionFillColor, y, (c) => settings.SameSubsectionFillColor = c); y += 45;
            AddColorOption("Other Slides Border:", settings.OtherSlidesBorderColor, y, (c) => settings.OtherSlidesBorderColor = c); y += 45;

            // Buttons
            Button btnReset = new Button { Text = "Reset Defaults", Location = new Point(20, y), Size = new Size(120, 30) };
            btnReset.Click += (s, e) => { 
                settings.ResetToDefaults(); 
                this.Close(); 
                new NavBarColorSettings(settings).ShowDialog(); 
            };
            this.Controls.Add(btnReset);

            Button btnOK = new Button { Text = "OK", Location = new Point(250, y), Size = new Size(80, 30), DialogResult = DialogResult.OK };
            Button btnCancel = new Button { Text = "Cancel", Location = new Point(340, y), Size = new Size(80, 30), DialogResult = DialogResult.Cancel };
            this.Controls.AddRange(new Control[] { btnOK, btnCancel });
            this.AcceptButton = btnOK;
            this.CancelButton = btnCancel;
        }

        private void AddColorOption(string label, Color color, int y, Action<Color> updateColor)
        {
            Label lbl = new Label { Text = label, Location = new Point(20, y + 5), Size = new Size(180, 20) };
            Panel preview = new Panel { BackColor = color, Location = new Point(210, y), Size = new Size(80, 30), BorderStyle = BorderStyle.FixedSingle };
            Button btn = new Button { Text = "Change", Location = new Point(300, y), Size = new Size(80, 30) };

            btn.Click += (s, e) => {
                using (ColorDialog dlg = new ColorDialog { Color = preview.BackColor, FullOpen = true })
                {
                    if (dlg.ShowDialog() == DialogResult.OK)
                    {
                        preview.BackColor = dlg.Color;
                        updateColor(dlg.Color);
                    }
                }
            };

            this.Controls.AddRange(new Control[] { lbl, preview, btn });
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            // 
            // NavBarColorSettings
            // 
            this.ClientSize = new System.Drawing.Size(282, 253);
            this.Name = "NavBarColorSettings";
            this.Load += new System.EventHandler(this.NavBarColorSettings_Load);
            this.ResumeLayout(false);

        }

        private void NavBarColorSettings_Load(object sender, EventArgs e)
        {

        }
    }
}
