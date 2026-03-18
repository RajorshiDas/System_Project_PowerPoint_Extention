using System;
using System.Drawing;
using System.Windows.Forms;

namespace PowerPointAddIn1
{
    public class ResizeLabSettingsForm : Form
    {
        private readonly RadioButton _firstSelectedRadio;
        private readonly RadioButton _outermostRadio;
        private readonly Button _okButton;
        private readonly Button _cancelButton;

        public ReferenceMode SelectedReferenceMode { get; private set; }

        public ResizeLabSettingsForm(ReferenceMode currentMode)
        {
            Text = "Resize Lab Settings";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(320, 150);

            var optionsGroup = new GroupBox
            {
                Text = "Reference Mode",
                Location = new Point(12, 10),
                Size = new Size(296, 85)
            };

            _firstSelectedRadio = new RadioButton
            {
                Text = "First selected object",
                Location = new Point(12, 24),
                AutoSize = true
            };

            _outermostRadio = new RadioButton
            {
                Text = "Outermost object",
                Location = new Point(12, 48),
                AutoSize = true
            };

            optionsGroup.Controls.Add(_firstSelectedRadio);
            optionsGroup.Controls.Add(_outermostRadio);

            _okButton = new Button
            {
                Text = "OK",
                Size = new Size(90, 28),
                Location = new Point(122, 108),
                DialogResult = DialogResult.OK
            };
            _okButton.Click += OkButton_Click;

            _cancelButton = new Button
            {
                Text = "Cancel",
                Size = new Size(90, 28),
                Location = new Point(218, 108),
                DialogResult = DialogResult.Cancel
            };

            AcceptButton = _okButton;
            CancelButton = _cancelButton;

            Controls.Add(optionsGroup);
            Controls.Add(_okButton);
            Controls.Add(_cancelButton);

            SetInitialSelection(currentMode);
        }

        private void SetInitialSelection(ReferenceMode currentMode)
        {
            SelectedReferenceMode = currentMode;

            switch (currentMode)
            {
                case ReferenceMode.OutermostObject:
                    _outermostRadio.Checked = true;
                    break;
                default:
                    _firstSelectedRadio.Checked = true;
                    break;
            }
        }

        private void OkButton_Click(object sender, EventArgs e)
        {
            SelectedReferenceMode = _outermostRadio.Checked
                ? ReferenceMode.OutermostObject
                : ReferenceMode.FirstSelected;
        }
    }
}
