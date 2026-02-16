using System;
using System.Windows.Forms;

namespace PowerPointAddIn1
{
    public class InputDialog : Form
    {
        private Label lblPrompt;
        private TextBox txtInput;
        private Button btnOK;
        private Button btnCancel;

        public string InputText { get; private set; }

        public InputDialog(string prompt, string title)
        {
            InitializeComponents(prompt, title);
        }

        private void InitializeComponents(string prompt, string title)
        {
            this.Text = title;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.ClientSize = new System.Drawing.Size(350, 120);

            lblPrompt = new Label
            {
                Text = prompt,
                Location = new System.Drawing.Point(10, 10),
                AutoSize = true
            };

            txtInput = new TextBox
            {
                Location = new System.Drawing.Point(10, 35),
                Width = 320
            };

            btnOK = new Button
            {
                Text = "OK",
                DialogResult = DialogResult.OK,
                Location = new System.Drawing.Point(170, 70),
                Width = 75
            };

            btnCancel = new Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Location = new System.Drawing.Point(255, 70),
                Width = 75
            };

            this.Controls.Add(lblPrompt);
            this.Controls.Add(txtInput);
            this.Controls.Add(btnOK);
            this.Controls.Add(btnCancel);

            this.AcceptButton = btnOK;
            this.CancelButton = btnCancel;
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (this.DialogResult == DialogResult.OK)
            {
                InputText = txtInput.Text;
            }
            base.OnFormClosing(e);
        }

        public static string Show(string prompt, string title)
        {
            using (InputDialog dialog = new InputDialog(prompt, title))
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    return dialog.InputText;
                }
                return null;
            }
        }
    }
}
