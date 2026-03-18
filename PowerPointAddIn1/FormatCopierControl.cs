using System.Drawing;
using System.Windows.Forms;

namespace PowerPointAddIn1
{
    public class FormatCopierControl : UserControl
    {
        public FormatCopierControl()
        {
            BackColor = Color.White;
            Dock = DockStyle.Fill;

            var label = new Label
            {
                Text = "Format Copier",
                AutoSize = true,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = Color.DarkSlateGray,
                Location = new Point(12, 12)
            };

            Controls.Add(label);
        }
    }
}
