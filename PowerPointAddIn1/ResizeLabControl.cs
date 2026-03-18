using System.Drawing;
using System.Windows.Forms;

namespace PowerPointAddIn1
{
    public class ResizeLabControl : UserControl
    {
        public ResizeLabControl()
        {
            BackColor = Color.White;
            Dock = DockStyle.Fill;

            var title = new Label
            {
                Text = "RESIZE LAB",
                AutoSize = true,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = Color.DarkSlateGray,
                Location = new Point(12, 12)
            };

            var hint = new Label
            {
                Text = "Resize tools can be added here.",
                AutoSize = true,
                Font = new Font("Segoe UI", 9),
                ForeColor = Color.DimGray,
                Location = new Point(12, 38)
            };

            Controls.Add(title);
            Controls.Add(hint);
        }
    }
}
