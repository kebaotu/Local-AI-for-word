using System.Drawing;
using System.Windows.Forms;

namespace LocalDocAI.UI
{
    public class RewriteDialog : Form
    {
        public bool UseTrackChanges { get; private set; } = true;

        public RewriteDialog(string original, string rewritten)
        {
            Text = "Áp dụng bản viết lại";
            Size = new Size(500, 380);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;

            var lblOrig = new Label { Text = "Bản gốc:", Font = new Font("Segoe UI", 9f, FontStyle.Bold), Location = new Point(12, 8), AutoSize = true };
            var txtOrig = new RichTextBox
            {
                Text = original, ReadOnly = true, Location = new Point(12, 28),
                Size = new Size(460, 100), Font = new Font("Segoe UI", 9f), BackColor = Color.FromArgb(254, 242, 242)
            };

            var lblNew = new Label { Text = "Bản viết lại:", Font = new Font("Segoe UI", 9f, FontStyle.Bold), Location = new Point(12, 138), AutoSize = true };
            var txtNew = new RichTextBox
            {
                Text = rewritten, ReadOnly = true, Location = new Point(12, 158),
                Size = new Size(460, 100), Font = new Font("Segoe UI", 9f), BackColor = Color.FromArgb(240, 253, 244)
            };

            var chkTrack = new CheckBox
            {
                Text = "Áp dụng bằng Track Changes (khuyến nghị)",
                Location = new Point(12, 270), AutoSize = true,
                Checked = true, Font = new Font("Segoe UI", 9f)
            };

            var btnOK = new Button
            {
                Text = "Áp dụng", DialogResult = DialogResult.OK,
                Location = new Point(280, 300), Size = new Size(90, 30),
                BackColor = Color.FromArgb(37, 99, 235), ForeColor = Color.White, FlatStyle = FlatStyle.Flat
            };
            var btnCancel = new Button
            {
                Text = "Hủy", DialogResult = DialogResult.Cancel,
                Location = new Point(380, 300), Size = new Size(90, 30), FlatStyle = FlatStyle.Flat
            };

            btnOK.Click += (s, e) => { UseTrackChanges = chkTrack.Checked; };

            Controls.AddRange(new Control[] { lblOrig, txtOrig, lblNew, txtNew, chkTrack, btnOK, btnCancel });
            AcceptButton = btnOK;
            CancelButton = btnCancel;
        }
    }
}
