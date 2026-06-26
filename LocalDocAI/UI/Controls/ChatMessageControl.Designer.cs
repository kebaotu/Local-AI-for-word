using System.Drawing;
using System.Windows.Forms;

namespace LocalDocAI.UI.Controls
{
    partial class ChatMessageControl
    {
        private Label lblRole;
        private RichTextBox txtMessage;

        private void InitializeComponent()
        {
            lblRole = new Label();
            txtMessage = new RichTextBox();
            SuspendLayout();

            lblRole.AutoSize = true;
            lblRole.Font = new Font("Segoe UI", 8f, FontStyle.Bold);
            lblRole.Location = new Point(4, 4);

            txtMessage.BorderStyle = BorderStyle.None;
            txtMessage.ReadOnly = true;
            txtMessage.ScrollBars = RichTextBoxScrollBars.None;
            txtMessage.BackColor = Color.White;
            txtMessage.Font = new Font("Segoe UI", 9f);
            txtMessage.Location = new Point(4, 22);
            txtMessage.Width = 300;
            txtMessage.WordWrap = true;

            AutoScaleMode = AutoScaleMode.Font;
            Padding = new Padding(4);
            Controls.Add(lblRole);
            Controls.Add(txtMessage);
            ResumeLayout(false);
        }

        protected override void OnLayout(LayoutEventArgs e)
        {
            base.OnLayout(e);
            if (txtMessage == null) return;
            txtMessage.Width = Width - 12;
            using (var g = CreateGraphics())
            {
                var size = g.MeasureString(txtMessage.Text, txtMessage.Font, txtMessage.Width);
                txtMessage.Height = (int)size.Height + 8;
            }
            Height = txtMessage.Top + txtMessage.Height + 6;
        }
    }
}
