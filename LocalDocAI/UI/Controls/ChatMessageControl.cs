using System.Drawing;
using System.Windows.Forms;

namespace LocalDocAI.UI.Controls
{
    public partial class ChatMessageControl : UserControl
    {
        public ChatMessageControl(string role, string text)
        {
            InitializeComponent();
            bool isUser = role == "user";

            lblRole.Text = isUser ? "Bạn" : "Local AI";
            lblRole.ForeColor = isUser ? Color.FromArgb(37, 99, 235) : Color.FromArgb(16, 124, 16);
            txtMessage.Text = text;
            BackColor = isUser ? Color.FromArgb(239, 246, 255) : Color.White;
            txtMessage.BackColor = BackColor;
        }
    }
}
