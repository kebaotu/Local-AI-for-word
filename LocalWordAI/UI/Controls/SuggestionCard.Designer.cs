using System.Drawing;
using System.Windows.Forms;

namespace LocalWordAI.UI.Controls
{
    partial class SuggestionCard
    {
        private Panel panelSeverity;
        private Label lblType;
        private Label lblOriginal;
        private Label lblReplacement;
        private Label lblExplanation;
        private Button btnTrack;
        private Button btnComment;
        private Button btnHighlight;
        private Button btnIgnore;

        private void InitializeComponent()
        {
            panelSeverity = new Panel();
            lblType = new Label();
            lblOriginal = new Label();
            lblReplacement = new Label();
            lblExplanation = new Label();
            btnTrack = new Button();
            btnComment = new Button();
            btnHighlight = new Button();
            btnIgnore = new Button();

            SuspendLayout();

            // Severity stripe
            panelSeverity.Width = 4;
            panelSeverity.Dock = DockStyle.Left;

            // Type label
            lblType.AutoSize = true;
            lblType.Font = new Font("Segoe UI", 8f, FontStyle.Bold);
            lblType.Location = new Point(10, 4);

            lblOriginal.AutoSize = false;
            lblOriginal.Font = new Font("Segoe UI", 8.5f);
            lblOriginal.ForeColor = Color.FromArgb(100, 100, 100);
            lblOriginal.Location = new Point(10, 22);
            lblOriginal.Size = new Size(300, 16);

            lblReplacement.AutoSize = false;
            lblReplacement.Font = new Font("Segoe UI", 8.5f);
            lblReplacement.ForeColor = Color.FromArgb(22, 163, 74);
            lblReplacement.Location = new Point(10, 38);
            lblReplacement.Size = new Size(300, 16);

            lblExplanation.AutoSize = false;
            lblExplanation.Font = new Font("Segoe UI", 8f);
            lblExplanation.ForeColor = Color.FromArgb(60, 60, 60);
            lblExplanation.Location = new Point(10, 56);
            lblExplanation.Size = new Size(300, 32);

            // Buttons row
            var bY = 92;
            SetupButton(btnTrack, "Track Change", bY, 0);
            SetupButton(btnComment, "Comment", bY, 100);
            SetupButton(btnHighlight, "Highlight", bY, 200);
            SetupButton(btnIgnore, "Bỏ qua", bY, 270);

            btnTrack.Click += btnTrack_Click;
            btnComment.Click += btnComment_Click;
            btnHighlight.Click += btnHighlight_Click;
            btnIgnore.Click += btnIgnore_Click;

            BackColor = Color.White;
            BorderStyle = BorderStyle.FixedSingle;
            Size = new Size(330, 120);
            Margin = new Padding(0, 0, 0, 4);

            Controls.AddRange(new Control[] {
                panelSeverity, lblType, lblOriginal, lblReplacement,
                lblExplanation, btnTrack, btnComment, btnHighlight, btnIgnore
            });

            ResumeLayout(false);
        }

        private void SetupButton(Button btn, string text, int y, int x)
        {
            btn.Text = text;
            btn.Location = new Point(10 + x, y);
            btn.Size = new Size(88, 22);
            btn.FlatStyle = FlatStyle.Flat;
            btn.FlatAppearance.BorderColor = Color.FromArgb(200, 200, 200);
            btn.Font = new Font("Segoe UI", 7.5f);
            btn.BackColor = Color.White;
        }
    }
}
