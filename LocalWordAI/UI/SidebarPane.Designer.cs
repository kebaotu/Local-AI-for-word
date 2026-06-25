using System.Drawing;
using System.Windows.Forms;

namespace LocalWordAI.UI
{
    partial class SidebarPane
    {
        private Label lblTitle;
        private Label lblModel;
        private Label lblStatus;
        private TabControl tabMain;
        private TabPage tabChat;
        private TabPage tabSuggestions;
        private TabPage tabSkills;
        private Panel panelChat;
        private Panel panelSuggestions;
        private Panel panelQuickActions;
        private FlowLayoutPanel flowQuickActions;
        private TextBox txtInput;
        private Button btnSend;
        private Button btnStop;
        private ProgressBar progressBar;
        private Label lblProgress;
        private Button btnApplyAllSafe;
        private ComboBox cmbSkills;
        private Button btnRunSkill;

        // Quick action buttons
        private Button btnReviewSel;
        private Button btnReviewDoc;
        private Button btnRewrite;
        private Button btnShorten;
        private Button btnFormal;
        private Button btnPassive;
        private Button btnComments;
        private Button btnChanges;
        private Button btnFinalCheck;

        private void InitializeComponent()
        {
            SuspendLayout();

            BackColor = Color.FromArgb(248, 250, 252);

            // ── Header ──
            var pnlHeader = new Panel { Dock = DockStyle.Top, Height = 56, BackColor = Color.FromArgb(30, 41, 59) };

            lblTitle = new Label
            {
                Text = "Local Word AI",
                Font = new Font("Segoe UI", 11f, FontStyle.Bold),
                ForeColor = Color.White,
                Location = new Point(8, 6),
                AutoSize = true
            };

            lblModel = new Label
            {
                Text = "...",
                Font = new Font("Segoe UI", 7.5f),
                ForeColor = Color.FromArgb(148, 163, 184),
                Location = new Point(8, 28),
                AutoSize = true
            };

            lblStatus = new Label
            {
                Text = "Đang kiểm tra...",
                Font = new Font("Segoe UI", 7.5f),
                ForeColor = Color.Gray,
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                AutoSize = true
            };

            pnlHeader.Controls.AddRange(new Control[] { lblTitle, lblModel, lblStatus });
            pnlHeader.Resize += (s, e) => lblStatus.Location = new Point(pnlHeader.Width - lblStatus.Width - 8, 8);

            // ── Quick actions ──
            panelQuickActions = new Panel
            {
                Dock = DockStyle.Top,
                Height = 68,
                BackColor = Color.FromArgb(241, 245, 249),
                Padding = new Padding(4, 4, 4, 0)
            };

            flowQuickActions = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = true,
                AutoScroll = false,
                Padding = new Padding(0)
            };

            btnReviewSel = MakeActionBtn("Review chọn", btnReviewSel_Click);
            btnReviewDoc = MakeActionBtn("Review tài liệu", btnReviewDoc_Click);
            btnRewrite = MakeActionBtn("Rewrite", btnRewrite_Click);
            btnShorten = MakeActionBtn("Rút gọn", btnShorten_Click);
            btnFormal = MakeActionBtn("Trang trọng", btnFormal_Click);
            btnPassive = MakeActionBtn("Câu bị động", btnPassive_Click);
            btnComments = MakeActionBtn("Comments", btnComments_Click);
            btnChanges = MakeActionBtn("Track Changes", btnChanges_Click);
            btnFinalCheck = MakeActionBtn("Final Check", btnFinalCheck_Click, Color.FromArgb(37, 99, 235));

            flowQuickActions.Controls.AddRange(new Control[] {
                btnReviewSel, btnReviewDoc, btnRewrite, btnShorten,
                btnFormal, btnPassive, btnComments, btnChanges, btnFinalCheck
            });

            panelQuickActions.Controls.Add(flowQuickActions);

            // ── Progress ──
            var pnlProgress = new Panel { Dock = DockStyle.Top, Height = 24, BackColor = Color.FromArgb(248, 250, 252) };
            progressBar = new ProgressBar { Style = ProgressBarStyle.Marquee, Dock = DockStyle.Left, Width = 120, Visible = false };
            btnStop = new Button { Text = "■ Stop", Dock = DockStyle.Left, Width = 60, FlatStyle = FlatStyle.Flat, Visible = false, Font = new Font("Segoe UI", 7.5f) };
            lblProgress = new Label { Dock = DockStyle.Fill, Font = new Font("Segoe UI", 7.5f), ForeColor = Color.Gray, TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(4, 0, 0, 0) };
            btnStop.Click += btnStop_Click;
            pnlProgress.Controls.AddRange(new Control[] { lblProgress, btnStop, progressBar });

            // ── Tab control ──
            tabMain = new TabControl { Dock = DockStyle.Fill, Font = new Font("Segoe UI", 8.5f) };
            tabChat = new TabPage { Text = "Chat", Padding = new Padding(0) };
            tabSuggestions = new TabPage { Text = "Gợi ý", Padding = new Padding(0) };
            tabSkills = new TabPage { Text = "Skills", Padding = new Padding(4) };

            // Chat tab
            panelChat = new Panel { Dock = DockStyle.Fill, AutoScroll = true, BackColor = Color.White };
            panelChat.Padding = new Padding(2);
            tabChat.Controls.Add(panelChat);

            // Suggestions tab
            var pnlSugTop = new Panel { Dock = DockStyle.Top, Height = 28, BackColor = Color.White };
            btnApplyAllSafe = new Button
            {
                Text = "Áp dụng tất cả an toàn",
                Dock = DockStyle.Fill,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 8f),
                BackColor = Color.FromArgb(240, 253, 244),
                ForeColor = Color.FromArgb(22, 163, 74)
            };
            btnApplyAllSafe.Click += btnApplyAllSafe_Click;
            pnlSugTop.Controls.Add(btnApplyAllSafe);

            panelSuggestions = new Panel { Dock = DockStyle.Fill, AutoScroll = true, BackColor = Color.FromArgb(248, 250, 252) };
            tabSuggestions.Controls.Add(panelSuggestions);
            tabSuggestions.Controls.Add(pnlSugTop);

            // Skills tab
            var lblSkillsHint = new Label { Text = "Chọn skill:", AutoSize = true, Location = new Point(0, 4), Font = new Font("Segoe UI", 8.5f) };
            cmbSkills = new ComboBox { Location = new Point(0, 22), Width = 280, DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Segoe UI", 8.5f) };
            btnRunSkill = new Button
            {
                Text = "Chạy Skill",
                Location = new Point(0, 50),
                Size = new Size(120, 28),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(37, 99, 235),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 8.5f)
            };
            btnRunSkill.Click += btnRunSkill_Click;
            tabSkills.Controls.AddRange(new Control[] { lblSkillsHint, cmbSkills, btnRunSkill });

            tabMain.TabPages.AddRange(new TabPage[] { tabChat, tabSuggestions, tabSkills });

            // ── Input bar ──
            var pnlInput = new Panel { Dock = DockStyle.Bottom, Height = 56, BackColor = Color.White, Padding = new Padding(4) };
            txtInput = new TextBox
            {
                Multiline = false,
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 9f),
                ForeColor = System.Drawing.Color.Black,
                BorderStyle = BorderStyle.FixedSingle
            };
            txtInput.KeyDown += txtInput_KeyDown;
            btnSend = new Button
            {
                Text = "Gửi",
                Dock = DockStyle.Right,
                Width = 56,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(37, 99, 235),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 8.5f)
            };
            btnSend.Click += btnSend_Click;
            pnlInput.Controls.Add(txtInput);
            pnlInput.Controls.Add(btnSend);

            // ── Assemble ──
            Controls.Add(tabMain);
            Controls.Add(pnlProgress);
            Controls.Add(panelQuickActions);
            Controls.Add(pnlHeader);
            Controls.Add(pnlInput);

            AutoScaleMode = AutoScaleMode.Font;
            Dock = DockStyle.Fill;
            MinimumSize = new Size(300, 400);

            ResumeLayout(false);
        }

        private Button MakeActionBtn(string text, System.EventHandler handler, Color? backColor = null)
        {
            var btn = new Button
            {
                Text = text,
                Height = 24,
                AutoSize = true,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 7.5f),
                BackColor = backColor ?? Color.White,
                ForeColor = backColor.HasValue ? Color.White : Color.FromArgb(30, 41, 59),
                Margin = new Padding(2, 2, 2, 2),
                Padding = new Padding(6, 2, 6, 2)
            };
            btn.FlatAppearance.BorderColor = Color.FromArgb(200, 200, 200);
            btn.Click += handler;
            return btn;
        }
    }
}
