using LocalDocAI.Models;
using System;
using System.Drawing;
using System.Windows.Forms;

namespace LocalDocAI.UI.Controls
{
    public partial class SuggestionCard : UserControl
    {
        public event EventHandler<Suggestion> TrackChangeClicked;
        public event EventHandler<Suggestion> CommentClicked;
        public event EventHandler<Suggestion> HighlightClicked;
        public event EventHandler<Suggestion> IgnoreClicked;

        private Suggestion _suggestion;

        public SuggestionCard(Suggestion suggestion)
        {
            _suggestion = suggestion;
            InitializeComponent();
            BindData();
        }

        private void BindData()
        {
            lblType.Text = GetTypeLabel(_suggestion.Type);
            lblExplanation.Text = _suggestion.Explanation;

            if (!string.IsNullOrEmpty(_suggestion.RangeText))
                lblOriginal.Text = "\"" + Truncate(_suggestion.RangeText, 60) + "\"";

            if (!string.IsNullOrEmpty(_suggestion.ReplacementText))
                lblReplacement.Text = "→ \"" + Truncate(_suggestion.ReplacementText, 60) + "\"";
            else
                lblReplacement.Visible = false;

            // Severity color
            panelSeverity.BackColor = GetSeverityColor(_suggestion.Severity);

            // Show/hide buttons based on action
            bool canApply = _suggestion.Action == SuggestionAction.Replace
                || _suggestion.Action == SuggestionAction.Delete;
            btnTrack.Visible = canApply;
        }

        private string GetTypeLabel(string type)
        {
            switch (type)
            {
                case "spelling": return "Chính tả";
                case "punctuation": return "Dấu câu";
                case "spacing": return "Khoảng trắng";
                case "repeated_word": return "Lặp từ";
                case "terminology": return "Thuật ngữ";
                case "capitalization": return "Viết hoa";
                case "date_logic": return "Ngày tháng";
                case "numbering": return "Đánh số";
                case "cross_reference": return "Tham chiếu";
                case "numeric": return "Số liệu";
                case "passive_voice": return "Câu bị động";
                case "legal_risk": return "Rủi ro pháp lý";
                case "placeholder": return "Placeholder";
                default: return type ?? "Lỗi";
            }
        }

        private Color GetSeverityColor(SuggestionSeverity sev)
        {
            switch (sev)
            {
                case SuggestionSeverity.Critical: return Color.FromArgb(220, 38, 38);
                case SuggestionSeverity.High: return Color.FromArgb(234, 88, 12);
                case SuggestionSeverity.Medium: return Color.FromArgb(202, 138, 4);
                default: return Color.FromArgb(22, 163, 74);
            }
        }

        private string Truncate(string s, int max) =>
            s?.Length > max ? s.Substring(0, max) + "…" : s;

        private void btnTrack_Click(object sender, EventArgs e) =>
            TrackChangeClicked?.Invoke(this, _suggestion);

        private void btnComment_Click(object sender, EventArgs e) =>
            CommentClicked?.Invoke(this, _suggestion);

        private void btnHighlight_Click(object sender, EventArgs e) =>
            HighlightClicked?.Invoke(this, _suggestion);

        private void btnIgnore_Click(object sender, EventArgs e) =>
            IgnoreClicked?.Invoke(this, _suggestion);
    }
}
