using LocalWordAI.AI;
using LocalWordAI.Models;
using LocalWordAI.Persistence;
using LocalWordAI.Rules;
using LocalWordAI.Skills;
using LocalWordAI.UI.Controls;
using LocalWordAI.WordIntegration;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace LocalWordAI.UI
{
    public partial class SidebarPane : UserControl
    {
        // Services
        private LmStudioClient _aiClient;
        private PromptBuilder _promptBuilder;
        private ContextCompressor _compressor;
        private DocumentContextManager _docContext;
        private TrackChangeService _trackChange;
        private CommentService _commentSvc;
        private HighlightService _highlightSvc;
        private RuleEngine _ruleEngine;
        private SkillManager _skillManager;
        private SkillRunner _skillRunner;
        private SettingsService _settings;

        private CancellationTokenSource _cts;
        private List<Suggestion> _currentSuggestions = new List<Suggestion>();
        private List<ChatMessage> _chatHistory = new List<ChatMessage>();

        public SidebarPane()
        {
            InitializeComponent();
            InitServices();
            LoadSkills();
            CheckConnectionAsync();
        }

        private void InitServices()
        {
            _settings = ThisAddIn.Instance.Settings;
            _promptBuilder = new PromptBuilder();
            _compressor = new ContextCompressor();
            _aiClient = new LmStudioClient(_settings);
            _docContext = new DocumentContextManager(ThisAddIn.Instance.WordApp);
            _trackChange = new TrackChangeService(ThisAddIn.Instance.WordApp);
            _commentSvc = new CommentService(ThisAddIn.Instance.WordApp, _settings);
            _highlightSvc = new HighlightService(ThisAddIn.Instance.WordApp);
            _ruleEngine = new RuleEngine();
            _skillManager = new SkillManager(_settings.Current.SkillsDirectory);
            _skillRunner = new SkillRunner(_aiClient, _promptBuilder);
        }

        private void LoadSkills()
        {
            _skillManager.Load();
            cmbSkills.Items.Clear();
            cmbSkills.Items.Add("-- Chọn Skill --");
            foreach (var s in _skillManager.GetAll())
                cmbSkills.Items.Add(s);
            cmbSkills.SelectedIndex = 0;
            cmbSkills.DisplayMember = "Name";
        }

        // Called from Ribbon (may come from COM thread)
        public void TriggerQuickAction(string action)
        {
            if (InvokeRequired) { BeginInvoke((Action)(() => TriggerQuickAction(action))); return; }
            switch (action)
            {
                case "review_selection": _ = RunReviewSelectionAsync(); break;
                case "review_document": _ = RunReviewDocumentAsync(); break;
                case "rewrite": _ = RunRewriteAsync(); break;
                case "check_comments": _ = RunCheckCommentsAsync(); break;
                case "summarize_changes": _ = RunSummarizeChangesAsync(); break;
                case "check_terms": _ = RunCheckTermsAsync(); break;
                case "final_check": _ = RunFinalCheckAsync(); break;
                case "redline": _ = RunRewriteAsync("redline"); break;
            }
        }

        private async void CheckConnectionAsync()
        {
            lblStatus.Text = "Đang kết nối...";
            lblStatus.ForeColor = Color.Gray;
            var ok = await _aiClient.TestConnectionAsync();
            if (ok)
            {
                lblStatus.Text = "● Đã kết nối";
                lblStatus.ForeColor = Color.FromArgb(22, 163, 74);
                // Load model list
                var models = await _aiClient.ListModelsAsync();
                if (models.Count > 0 && string.IsNullOrEmpty(_settings.Current.ModelName))
                {
                    _settings.Current.ModelName = models[0];
                    lblModel.Text = models[0];
                }
                else
                {
                    lblModel.Text = _settings.Current.ModelName;
                }
            }
            else
            {
                lblStatus.Text = "● Mất kết nối";
                lblStatus.ForeColor = Color.FromArgb(220, 38, 38);
                lblModel.Text = "LM Studio chưa chạy";
            }
        }

        // ─── Quick Actions ────────────────────────────────────────────────

        private async void btnReviewSel_Click(object sender, EventArgs e) => await RunReviewSelectionAsync();
        private async void btnReviewDoc_Click(object sender, EventArgs e) => await RunReviewDocumentAsync();
        private async void btnRewrite_Click(object sender, EventArgs e) => await RunRewriteAsync();
        private async void btnShorten_Click(object sender, EventArgs e) => await RunRewriteAsync("Rút gọn, giữ ý chính, không thêm dữ kiện");
        private async void btnFormal_Click(object sender, EventArgs e) => await RunRewriteAsync("Chuyển sang văn phong trang trọng, hành chính");
        private async void btnPassive_Click(object sender, EventArgs e) => await RunCheckPassiveAsync();
        private async void btnComments_Click(object sender, EventArgs e) => await RunCheckCommentsAsync();
        private async void btnChanges_Click(object sender, EventArgs e) => await RunSummarizeChangesAsync();
        private async void btnFinalCheck_Click(object sender, EventArgs e) => await RunFinalCheckAsync();

        private async void btnSend_Click(object sender, EventArgs e)
        {
            var userText = txtInput.Text.Trim();
            if (string.IsNullOrEmpty(userText)) return;
            txtInput.Clear();
            try
            {
                await RunChatAsync(userText);
            }
            catch (Exception ex)
            {
                ShowChatError("Lỗi không xử lý được: " + ex.Message + "\n" + ex.StackTrace);
                SetBusy(false);
            }
        }

        private void txtInput_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter && !e.Shift)
            {
                e.SuppressKeyPress = true;
                btnSend_Click(sender, e);
            }
        }

        private void btnStop_Click(object sender, EventArgs e)
        {
            _cts?.Cancel();
            SetBusy(false);
        }

        private async void btnRunSkill_Click(object sender, EventArgs e)
        {
            if (cmbSkills.SelectedItem is Skill skill)
                await RunSkillAsync(skill);
        }

        // ─── Core AI Operations ──────────────────────────────────────────

        private async Task RunReviewSelectionAsync()
        {
            var ctx = _docContext.GetContext();
            var text = ctx.HasSelection ? ctx.SelectedText : ctx.FullText;
            if (string.IsNullOrWhiteSpace(text)) { ShowError("Không có văn bản để kiểm tra."); return; }

            SetBusy(true, "Đang kiểm tra...");
            AddMessage("system", $"Kiểm tra: {(ctx.HasSelection ? "vùng chọn" : "tài liệu")}");

            try
            {
                _cts = new CancellationTokenSource();

                // Rule-based first
                var ruleResults = _ruleEngine.RunAll(text);

                // AI review
                var compressed = _compressor.CompressForReview(text, ctx.SelectedText);
                var messages = _promptBuilder.BuildReviewPrompt(compressed);
                var parsed = await _aiClient.ChatJsonAsync<SuggestionList>(messages, _cts.Token);
                var aiResults = parsed?.Suggestions ?? new List<Suggestion>();

                // Merge
                _currentSuggestions = MergeSuggestions(ruleResults, aiResults);
                ShowSuggestions(_currentSuggestions);
                AddMessage("assistant", $"Tìm thấy {_currentSuggestions.Count} vấn đề.");
            }
            catch (OperationCanceledException) { AddMessage("assistant", "Đã dừng."); }
            catch (Exception ex) { ShowError("Lỗi: " + ex.Message); }
            finally { SetBusy(false); }
        }

        private async Task RunReviewDocumentAsync()
        {
            SetBusy(true, "Đang kiểm tra tài liệu...");
            var doc = ThisAddIn.Instance.ActiveDoc;
            if (doc == null) { ShowError("Không có tài liệu đang mở."); SetBusy(false); return; }

            try
            {
                _cts = new CancellationTokenSource();
                var fullText = _docContext.GetFullText();
                var ruleResults = _ruleEngine.RunAll(fullText);

                // Chunk if long
                var chunks = _compressor.ChunkDocument(fullText);
                var aiResults = new List<Suggestion>();

                for (int i = 0; i < chunks.Count && !_cts.Token.IsCancellationRequested; i++)
                {
                    SetBusy(true, $"Đang kiểm tra phần {i + 1}/{chunks.Count}...");
                    var messages = _promptBuilder.BuildReviewPrompt(chunks[i]);
                    var parsed = await _aiClient.ChatJsonAsync<SuggestionList>(messages, _cts.Token);
                    if (parsed?.Suggestions != null) aiResults.AddRange(parsed.Suggestions);
                }

                _currentSuggestions = MergeSuggestions(ruleResults, aiResults);
                ShowSuggestions(_currentSuggestions);
                AddMessage("assistant", $"Kiểm tra xong. Tìm thấy {_currentSuggestions.Count} vấn đề.");
            }
            catch (OperationCanceledException) { AddMessage("assistant", "Đã dừng."); }
            catch (Exception ex) { ShowError("Lỗi: " + ex.Message); }
            finally { SetBusy(false); }
        }

        private async Task RunRewriteAsync(string instruction = null)
        {
            var selText = _docContext.GetSelectedText();
            if (string.IsNullOrWhiteSpace(selText))
            {
                ShowError("Hãy chọn đoạn văn bản cần viết lại.");
                return;
            }

            if (string.IsNullOrEmpty(instruction))
            {
                instruction = Microsoft.VisualBasic.Interaction.InputBox(
                    "Yêu cầu viết lại (vd: 'trang trọng hơn', 'ngắn gọn hơn'):",
                    "Rewrite", "");
                if (string.IsNullOrEmpty(instruction)) return;
            }

            SetBusy(true, "Đang viết lại...");
            try
            {
                _cts = new CancellationTokenSource();
                var messages = _promptBuilder.BuildRewritePrompt(selText, instruction);
                var raw = await _aiClient.ChatAsync(messages, _cts.Token);
                var result = JsonOutputParser.Parse<RewriteResult>(raw);

                if (!string.IsNullOrEmpty(result?.RewrittenText))
                {
                    AddMessage("assistant", $"**Bản gốc:**\n{result.OriginalText}\n\n**Viết lại:**\n{result.RewrittenText}\n\n{result.Explanation}");
                    ShowRewriteOptions(selText, result.RewrittenText);
                }
                else
                {
                    AddMessage("assistant", raw);
                }
            }
            catch (OperationCanceledException) { AddMessage("assistant", "Đã dừng."); }
            catch (Exception ex) { ShowError("Lỗi: " + ex.Message); }
            finally { SetBusy(false); }
        }

        private async Task RunCheckCommentsAsync()
        {
            var doc = ThisAddIn.Instance.ActiveDoc;
            if (doc == null) { ShowError("Không có tài liệu."); return; }

            var comments = CommentReader.ReadAll(doc);
            if (comments.Count == 0) { AddMessage("assistant", "Tài liệu không có comment."); return; }

            SetBusy(true, $"Đang phân tích {comments.Count} comments...");
            try
            {
                _cts = new CancellationTokenSource();
                var commentsJson = JsonConvert.SerializeObject(comments, Formatting.Indented);
                var messages = _promptBuilder.BuildCommentPrompt(commentsJson);
                var raw = await _aiClient.ChatAsync(messages, _cts.Token);
                AddMessage("assistant", raw);
            }
            catch (OperationCanceledException) { AddMessage("assistant", "Đã dừng."); }
            catch (Exception ex) { ShowError("Lỗi: " + ex.Message); }
            finally { SetBusy(false); }
        }

        private async Task RunSummarizeChangesAsync()
        {
            var doc = ThisAddIn.Instance.ActiveDoc;
            if (doc == null) { ShowError("Không có tài liệu."); return; }

            var revisions = RevisionReader.ReadAll(doc);
            if (revisions.Count == 0) { AddMessage("assistant", "Tài liệu không có tracked changes."); return; }

            SetBusy(true, $"Đang phân tích {revisions.Count} tracked changes...");
            try
            {
                _cts = new CancellationTokenSource();
                var revisionsJson = JsonConvert.SerializeObject(revisions, Formatting.Indented);
                var messages = _promptBuilder.BuildRevisionPrompt(revisionsJson);
                var raw = await _aiClient.ChatAsync(messages, _cts.Token);
                AddMessage("assistant", raw);
            }
            catch (OperationCanceledException) { AddMessage("assistant", "Đã dừng."); }
            catch (Exception ex) { ShowError("Lỗi: " + ex.Message); }
            finally { SetBusy(false); }
        }

        private async Task RunCheckPassiveAsync()
        {
            var ctx = _docContext.GetContext();
            var text = ctx.HasSelection ? ctx.SelectedText : ctx.FullText;
            if (string.IsNullOrWhiteSpace(text)) { ShowError("Không có văn bản."); return; }

            SetBusy(true, "Đang kiểm tra câu bị động...");
            try
            {
                _cts = new CancellationTokenSource();
                var messages = _promptBuilder.BuildReviewPrompt(text, "Tìm câu bị động, câu mơ hồ trách nhiệm, đề xuất sửa thành câu chủ động");
                var parsed = await _aiClient.ChatJsonAsync<SuggestionList>(messages, _cts.Token);
                var results = parsed?.Suggestions?.Where(s => s.Type == "passive_voice").ToList()
                    ?? new List<Suggestion>();
                _currentSuggestions = results;
                ShowSuggestions(results);
                AddMessage("assistant", $"Tìm thấy {results.Count} câu bị động/mơ hồ.");
            }
            catch (OperationCanceledException) { }
            catch (Exception ex) { ShowError("Lỗi: " + ex.Message); }
            finally { SetBusy(false); }
        }

        private async Task RunCheckTermsAsync()
        {
            var text = _docContext.GetFullText();
            if (string.IsNullOrWhiteSpace(text)) { ShowError("Không có tài liệu."); return; }

            SetBusy(true, "Đang kiểm tra thuật ngữ...");
            try
            {
                _cts = new CancellationTokenSource();
                var messages = _promptBuilder.BuildReviewPrompt(text,
                    "Kiểm tra thuật ngữ chuyên ngành không nhất quán, tên riêng viết khác nhau, defined terms, viết hoa không thống nhất");
                var parsed = await _aiClient.ChatJsonAsync<SuggestionList>(messages, _cts.Token);
                var results = parsed?.Suggestions ?? new List<Suggestion>();
                _currentSuggestions = results;
                ShowSuggestions(results);
                AddMessage("assistant", $"Tìm thấy {results.Count} vấn đề thuật ngữ.");
            }
            catch (OperationCanceledException) { }
            catch (Exception ex) { ShowError("Lỗi: " + ex.Message); }
            finally { SetBusy(false); }
        }

        private async Task RunFinalCheckAsync()
        {
            var doc = ThisAddIn.Instance.ActiveDoc;
            if (doc == null) { ShowError("Không có tài liệu."); return; }

            SetBusy(true, "Đang kiểm tra toàn diện...");
            try
            {
                _cts = new CancellationTokenSource();
                var ctx = _docContext.GetContext(true, true, true, true);
                var allSuggestions = new List<Suggestion>();

                // Rule engine
                allSuggestions.AddRange(_ruleEngine.RunAll(ctx.FullText));

                // AI review
                var messages = _promptBuilder.BuildReviewPrompt(
                    _compressor.CompressForReview(ctx.FullText, null),
                    "Kiểm tra toàn diện: chính tả, logic, số liệu, tham chiếu, placeholder, rủi ro pháp lý, thuật ngữ");
                var parsed = await _aiClient.ChatJsonAsync<SuggestionList>(messages, _cts.Token);
                if (parsed?.Suggestions != null) allSuggestions.AddRange(parsed.Suggestions);

                _currentSuggestions = MergeSuggestions(allSuggestions, new List<Suggestion>());
                ShowSuggestions(_currentSuggestions);
                AddMessage("assistant", $"Final check hoàn tất. Tổng cộng {_currentSuggestions.Count} vấn đề.");
            }
            catch (OperationCanceledException) { }
            catch (Exception ex) { ShowError("Lỗi: " + ex.Message); }
            finally { SetBusy(false); }
        }

        private async Task RunChatAsync(string userText)
        {
            try
            {
                AddMessage("user", userText, switchToChat: true);
                SetBusy(true, "Đang trả lời...");
                _cts = new CancellationTokenSource();
                var ctx = _docContext.GetContext();

                var messages = _promptBuilder.BuildChatWithActionsMessages(userText, ctx, _chatHistory);
                var raw = await _aiClient.ChatAsync(messages, _cts.Token);

                // Try parse as action response
                var actionResp = JsonOutputParser.Parse<ChatActionResponse>(raw);
                string replyText = (actionResp?.Message) ?? raw;

                // Execute word actions if any
                if (actionResp?.Actions != null && actionResp.Actions.Count > 0)
                {
                    SetBusy(true, $"Đang thực hiện {actionResp.Actions.Count} thao tác...");
                    var resultMsg = await ExecuteChatActionsAsync(actionResp.Actions);
                    if (!string.IsNullOrEmpty(resultMsg))
                        replyText += "\n\n" + resultMsg;
                }

                _chatHistory.Add(ChatMessage.User(userText));
                _chatHistory.Add(ChatMessage.Assistant(replyText));
                AddMessage("assistant", replyText, switchToChat: true);
            }
            catch (OperationCanceledException) { SafeAddMessage("assistant", "Đã dừng."); }
            catch (Exception ex)
            {
                WriteLog($"RunChatAsync EXCEPTION: {ex.GetType().Name}: {ex.Message}");
                SafeAddMessage("system", $"⚠ Lỗi: {ex.Message}");
            }
            finally { SetBusy(false); }
        }

        private async Task<string> ExecuteChatActionsAsync(List<ChatAction> actions)
        {
            var doc = ThisAddIn.Instance.ActiveDoc;
            if (doc == null) return "⚠ Không có tài liệu đang mở.";

            // Show confirmation dialog
            var actionList = string.Join("\n", actions.ConvertAll(a =>
                $"• {a.Description ?? a.Type}" + (a.Find != null ? $": \"{Truncate(a.Find, 40)}\"" : "")));

            var confirm = System.Windows.Forms.MessageBox.Show(
                $"Thực hiện {actions.Count} thao tác sau?\n\n{actionList}",
                "Xác nhận thao tác Word",
                System.Windows.Forms.MessageBoxButtons.YesNo,
                System.Windows.Forms.MessageBoxIcon.Question);

            if (confirm != System.Windows.Forms.DialogResult.Yes)
                return "Đã hủy thao tác.";

            var results = new System.Text.StringBuilder();
            int ok = 0, fail = 0;

            foreach (var action in actions)
            {
                try
                {
                    bool success = await Task.Run(() => ExecuteSingleAction(doc, action));
                    if (success) ok++;
                    else { fail++; results.AppendLine($"⚠ Không tìm thấy: \"{Truncate(action.Find, 40)}\""); }
                }
                catch (Exception ex)
                {
                    fail++;
                    results.AppendLine($"⚠ Lỗi ({action.Type}): {ex.Message}");
                }
            }

            if (ok > 0) results.Insert(0, $"✓ Đã thực hiện {ok} thao tác thành công.\n");
            if (fail > 0) results.AppendLine($"✗ {fail} thao tác thất bại.");
            return results.ToString().Trim();
        }

        private bool ExecuteSingleAction(Word.Document doc, ChatAction action)
        {
            switch (action.Type?.ToLower())
            {
                case "insert_at_end":
                    var endRange = doc.Content;
                    endRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    if (!string.IsNullOrEmpty(action.Text))
                    {
                        endRange.InsertParagraphAfter();
                        endRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                        endRange.Text = action.Text;
                    }
                    return true;

                case "insert_at_cursor":
                    var sel = ThisAddIn.Instance.WordApp.Selection;
                    sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    if (!string.IsNullOrEmpty(action.Text))
                        sel.TypeText(action.Text);
                    return true;

                case "track_change_replace":
                    if (string.IsNullOrEmpty(action.Find)) return false;
                    var sug = new Suggestion
                    {
                        RangeText = action.Find,
                        ReplacementText = action.Replacement ?? "",
                        Action = SuggestionAction.Replace
                    };
                    return _trackChange.ApplySuggestion(doc, sug);

                case "direct_replace":
                    if (string.IsNullOrEmpty(action.Find)) return false;
                    bool wasTracking = doc.TrackRevisions;
                    doc.TrackRevisions = false;
                    var sugDirect = new Suggestion
                    {
                        RangeText = action.Find,
                        ReplacementText = action.Replacement ?? "",
                        Action = SuggestionAction.Replace
                    };
                    bool result = _trackChange.ApplySuggestion(doc, sugDirect);
                    doc.TrackRevisions = wasTracking;
                    return result;

                case "add_comment":
                    if (string.IsNullOrEmpty(action.Find)) return false;
                    return _commentSvc.AddCommentOnText(doc, action.Find, action.Comment ?? "");

                case "highlight":
                    if (string.IsNullOrEmpty(action.Find)) return false;
                    return _highlightSvc.HighlightText(doc, action.Find,
                        Word.WdColorIndex.wdYellow);

                default:
                    return false;
            }
        }

        private static string Truncate(string s, int max) =>
            s == null ? "" : (s.Length <= max ? s : s.Substring(0, max) + "...");

        private void SafeAddMessage(string role, string text)
        {
            try { AddMessage(role, text); }
            catch (Exception ex) { WriteLog($"SafeAddMessage failed: {ex.Message}"); }
        }

        private async Task RunSkillAsync(Skill skill)
        {
            SetBusy(true, $"Đang chạy skill: {skill.Name}...");
            try
            {
                _cts = new CancellationTokenSource();
                bool withComments = skill.Context.Contains("comments");
                bool withRevisions = skill.Context.Contains("revisions");
                var ctx = _docContext.GetContext(withComments, withRevisions, true, false);

                var result = await _skillRunner.RunAsync(skill, ctx, _cts.Token);
                if (result.Success)
                {
                    if (result.Suggestions.Count > 0)
                    {
                        _currentSuggestions = result.Suggestions;
                        ShowSuggestions(result.Suggestions);
                        AddMessage("assistant", $"Skill '{skill.Name}' hoàn tất. {result.Suggestions.Count} vấn đề.");
                    }
                    else
                    {
                        AddMessage("assistant", result.Summary ?? $"Skill '{skill.Name}' hoàn tất.");
                    }
                }
                else
                {
                    ShowError("Skill lỗi: " + result.ErrorMessage);
                }
            }
            catch (Exception ex) { ShowError("Lỗi: " + ex.Message); }
            finally { SetBusy(false); }
        }

        // ─── Suggestion Handling ─────────────────────────────────────────

        private void ShowSuggestions(List<Suggestion> suggestions)
        {
            if (InvokeRequired) { Invoke((Action)(() => ShowSuggestions(suggestions))); return; }

            panelSuggestions.SuspendLayout();

            // Dispose old cards to release GDI resources
            foreach (Control c in panelSuggestions.Controls)
                c.Dispose();
            panelSuggestions.Controls.Clear();

            int cardWidth = Math.Max(200, panelSuggestions.ClientSize.Width - 8);
            int yPos = 4;

            foreach (var s in suggestions.Take(30))
            {
                var card = new SuggestionCard(s);
                card.Width = cardWidth;
                card.Location = new Point(4, yPos);   // CRITICAL: must set position explicitly
                card.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
                card.TrackChangeClicked += OnTrackChange;
                card.CommentClicked += OnAddComment;
                card.HighlightClicked += OnHighlight;
                card.IgnoreClicked += OnIgnore;
                panelSuggestions.Controls.Add(card);
                yPos += card.Height + 6;
            }

            // Update scrollable height so AutoScroll works
            panelSuggestions.AutoScrollMinSize = new Size(0, yPos + 4);
            panelSuggestions.ResumeLayout(true);

            if (suggestions.Count > 0)
                tabMain.SelectedTab = tabSuggestions;
        }

        private void ShowRewriteOptions(string original, string rewritten)
        {
            var dlg = new RewriteDialog(original, rewritten);
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                ApplyRewrite(original, rewritten, dlg.UseTrackChanges);
            }
        }

        private void ApplyRewrite(string original, string rewritten, bool trackChanges)
        {
            var doc = ThisAddIn.Instance.ActiveDoc;
            if (doc == null) return;

            var sug = new Suggestion
            {
                RangeText = original,
                ReplacementText = rewritten,
                Action = SuggestionAction.Replace,
                Type = "rewrite"
            };

            if (trackChanges)
                _trackChange.ApplySuggestion(doc, sug);
            else
            {
                // Direct replace without track
                bool oldTrack = doc.TrackRevisions;
                doc.TrackRevisions = false;
                _trackChange.ApplySuggestion(doc, sug);
                doc.TrackRevisions = oldTrack;
            }
        }

        private void OnTrackChange(object sender, Suggestion s)
        {
            var doc = ThisAddIn.Instance.ActiveDoc;
            if (doc == null) return;
            bool ok = _trackChange.ApplySuggestion(doc, s);
            if (ok)
            {
                s.Status = SuggestionStatus.Accepted;
                ((Control)sender).Parent?.Controls.Remove((Control)sender);
                AddMessage("assistant", $"Đã áp dụng Track Change: \"{s.RangeText}\"");
            }
            else
                ShowError("Không thể áp dụng — không tìm thấy đoạn văn.");
        }

        private void OnAddComment(object sender, Suggestion s)
        {
            var doc = ThisAddIn.Instance.ActiveDoc;
            if (doc == null) return;
            bool ok = _commentSvc.AddCommentOnText(doc, s.RangeText, s.Explanation);
            if (ok)
            {
                s.Status = SuggestionStatus.Accepted;
                ((Control)sender).Parent?.Controls.Remove((Control)sender);
                AddMessage("assistant", "Đã thêm comment.");
            }
            else
                ShowError("Không thể thêm comment — không tìm thấy đoạn văn.");
        }

        private void OnHighlight(object sender, Suggestion s)
        {
            var doc = ThisAddIn.Instance.ActiveDoc;
            if (doc == null) return;
            var color = _highlightSvc.GetColorBySeverity(s.SeverityStr);
            bool ok = _highlightSvc.HighlightText(doc, s.RangeText, color);
            if (ok) AddMessage("assistant", "Đã highlight.");
        }

        private void OnIgnore(object sender, Suggestion s)
        {
            s.Status = SuggestionStatus.Ignored;
            if (sender is Control ctrl && ctrl.Parent != null)
            {
                ctrl.Parent.Controls.Remove(ctrl);
                ctrl.Dispose();
                RelayoutSuggestionCards();
            }
        }

        private void RelayoutSuggestionCards()
        {
            panelSuggestions.SuspendLayout();
            int yPos = 4;
            foreach (Control c in panelSuggestions.Controls)
            {
                c.Location = new Point(4, yPos);
                yPos += c.Height + 6;
            }
            panelSuggestions.AutoScrollMinSize = new Size(0, yPos + 4);
            panelSuggestions.ResumeLayout(true);
        }

        private void btnApplyAllSafe_Click(object sender, EventArgs e)
        {
            var doc = ThisAddIn.Instance.ActiveDoc;
            if (doc == null) return;
            var safe = _currentSuggestions
                .Where(s => s.SafeToAutoApply && s.Status == SuggestionStatus.Pending)
                .ToList();
            if (safe.Count == 0) { ShowError("Không có gợi ý nào an toàn để tự động áp dụng."); return; }
            int n = _trackChange.ApplyMultiple(doc, safe);
            foreach (var s in safe) s.Status = SuggestionStatus.Accepted;
            ShowSuggestions(_currentSuggestions.Where(s => s.Status == SuggestionStatus.Pending).ToList());
            AddMessage("assistant", $"Đã áp dụng {n} sửa đổi an toàn.");
        }

        // ─── UI Helpers ───────────────────────────────────────────────────

        private void AddMessage(string role, string text, bool switchToChat = false)
        {
            if (InvokeRequired) { Invoke((Action)(() => AddMessage(role, text, switchToChat))); return; }

            var msg = new ChatMessageControl(role, text);
            msg.Width = Math.Max(200, panelChat.ClientSize.Width - 8);
            msg.Location = new Point(4, GetChatBottom());
            msg.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            panelChat.Controls.Add(msg);
            panelChat.AutoScrollMinSize = new Size(0, msg.Bottom + 8);
            panelChat.ScrollControlIntoView(msg);
            if (switchToChat)
                tabMain.SelectedTab = tabChat;
        }

        private int GetChatBottom()
        {
            int max = 4;
            foreach (Control c in panelChat.Controls)
                if (c.Bottom + 4 > max) max = c.Bottom + 4;
            return max;
        }

        private void ShowError(string msg)
        {
            if (InvokeRequired) { Invoke((Action)(() => ShowError(msg))); return; }
            lblProgress.Text = msg;
            lblProgress.ForeColor = Color.FromArgb(220, 38, 38);
        }

        private void ShowChatError(string msg)
        {
            if (InvokeRequired) { Invoke((Action)(() => ShowChatError(msg))); return; }
            lblProgress.Text = "Lỗi";
            lblProgress.ForeColor = Color.FromArgb(220, 38, 38);
            // Show error as a visible message in the chat area
            var errMsg = new ChatMessageControl("system", "⚠ " + msg);
            errMsg.Width = Math.Max(200, panelChat.ClientSize.Width - 8);
            errMsg.Location = new Point(4, GetChatBottom());
            errMsg.BackColor = Color.FromArgb(254, 242, 242);
            panelChat.Controls.Add(errMsg);
            panelChat.AutoScrollMinSize = new Size(0, errMsg.Bottom + 8);
            panelChat.ScrollControlIntoView(errMsg);
            tabMain.SelectedTab = tabChat;
            WriteLog(msg);
        }

        private static readonly string LogPath = System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "LocalWordAI", "debug.log");

        private static void WriteLog(string msg)
        {
            try
            {
                System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(LogPath));
                System.IO.File.AppendAllText(LogPath,
                    $"[{DateTime.Now:HH:mm:ss}] {msg}{Environment.NewLine}");
            }
            catch { }
        }

        private void SetBusy(bool busy, string msg = null)
        {
            if (InvokeRequired) { Invoke((Action)(() => SetBusy(busy, msg))); return; }
            progressBar.Visible = busy;
            btnStop.Visible = busy;
            btnSend.Enabled = !busy;
            lblProgress.Text = msg ?? "";
            lblProgress.ForeColor = Color.FromArgb(60, 60, 60);
        }

        private List<Suggestion> MergeSuggestions(List<Suggestion> rule, List<Suggestion> ai)
        {
            var all = new List<Suggestion>(rule);
            int nextId = rule.Count + 1;

            foreach (var s in ai)
            {
                // Deduplicate: skip if same rangeText already exists from rules
                bool dup = rule.Any(r => string.Equals(r.RangeText, s.RangeText, StringComparison.OrdinalIgnoreCase)
                    && r.Type == s.Type);
                if (!dup)
                {
                    s.Id = $"AI-{nextId++:D4}";
                    all.Add(s);
                }
            }

            // Sort by severity desc
            return all.OrderByDescending(s => (int)s.Severity).ToList();
        }
    }

    // Small helper for rewrite JSON
    internal class RewriteResult
    {
        [JsonProperty("originalText")]
        public string OriginalText { get; set; }
        [JsonProperty("rewrittenText")]
        public string RewrittenText { get; set; }
        [JsonProperty("explanation")]
        public string Explanation { get; set; }
    }
}
