using LocalDocAI.AI;
using LocalDocAI.Persistence;
using System.Drawing;
using System.Windows.Forms;

namespace LocalDocAI.UI
{
    public partial class SettingsDialog : Form
    {
        private AppSettings _settings;
        private SettingsService _service;

        public SettingsDialog()
        {
            _service = ThisAddIn.Instance.Settings;
            _settings = _service.Current;
            InitializeComponent();
            LoadValues();
        }

        private void LoadValues()
        {
            txtBaseUrl.Text = _settings.LmStudioBaseUrl;
            txtModel.Text = _settings.ModelName;
            nudTimeout.Value = _settings.TimeoutSeconds;
            nudMaxTokens.Value = _settings.MaxTokens;
            nudTemperature.Value = (decimal)_settings.Temperature;
            chkTrackChanges.Checked = _settings.UseTrackChangesDefault;
            chkConfirmBefore.Checked = _settings.ShowConfirmBeforeApply;
            txtSkillsDir.Text = _settings.SkillsDirectory;
            txtAuthor.Text = _settings.AuthorName;
        }

        private async void btnTest_Click(object sender, System.EventArgs e)
        {
            btnTest.Enabled = false;
            lblTestResult.Text = "Đang kiểm tra...";
            lblTestResult.ForeColor = Color.Gray;

            var testSettings = new SettingsService();
            testSettings.Current.LmStudioBaseUrl = txtBaseUrl.Text.Trim();
            testSettings.Current.TimeoutSeconds = (int)nudTimeout.Value;

            if (!testSettings.IsValidLocalUrl(testSettings.Current.LmStudioBaseUrl))
            {
                lblTestResult.Text = "URL không hợp lệ (chỉ cho phép localhost/127.0.0.1)";
                lblTestResult.ForeColor = Color.FromArgb(220, 38, 38);
                btnTest.Enabled = true;
                return;
            }

            var client = new LmStudioClient(testSettings);
            bool ok = await client.TestConnectionAsync();
            lblTestResult.Text = ok ? "✓ Kết nối thành công" : "✗ Không kết nối được";
            lblTestResult.ForeColor = ok ? Color.FromArgb(22, 163, 74) : Color.FromArgb(220, 38, 38);

            if (ok)
            {
                var models = await client.ListModelsAsync();
                if (models.Count > 0)
                {
                    cmbModels.Items.Clear();
                    foreach (var m in models) cmbModels.Items.Add(m);
                    cmbModels.Visible = true;
                    lblModels.Visible = true;
                    if (models.Count > 0) cmbModels.SelectedIndex = 0;
                }
            }

            btnTest.Enabled = true;
        }

        private void btnSave_Click(object sender, System.EventArgs e)
        {
            if (!_service.IsValidLocalUrl(txtBaseUrl.Text.Trim()))
            {
                MessageBox.Show("URL không hợp lệ. Chỉ cho phép localhost hoặc 127.0.0.1.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            _settings.LmStudioBaseUrl = txtBaseUrl.Text.Trim();
            _settings.ModelName = cmbModels.Visible && cmbModels.SelectedItem != null
                ? cmbModels.SelectedItem.ToString()
                : txtModel.Text.Trim();
            _settings.TimeoutSeconds = (int)nudTimeout.Value;
            _settings.MaxTokens = (int)nudMaxTokens.Value;
            _settings.Temperature = (double)nudTemperature.Value;
            _settings.UseTrackChangesDefault = chkTrackChanges.Checked;
            _settings.ShowConfirmBeforeApply = chkConfirmBefore.Checked;
            _settings.SkillsDirectory = txtSkillsDir.Text.Trim();
            _settings.AuthorName = txtAuthor.Text.Trim();

            _service.Save();
            DialogResult = DialogResult.OK;
            Close();
        }

        private void btnCancel_Click(object sender, System.EventArgs e) => Close();
    }
}
