using System.Drawing;
using System.Windows.Forms;

namespace LocalDocAI.UI
{
    partial class SettingsDialog
    {
        private TextBox txtBaseUrl;
        private TextBox txtModel;
        private NumericUpDown nudTimeout;
        private NumericUpDown nudMaxTokens;
        private NumericUpDown nudTemperature;
        private CheckBox chkTrackChanges;
        private CheckBox chkConfirmBefore;
        private TextBox txtSkillsDir;
        private TextBox txtAuthor;
        private Button btnTest;
        private Button btnSave;
        private Button btnCancel;
        private Label lblTestResult;
        private ComboBox cmbModels;
        private Label lblModels;
        private ComboBox cmbLanguage;

        private void InitializeComponent()
        {
            Text = LocalDocAI.Persistence.LocalizationService.Get("settingsTitle");
            Size = new Size(440, 520);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;

            int y = 10, lx = 10, cx = 200, w = 210;

            void AddRow(string label, Control ctrl)
            {
                Controls.Add(new Label { Text = label, Location = new Point(lx, y + 3), AutoSize = true, Font = new Font("Segoe UI", 8.5f) });
                ctrl.Location = new Point(cx, y);
                if (ctrl is TextBox tb) { tb.Width = w; }
                Controls.Add(ctrl);
                y += 30;
            }

            txtBaseUrl = new TextBox { Font = new Font("Segoe UI", 8.5f) };
            txtModel = new TextBox { Font = new Font("Segoe UI", 8.5f) };
            nudTimeout = new NumericUpDown { Minimum = 10, Maximum = 600, Value = 120, Width = 80 };
            nudMaxTokens = new NumericUpDown { Minimum = 256, Maximum = 32768, Value = 4096, Width = 80 };
            nudTemperature = new NumericUpDown { Minimum = 0, Maximum = 2, DecimalPlaces = 2, Increment = 0.05m, Value = 0.2m, Width = 80 };
            chkTrackChanges = new CheckBox { Text = LocalDocAI.Persistence.LocalizationService.Get("chkTrackChanges"), Location = new Point(cx, y), AutoSize = true };
            chkConfirmBefore = new CheckBox { Text = LocalDocAI.Persistence.LocalizationService.Get("chkConfirmBefore"), AutoSize = true };
            txtSkillsDir = new TextBox { Font = new Font("Segoe UI", 8.5f) };
            txtAuthor = new TextBox { Font = new Font("Segoe UI", 8.5f) };

            AddRow(LocalDocAI.Persistence.LocalizationService.Get("lblBaseUrl"), txtBaseUrl);
            AddRow(LocalDocAI.Persistence.LocalizationService.Get("lblModel"), txtModel);
            AddRow(LocalDocAI.Persistence.LocalizationService.Get("lblTimeout"), nudTimeout);
            AddRow(LocalDocAI.Persistence.LocalizationService.Get("lblMaxTokens"), nudMaxTokens);
            AddRow(LocalDocAI.Persistence.LocalizationService.Get("lblTemperature"), nudTemperature);

            Controls.Add(chkTrackChanges);
            y += 30;
            chkConfirmBefore.Location = new Point(cx, y);
            Controls.Add(chkConfirmBefore);
            y += 30;

            AddRow(LocalDocAI.Persistence.LocalizationService.Get("lblSkillsDir"), txtSkillsDir);
            AddRow(LocalDocAI.Persistence.LocalizationService.Get("lblAuthor"), txtAuthor);
            y += 5;

            var langLabel = new Label { Text = LocalDocAI.Persistence.LocalizationService.Get("lblLanguage"), Location = new Point(lx, y + 3), AutoSize = true, Font = new Font("Segoe UI", 8.5f) };
            cmbLanguage = new ComboBox { Location = new Point(cx, y), Width = w, DropDownStyle = ComboBoxStyle.DropDownList };
            cmbLanguage.Items.AddRange(new object[] { "Tiếng Việt", "English", "Français" });
            Controls.Add(langLabel);
            Controls.Add(cmbLanguage);
            y += 35;

            btnTest = new Button { Text = LocalDocAI.Persistence.LocalizationService.Get("btnTest"), Location = new Point(10, y), Size = new Size(100, 28), FlatStyle = FlatStyle.Flat };
            lblTestResult = new Label { Location = new Point(120, y + 6), AutoSize = true, Font = new Font("Segoe UI", 8.5f) };
            btnTest.Click += btnTest_Click;
            Controls.Add(btnTest);
            Controls.Add(lblTestResult);
            y += 35;

            lblModels = new Label { Text = LocalDocAI.Persistence.LocalizationService.Get("lblModels"), Location = new Point(lx, y + 3), AutoSize = true, Visible = false };
            cmbModels = new ComboBox { Location = new Point(cx, y), Width = w, DropDownStyle = ComboBoxStyle.DropDownList, Visible = false };
            Controls.Add(lblModels);
            Controls.Add(cmbModels);
            y += 35;

            btnSave = new Button
            {
                Text = LocalDocAI.Persistence.LocalizationService.Get("btnSave"), Location = new Point(240, y), Size = new Size(80, 30),
                BackColor = Color.FromArgb(37, 99, 235), ForeColor = Color.White, FlatStyle = FlatStyle.Flat
            };
            btnCancel = new Button { Text = LocalDocAI.Persistence.LocalizationService.Get("btnCancel"), Location = new Point(330, y), Size = new Size(80, 30), FlatStyle = FlatStyle.Flat };
            btnSave.Click += btnSave_Click;
            btnCancel.Click += btnCancel_Click;
            Controls.Add(btnSave);
            Controls.Add(btnCancel);

            Height = y + 70;
            AutoScaleMode = AutoScaleMode.Font;
        }
    }
}
