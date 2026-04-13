using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace AddinsSupport.Forms
{
    /// <summary>
    /// Dialog đổi tên sheet hàng loạt: thêm tiền tố, hậu tố hoặc đánh số thứ tự.
    /// </summary>
    public sealed class SheetRenameForm : Form
    {
        private readonly Excel.Workbook _wb;

        private Label _lblPrefix, _lblSuffix, _lblPreview;
        private TextBox _txtPrefix, _txtSuffix;
        private CheckBox _chkAddNumber;
        private ComboBox _cboNumberPos;
        private ListBox _lstPreview;
        private Button _btnApply, _btnCancel;

        public SheetRenameForm(Excel.Workbook wb)
        {
            _wb = wb ?? throw new ArgumentNullException("wb");
            BuildUI();
            UpdatePreview();
        }

        private void BuildUI()
        {
            Text = "Đổi Tên Sheet Hàng Loạt";
            Size = new System.Drawing.Size(460, 420);
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;

            var panel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(12),
                RowCount = 6,
                ColumnCount = 2
            };

            // Tiền tố
            _lblPrefix = new Label { Text = "Tiền tố (Prefix):", Anchor = AnchorStyles.Left, AutoSize = true };
            _txtPrefix = new TextBox { Width = 220 };
            _txtPrefix.TextChanged += (s, e) => UpdatePreview();

            // Hậu tố
            _lblSuffix = new Label { Text = "Hậu tố (Suffix):", Anchor = AnchorStyles.Left, AutoSize = true };
            _txtSuffix = new TextBox { Width = 220 };
            _txtSuffix.TextChanged += (s, e) => UpdatePreview();

            // Đánh số
            _chkAddNumber = new CheckBox { Text = "Thêm số thứ tự", AutoSize = true };
            _chkAddNumber.CheckedChanged += (s, e) => UpdatePreview();

            _cboNumberPos = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Width = 180
            };
            _cboNumberPos.Items.AddRange(new object[] { "Đặt số ở đầu tên", "Đặt số ở cuối tên" });
            _cboNumberPos.SelectedIndex = 0;
            _cboNumberPos.SelectedIndexChanged += (s, e) => UpdatePreview();

            // Preview
            _lblPreview = new Label { Text = "Xem trước:", AutoSize = true };
            _lstPreview = new ListBox { Height = 150, Width = 400, Dock = DockStyle.Fill };

            // Buttons
            _btnApply = new Button { Text = "Áp Dụng", Width = 120, Height = 32, DialogResult = DialogResult.None };
            _btnApply.Click += BtnApply_Click;
            _btnCancel = new Button { Text = "Hủy", Width = 80, Height = 32, DialogResult = DialogResult.Cancel };

            panel.Controls.Add(_lblPrefix, 0, 0);
            panel.Controls.Add(_txtPrefix, 1, 0);
            panel.Controls.Add(_lblSuffix, 0, 1);
            panel.Controls.Add(_txtSuffix, 1, 1);
            panel.Controls.Add(_chkAddNumber, 0, 2);
            panel.Controls.Add(_cboNumberPos, 1, 2);
            panel.Controls.Add(_lblPreview, 0, 3);
            panel.SetColumnSpan(_lstPreview, 2);
            panel.Controls.Add(_lstPreview, 0, 4);

            var btnPanel = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.Bottom,
                Height = 48,
                Padding = new Padding(8)
            };
            btnPanel.Controls.Add(_btnCancel);
            btnPanel.Controls.Add(_btnApply);

            Controls.Add(panel);
            Controls.Add(btnPanel);
            AcceptButton = _btnApply;
            CancelButton = _btnCancel;
        }

        private void UpdatePreview()
        {
            _lstPreview.Items.Clear();
            int idx = 1;
            foreach (Excel.Worksheet ws in _wb.Worksheets)
            {
                string newName = BuildName(ws.Name, idx);
                _lstPreview.Items.Add($"{ws.Name}  →  {newName}");
                idx++;
            }
        }

        private string BuildName(string original, int index)
        {
            string name = original;
            string prefix = _txtPrefix.Text;
            string suffix = _txtSuffix.Text;
            bool addNum = _chkAddNumber.Checked;
            bool numFirst = _cboNumberPos.SelectedIndex == 0;

            if (addNum && numFirst)
                name = $"{index:D2}_{name}";

            name = prefix + name + suffix;

            if (addNum && !numFirst)
                name = $"{name}_{index:D2}";

            return Features.SheetNameManager.SanitizeSheetName(name);
        }

        private void BtnApply_Click(object sender, EventArgs e)
        {
            int renamed = 0, skipped = 0;
            int idx = 1;

            foreach (Excel.Worksheet ws in _wb.Worksheets)
            {
                string newName = BuildName(ws.Name, idx);
                try
                {
                    ws.Name = newName;
                    renamed++;
                }
                catch
                {
                    skipped++;
                }
                idx++;
            }

            MessageBox.Show(
                $"Hoàn thành!\n• Đã đổi tên: {renamed} sheet\n• Lỗi/bỏ qua: {skipped} sheet",
                "Kết Quả",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);

            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
