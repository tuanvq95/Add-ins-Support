using System;
using System.Collections.Generic;
using System.Windows.Forms;
using ExcelAddIn.Features;

namespace ExcelAddIn.Forms
{
  /// <summary>
  /// Dialog chọn macro VBA để nhúng vào workbook.
  /// </summary>
  public sealed class MacroSelectorForm : Form
  {
    private CheckedListBox _listMacros;
    private Label _lblDesc;
    private Button _btnOk, _btnCancel, _btnSelectAll, _btnClearAll;

    public IList<MacroDefinition> SelectedMacros { get; private set; } = new List<MacroDefinition>();

    public MacroSelectorForm()
    {
      BuildUI();
    }

    private void BuildUI()
    {
      Text = "Chọn Macro VBA Để Nhúng";
      Size = new System.Drawing.Size(500, 420);
      StartPosition = FormStartPosition.CenterScreen;
      FormBorderStyle = FormBorderStyle.FixedDialog;
      MaximizeBox = false;
      MinimizeBox = false;

      var lblTitle = new Label
      {
        Text = "Chọn các macro bạn muốn nhúng vào workbook hiện tại:",
        AutoSize = true,
        Font = new System.Drawing.Font("Segoe UI", 9f),
        Location = new System.Drawing.Point(12, 12)
      };

      _listMacros = new CheckedListBox
      {
        Location = new System.Drawing.Point(12, 36),
        Size = new System.Drawing.Size(460, 220),
        CheckOnClick = true,
        IntegralHeight = false
      };

      foreach (var macro in VbaMacroManager.AvailableMacros)
        _listMacros.Items.Add(macro, false);

      _listMacros.SelectedIndexChanged += ListMacros_SelectedIndexChanged;

      _lblDesc = new Label
      {
        Location = new System.Drawing.Point(12, 265),
        Size = new System.Drawing.Size(460, 50),
        Text = "← Chọn một macro để xem mô tả",
        ForeColor = System.Drawing.Color.Gray,
        Font = new System.Drawing.Font("Segoe UI", 8.5f, System.Drawing.FontStyle.Italic)
      };

      _btnSelectAll = new Button
      {
        Text = "Chọn Tất Cả",
        Location = new System.Drawing.Point(12, 325),
        Width = 110,
        Height = 28
      };
      _btnSelectAll.Click += (s, e) =>
      {
        for (int i = 0; i < _listMacros.Items.Count; i++)
          _listMacros.SetItemChecked(i, true);
      };

      _btnClearAll = new Button
      {
        Text = "Bỏ Chọn",
        Location = new System.Drawing.Point(130, 325),
        Width = 90,
        Height = 28
      };
      _btnClearAll.Click += (s, e) =>
      {
        for (int i = 0; i < _listMacros.Items.Count; i++)
          _listMacros.SetItemChecked(i, false);
      };

      _btnOk = new Button
      {
        Text = "Nhúng Macro",
        Location = new System.Drawing.Point(292, 325),
        Width = 110,
        Height = 28,
        DialogResult = DialogResult.None
      };
      _btnOk.Click += BtnOk_Click;

      _btnCancel = new Button
      {
        Text = "Hủy",
        Location = new System.Drawing.Point(410, 325),
        Width = 64,
        Height = 28,
        DialogResult = DialogResult.Cancel
      };

      Controls.AddRange(new Control[]
      {
                lblTitle, _listMacros, _lblDesc,
                _btnSelectAll, _btnClearAll, _btnOk, _btnCancel
      });

      AcceptButton = _btnOk;
      CancelButton = _btnCancel;
    }

    private void ListMacros_SelectedIndexChanged(object sender, EventArgs e)
    {
      if (_listMacros.SelectedItem is MacroDefinition macro)
        _lblDesc.Text = macro.Description;
    }

    private void BtnOk_Click(object sender, EventArgs e)
    {
      var selected = new List<MacroDefinition>();
      foreach (MacroDefinition macro in _listMacros.CheckedItems)
        selected.Add(macro);

      if (selected.Count == 0)
      {
        MessageBox.Show(
            "Vui lòng chọn ít nhất một macro.",
            "Chưa Chọn Macro",
            MessageBoxButtons.OK,
            MessageBoxIcon.Warning);
        return;
      }

      SelectedMacros = selected;
      DialogResult = DialogResult.OK;
      Close();
    }
  }
}
