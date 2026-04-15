using System;
using System.Drawing;
using System.Windows.Forms;
using AddinsSupport.Features;

namespace AddinsSupport.Forms
{
  /// <summary>
  /// Dialog cài đặt tô màu:
  ///   • Chọn màu nền (ColorDialog)
  ///   • Bật/tắt giới hạn cột và nhập cột bắt đầu / kết thúc
  /// Kết quả được lưu vào <see cref="ColorSelectionSettings"/> khi nhấn OK.
  /// </summary>
  public class ColorSelectionSettingsForm : Form
  {
    private CheckBox chkUseColumnRange;
    private Label lblFrom;
    private NumericUpDown nudColFrom;
    private Label lblTo;
    private NumericUpDown nudColTo;
    private Label lblColor;
    private Panel pnlColorPreview;
    private Button btnPickColor;
    private Label lblSheetNormalize;
    private Label lblZoom;
    private NumericUpDown nudZoom;
    private Button btnOK;
    private Button btnCancel;

    public ColorSelectionSettingsForm()
    {
      BuildUI();
      LoadSettings();
    }

    // ─── Xây dựng UI ─────────────────────────────────────────────────────

    private void BuildUI()
    {
      this.Text = "Cài Đặt Tô Màu & Chuẩn Hóa Sheet";
      this.FormBorderStyle = FormBorderStyle.FixedDialog;
      this.MaximizeBox = false;
      this.MinimizeBox = false;
      this.StartPosition = FormStartPosition.CenterScreen;
      this.ClientSize = new Size(340, 215);
      this.Font = new Font("Segoe UI", 9f);

      // ── Checkbox chọn giới hạn cột ──────────────────────────────────
      chkUseColumnRange = new CheckBox
      {
        Text = "Giới hạn cột tô màu",
        Location = new Point(12, 14),
        AutoSize = true
      };
      chkUseColumnRange.CheckedChanged += (s, e) => RefreshColumnControls();

      // ── Dải cột ─────────────────────────────────────────────────────
      lblFrom = new Label
      {
        Text = "Từ cột:",
        Location = new Point(28, 46),
        AutoSize = true,
        TextAlign = System.Drawing.ContentAlignment.MiddleLeft
      };
      nudColFrom = new NumericUpDown
      {
        Location = new Point(88, 43),
        Width = 65,
        Minimum = 1,
        Maximum = 16384,
        Value = 1
      };

      lblTo = new Label
      {
        Text = "Đến cột:",
        Location = new Point(168, 46),
        AutoSize = true,
        TextAlign = System.Drawing.ContentAlignment.MiddleLeft
      };
      nudColTo = new NumericUpDown
      {
        Location = new Point(228, 43),
        Width = 65,
        Minimum = 1,
        Maximum = 16384,
        Value = 10
      };

      // ── Màu nền ─────────────────────────────────────────────────────
      lblColor = new Label
      {
        Text = "Màu nền:",
        Location = new Point(12, 85),
        AutoSize = true
      };
      pnlColorPreview = new Panel
      {
        Location = new Point(88, 82),
        Size = new Size(32, 22),
        BorderStyle = BorderStyle.FixedSingle
      };
      btnPickColor = new Button
      {
        Text = "Chọn Màu...",
        Location = new Point(130, 80),
        Width = 100
      };
      btnPickColor.Click += BtnPickColor_Click;

      // ── Chuẩn hóa Sheet ────────────────────────────────────────
      lblSheetNormalize = new Label
      {
        Text = "───  Chuẩn Hóa Sheet  ───",
        Location = new Point(12, 118),
        AutoSize = true,
        ForeColor = System.Drawing.SystemColors.GrayText
      };

      lblZoom = new Label
      {
        Text = "Zoom chuẩn hóa (%):",
        Location = new Point(12, 148),
        AutoSize = true,
        TextAlign = System.Drawing.ContentAlignment.MiddleLeft
      };
      nudZoom = new NumericUpDown
      {
        Location = new Point(148, 145),
        Width = 75,
        Minimum = 10,
        Maximum = 400,
        Value = 100
      };

      // ── OK / Huỷ ────────────────────────────────────────────
      btnOK = new Button
      {
        Text = "OK",
        Location = new Point(148, 175),
        Width = 80,
        DialogResult = DialogResult.OK
      };
      btnOK.Click += BtnOK_Click;

      btnCancel = new Button
      {
        Text = "Huỷ",
        Location = new Point(244, 175),
        Width = 80,
        DialogResult = DialogResult.Cancel
      };

      this.AcceptButton = btnOK;
      this.CancelButton = btnCancel;

      this.Controls.AddRange(new Control[]
      {
                chkUseColumnRange,
                lblFrom, nudColFrom,
                lblTo, nudColTo,
                lblColor, pnlColorPreview, btnPickColor,
                lblSheetNormalize, lblZoom, nudZoom,
                btnOK, btnCancel
      });
    }

    // ─── Load / Save settings ─────────────────────────────────────────────

    private void LoadSettings()
    {
      chkUseColumnRange.Checked = ColorSelectionSettings.UseColumnRange;
      nudColFrom.Value = Clamp(ColorSelectionSettings.ColFrom, 1, 16384);
      nudColTo.Value = Clamp(ColorSelectionSettings.ColTo, 1, 16384);
      pnlColorPreview.BackColor = ColorSelectionSettings.FillColor;
      nudZoom.Value = Clamp(ColorSelectionSettings.SheetZoomPercent, 10, 400);
      RefreshColumnControls();
    }

    private void RefreshColumnControls()
    {
      bool on = chkUseColumnRange.Checked;
      lblFrom.Enabled = on;
      nudColFrom.Enabled = on;
      lblTo.Enabled = on;
      nudColTo.Enabled = on;
    }

    // ─── Event handlers ──────────────────────────────────────────────────

    private void BtnPickColor_Click(object sender, EventArgs e)
    {
      using (var dlg = new ColorDialog { Color = pnlColorPreview.BackColor, FullOpen = true })
      {
        if (dlg.ShowDialog() == DialogResult.OK)
          pnlColorPreview.BackColor = dlg.Color;
      }
    }

    private void BtnOK_Click(object sender, EventArgs e)
    {
      if (chkUseColumnRange.Checked && nudColFrom.Value > nudColTo.Value)
      {
        MessageBox.Show(
            "Cột bắt đầu không được lớn hơn cột kết thúc.",
            "Cài Đặt Tô Màu",
            MessageBoxButtons.OK,
            MessageBoxIcon.Warning);
        this.DialogResult = DialogResult.None;
        return;
      }

      ColorSelectionSettings.UseColumnRange = chkUseColumnRange.Checked;
      ColorSelectionSettings.ColFrom = (int)nudColFrom.Value;
      ColorSelectionSettings.ColTo = (int)nudColTo.Value;
      ColorSelectionSettings.FillColor = pnlColorPreview.BackColor;
      ColorSelectionSettings.SheetZoomPercent = (int)nudZoom.Value;
    }

    // ─── Helper ──────────────────────────────────────────────────────────

    private static int Clamp(int value, int min, int max)
        => value < min ? min : value > max ? max : value;
  }
}
