using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn.Features
{
  /// <summary>
  /// Quản lý tên sheet: đổi tên hàng loạt theo prefix/suffix hoặc theo nội dung cell A1.
  /// </summary>
  public static class SheetNameManager
  {
    /// <summary>
    /// Hiển thị dialog cho phép người dùng đổi tên sheet hàng loạt.
    /// </summary>
    public static void ShowRenameDialog(Excel.Workbook wb)
    {
      using (var form = new Forms.SheetRenameForm(wb))
        form.ShowDialog();
    }

    /// <summary>
    /// Đổi tên mỗi sheet bằng nội dung của cell A1 trong sheet đó.
    /// Tên sheet tối đa 31 ký tự và không chứa ký tự cấm của Excel.
    /// </summary>
    public static void RenameSheetsByCell(Excel.Workbook wb)
    {
      if (wb == null) throw new ArgumentNullException("wb");

      int renamed = 0;
      int skipped = 0;

      foreach (Excel.Worksheet ws in wb.Worksheets)
      {
        try
        {
          string cellValue = ws.Range["A1"].Value2 as string;
          if (string.IsNullOrWhiteSpace(cellValue)) { skipped++; continue; }

          string newName = SanitizeSheetName(cellValue.Trim());
          if (string.IsNullOrEmpty(newName)) { skipped++; continue; }

          // Tránh trùng tên với sheet khác
          newName = EnsureUnique(wb, ws, newName);
          ws.Name = newName;
          renamed++;
        }
        catch
        {
          skipped++;
        }
      }

      MessageBox.Show(
          $"Hoàn thành!\n• Đã đổi tên: {renamed} sheet\n• Bỏ qua: {skipped} sheet",
          "Đổi Tên Theo Cell A1",
          MessageBoxButtons.OK,
          MessageBoxIcon.Information);
    }

    /// <summary>
    /// Loại bỏ các ký tự không hợp lệ trong tên sheet Excel và cắt bớt nếu quá 31 ký tự.
    /// </summary>
    public static string SanitizeSheetName(string name)
    {
      if (string.IsNullOrEmpty(name)) return string.Empty;

      // Ký tự bị cấm trong tên sheet Excel: \ / ? * [ ] :
      char[] forbidden = { '\\', '/', '?', '*', '[', ']', ':' };
      foreach (char c in forbidden)
        name = name.Replace(c, '_');

      // Tối đa 31 ký tự
      return name.Length > 31 ? name.Substring(0, 31) : name;
    }

    /// <summary>
    /// Nếu tên đã tồn tại, thêm số hậu tố để tạo tên duy nhất.
    /// </summary>
    private static string EnsureUnique(Excel.Workbook wb, Excel.Worksheet currentSheet, string name)
    {
      string candidate = name;
      int counter = 2;

      while (true)
      {
        bool conflict = false;
        foreach (Excel.Worksheet ws in wb.Worksheets)
        {
          if (ws == currentSheet) continue;
          if (string.Equals(ws.Name, candidate, StringComparison.OrdinalIgnoreCase))
          {
            conflict = true;
            break;
          }
        }
        if (!conflict) return candidate;

        string suffix = $"_{counter}";
        int maxBase = 31 - suffix.Length;
        candidate = (name.Length > maxBase ? name.Substring(0, maxBase) : name) + suffix;
        counter++;
      }
    }
  }
}
