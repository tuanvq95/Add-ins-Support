using System;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddIn.Features
{
  /// <summary>
  /// Quản lý Hyperlink: phát hiện URL và tự động chuyển thành hyperlink có thể click.
  /// </summary>
  public static class HyperlinkManager
  {
    // Regex phát hiện URL bắt đầu bằng http/https hoặc www.
    private static readonly Regex UrlRegex = new Regex(
        @"^(https?://|www\.)[^\s""'<>]{2,}$",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    /// <summary>
    /// Kiểm tra chuỗi có phải URL không (dùng cho việc tự động nhận diện khi nhập liệu).
    /// </summary>
    public static bool IsUrl(string value)
    {
      if (string.IsNullOrWhiteSpace(value)) return false;
      return UrlRegex.IsMatch(value.Trim());
    }

    /// <summary>
    /// Quét toàn bộ UsedRange của sheet và chuyển tất cả cell chứa URL thành hyperlink.
    /// </summary>
    public static void AutoAddHyperlinks(Excel.Worksheet ws)
    {
      if (ws == null) throw new ArgumentNullException("ws");

      Excel.Range usedRange = ws.UsedRange;
      int added = 0;
      int skipped = 0;

      foreach (Excel.Range cell in usedRange.Cells)
      {
        string raw = cell.Value2 as string;
        if (string.IsNullOrEmpty(raw)) continue;

        string trimmed = raw.Trim();
        if (!IsUrl(trimmed)) continue;

        // Bỏ qua nếu đã có hyperlink
        if (cell.Hyperlinks.Count > 0) { skipped++; continue; }

        string address = trimmed.StartsWith("www.", StringComparison.OrdinalIgnoreCase)
            ? "http://" + trimmed
            : trimmed;

        try
        {
          ws.Hyperlinks.Add(
              Anchor: cell,
              Address: address,
              TextToDisplay: trimmed);
          added++;
        }
        catch
        {
          // Cell bị khóa, bỏ qua
          skipped++;
        }
      }

      MessageBox.Show(
          $"Hoàn thành!\n• Đã thêm: {added} hyperlink\n• Bỏ qua: {skipped} cell (đã có hoặc bị khóa)",
          "Auto Hyperlink — " + ws.Name,
          MessageBoxButtons.OK,
          MessageBoxIcon.Information);
    }
  }
}
