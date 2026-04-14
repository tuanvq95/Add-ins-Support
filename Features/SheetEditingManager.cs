using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using MsoShapeType = Microsoft.Office.Core.MsoShapeType;
using MsoTriState = Microsoft.Office.Core.MsoTriState;

namespace AddinsSupport.Features
{
  /// <summary>
  /// Quản lý các thao tác chỉnh sửa sheet:
  ///   • Tô màu nền vùng chọn
  ///   • Thêm sheet mới theo định dạng mẫu
  ///   • Chuẩn hóa kích thước hình ảnh
  ///   • Unhide tất cả sheet ẩn
  /// </summary>
  public static class SheetEditingManager
  {
    // ─── Hằng số cấu hình ────────────────────────────────────────────────

    /// <summary>Màu nền mặc định khi tô màu vùng chọn: vàng nhạt (RGB #FFFFCC).</summary>
    public const int DEFAULT_FILL_COLOR = 0xCCFFFF; // BGR trong COM = 0xCCFFFF → RGB #FFFFCC

    /// <summary>Chiều rộng chuẩn (đơn vị point) khi chuẩn hóa hình ảnh.</summary>
    public const float STANDARD_IMAGE_WIDTH = 120f;

    /// <summary>Chiều cao chuẩn (đơn vị point) khi chuẩn hóa hình ảnh.</summary>
    public const float STANDARD_IMAGE_HEIGHT = 80f;

    // ─── Tô Màu Vùng Chọn ────────────────────────────────────────────────

    /// <summary>
    /// Tô màu nền cho toàn bộ dòng (hoặc dải cột theo cài đặt) của vùng đang chọn.
    /// <list type="bullet">
    ///   <item>Chọn 1 ô → tô toàn bộ dòng đó.</item>
    ///   <item>Chọn vùng nhiều dòng → tô tất cả các dòng trong vùng.</item>
    ///   <item>Nếu cài đặt "Giới hạn cột" → chỉ tô từ cột ColFrom đến ColTo.</item>
    /// </list>
    /// Màu và phạm vi cột lấy từ <see cref="ColorSelectionSettings"/>.
    /// </summary>
    /// <param name="ws">Sheet đang hoạt động.</param>
    public static void ColorSelection(Excel.Worksheet ws)
    {
      if (ws == null) throw new ArgumentNullException("ws");

      Excel.Range selection = ws.Application.Selection as Excel.Range;
      if (selection == null)
      {
        MessageBox.Show(
            "Vui lòng chọn một ô hoặc vùng ô trước khi thực hiện.",
            "Tô Màu",
            MessageBoxButtons.OK,
            MessageBoxIcon.Warning);
        return;
      }

      int colorBgr = ColorSelectionSettings.FillColorBgr;
      bool limitCols = ColorSelectionSettings.UseColumnRange;
      int colFrom = ColorSelectionSettings.ColFrom;
      int colTo = ColorSelectionSettings.ColTo;

      // Xác định dải dòng từ vùng chọn (không lặp từng ô — nhanh hơn)
      int firstRow = selection.Row;
      int lastRow = firstRow + selection.Rows.Count - 1;

      for (int rowNum = firstRow; rowNum <= lastRow; rowNum++)
      {
        Excel.Range target;
        if (limitCols)
          target = ws.Range[ws.Cells[rowNum, colFrom], ws.Cells[rowNum, colTo]] as Excel.Range;
        else
          target = ws.Rows[rowNum] as Excel.Range;

        if (target != null)
          target.Interior.Color = colorBgr;
      }
    }

    // ─── Thêm Sheet Theo Format ───────────────────────────────────────────

    /// <summary>
    /// Tạo sheet mới bằng cách sao chép sheet cuối cùng trong workbook
    /// (giữ nguyên định dạng, column width, row height, v.v.) rồi xóa nội dung.
    /// Sheet mới được tự động đặt tên theo format "Sheet_yyyyMMdd".
    /// </summary>
    /// <param name="wb">Workbook đang hoạt động.</param>
    public static void AddSheetWithFormat(Excel.Workbook wb)
    {
      if (wb == null) throw new ArgumentNullException("wb");

      int sheetCount = wb.Worksheets.Count;
      Excel.Worksheet templateSheet = wb.Worksheets[sheetCount] as Excel.Worksheet;
      if (templateSheet == null)
      {
        MessageBox.Show(
            "Không tìm thấy sheet mẫu để sao chép.",
            "Thêm Sheet Theo Format",
            MessageBoxButtons.OK,
            MessageBoxIcon.Warning);
        return;
      }

      // Copy sheet cuối → sheet mới xuất hiện sau nó
      templateSheet.Copy(After: templateSheet);

      // Lấy sheet vừa được tạo (luôn là sheet cuối mới nhất)
      Excel.Worksheet newSheet = wb.Worksheets[wb.Worksheets.Count] as Excel.Worksheet;
      if (newSheet == null) return;

      // Xóa nội dung nhưng giữ lại định dạng ô
      newSheet.UsedRange.ClearContents();

      // Đặt tên mặc định; đảm bảo không trùng với sheet đã có
      string baseName = "Sheet_" + DateTime.Now.ToString("yyyyMMdd");
      newSheet.Name = SheetNameManager.EnsureUnique(wb, newSheet, baseName);

      // Kích hoạt sheet mới để người dùng thấy ngay
      newSheet.Activate();

      MessageBox.Show(
          $"Đã tạo sheet mới '{newSheet.Name}'\n"
          + $"từ định dạng của sheet '{templateSheet.Name}'.",
          "Thêm Sheet Theo Format",
          MessageBoxButtons.OK,
          MessageBoxIcon.Information);
    }

    // ─── Chuẩn Hóa Kích Thước Hình Ảnh ──────────────────────────────────

    /// <summary>
    /// Căn chỉnh kích thước tất cả hình ảnh (Picture / LinkedPicture) trong sheet
    /// hiện tại về <see cref="STANDARD_IMAGE_WIDTH"/> × <see cref="STANDARD_IMAGE_HEIGHT"/>.
    /// Tỉ lệ khung hình sẽ bị bỏ qua để đảm bảo kích thước đồng nhất.
    /// </summary>
    /// <param name="ws">Sheet đang hoạt động.</param>
    public static void ResizeImages(Excel.Worksheet ws)
    {
      if (ws == null) throw new ArgumentNullException("ws");

      int count = 0;

      foreach (Excel.Shape shape in ws.Shapes)
      {
        // Chỉ xử lý đối tượng là hình ảnh hoặc ảnh liên kết
        if (shape.Type != MsoShapeType.msoPicture &&
            shape.Type != MsoShapeType.msoLinkedPicture)
          continue;

        // Tắt khóa tỉ lệ để có thể đặt width/height riêng biệt
        shape.LockAspectRatio = MsoTriState.msoFalse;
        shape.Width = STANDARD_IMAGE_WIDTH;
        shape.Height = STANDARD_IMAGE_HEIGHT;
        count++;
      }

      if (count == 0)
        MessageBox.Show(
            "Không tìm thấy hình ảnh nào trong sheet hiện tại.",
            "Chuẩn Hóa Hình Ảnh",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information);
      else
        MessageBox.Show(
            $"Đã chuẩn hóa {count} hình ảnh về kích thước "
            + $"{STANDARD_IMAGE_WIDTH} × {STANDARD_IMAGE_HEIGHT} pt.",
            "Chuẩn Hóa Hình Ảnh",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information);
    }

    // ─── Unhide Tất Cả Sheet ─────────────────────────────────────────────

    /// <summary>
    /// Bỏ ẩn tất cả sheet đang bị ẩn (xlSheetHidden / xlSheetVeryHidden)
    /// trong workbook hiện tại.
    /// </summary>
    /// <param name="wb">Workbook đang hoạt động.</param>
    public static void UnhideAllSheets(Excel.Workbook wb)
    {
      if (wb == null) throw new ArgumentNullException("wb");

      int count = 0;

      foreach (Excel.Worksheet ws in wb.Worksheets)
      {
        // Unhide cả xlSheetHidden lẫn xlSheetVeryHidden
        if (ws.Visible != Excel.XlSheetVisibility.xlSheetVisible)
        {
          ws.Visible = Excel.XlSheetVisibility.xlSheetVisible;
          count++;
        }
      }

      if (count == 0)
        MessageBox.Show(
            "Không có sheet nào đang bị ẩn trong workbook hiện tại.",
            "Unhide Tất Cả Sheet",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information);
      else
        MessageBox.Show(
            $"Đã hiển thị lại {count} sheet bị ẩn.",
            "Unhide Tất Cả Sheet",
            MessageBoxButtons.OK,
            MessageBoxIcon.Information);
    }
  }
}
