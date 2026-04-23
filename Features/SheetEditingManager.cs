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

      // Dùng HashSet để tránh tô lại cùng một dòng khi nhiều cell rời rạc
      // nằm trên cùng dòng (vd: chọn A1 và C1 riêng lẻ → chỉ tô dòng 1 một lần).
      var coloredRows = new System.Collections.Generic.HashSet<int>();

      // selection.Areas hỗ trợ cả 3 case:
      //   • 1 cell      → 1 Area, 1 dòng
      //   • Vùng liên tục → 1 Area, nhiều dòng
      //   • Cell rời rạc  → nhiều Area, mỗi Area 1 hoặc nhiều dòng
      foreach (Excel.Range area in selection.Areas)
      {
        int firstRow = area.Row;
        int lastRow = firstRow + area.Rows.Count - 1;

        for (int rowNum = firstRow; rowNum <= lastRow; rowNum++)
        {
          if (!coloredRows.Add(rowNum)) continue; // đã tô dòng này rồi

          Excel.Range target;
          if (limitCols)
            target = ws.Range[ws.Cells[rowNum, colFrom], ws.Cells[rowNum, colTo]] as Excel.Range;
          else
            target = ws.Rows[rowNum] as Excel.Range;

          if (target != null)
            target.Interior.Color = colorBgr;
        }
      }
    }

    // ─── Thêm Sheet Theo Format ───────────────────────────────────────────

    /// <summary>
    /// Tạo sheet mới bằng cách sao chép sheet cuối cùng trong workbook
    /// (giữ nguyên định dạng, column width, row height, v.v.) rồi xóa nội dung.
    /// <para>Tên sheet mới được tính từ tên sheet cuối theo định dạng SEQ tương ứng:</para>
    /// <list type="bullet">
    ///   <item>seqMode = -1 (None)  → "Sheet_yyyyMMdd"</item>
    ///   <item>seqMode =  0 (SEQ.xxx)  → "SEQ.{N+1}" kế tiếp sheet cuối</item>
    ///   <item>seqMode =  1 (SEQg.xxx) → "SEQ{g}.{N+1}" kế tiếp trong nhóm</item>
    /// </list>
    /// Nếu tên sheet cuối không khớp định dạng đã chọn, fallback về "Sheet_yyyyMMdd".
    /// </summary>
    /// <param name="wb">Workbook đang hoạt động.</param>
    /// <param name="seqMode">-1=None, 0=SEQ.xxx, 1=SEQg.xxx.</param>
    public static void AddSheetWithFormat(Excel.Workbook wb, int seqMode = -1)
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

      // Ghi nhớ tên tất cả sheet hiện có để tìm sheet mới sau khi copy
      var existingNames = new System.Collections.Generic.HashSet<string>(
          StringComparer.OrdinalIgnoreCase);
      foreach (Excel.Worksheet s in wb.Worksheets)
        existingNames.Add(s.Name);

      // Copy sheet cuối → sheet mới xuất hiện sau nó
      templateSheet.Copy(After: templateSheet);

      // Tìm sheet vừa được tạo bằng cách so sánh tên (tránh lệ thuộc vào index)
      Excel.Worksheet newSheet = null;
      foreach (Excel.Worksheet s in wb.Worksheets)
      {
        if (!existingNames.Contains(s.Name))
        {
          newSheet = s;
          break;
        }
      }
      if (newSheet == null) return;

      // Tính tên mới TRƯỚC khi xóa nội dung để tránh mọi exception làm tên sai
      string baseName = BuildSheetName(templateSheet.Name, seqMode);
      newSheet.Name = SheetNameManager.EnsureUnique(wb, newSheet, baseName);

      // Xóa nội dung nhưng giữ lại định dạng ô
      try { newSheet.UsedRange.ClearContents(); }
      catch { /* sheet rỗng hoặc protected — bỏ qua */ }

      // Kích hoạt sheet mới để người dùng thấy ngay
      newSheet.Activate();

      MessageBox.Show(
          $"Đã tạo sheet mới '{newSheet.Name}'\n"
          + $"từ định dạng của sheet '{templateSheet.Name}'.",
          "Thêm Sheet Theo Format",
          MessageBoxButtons.OK,
          MessageBoxIcon.Information);
    }

    // ─── Tính Tên Sheet Mới Theo SEQ Format ────────────────────────────────

    /// <summary>
    /// Tính tên sheet kế tiếp dựa trên tên sheet mẫu và chế độ SEQ đang chọn.
    /// </summary>
    private static string BuildSheetName(string templateName, int seqMode)
    {
      if (seqMode == 0)
      {
        long s, e;
        if (SheetSeqRenamer.ParseMode0(templateName, out s, out e))
          return $"SEQ.{(e >= 0 ? e + 1 : s + 1)}";
      }
      else if (seqMode == 1)
      {
        long g, s, e;
        if (SheetSeqRenamer.ParseMode1(templateName, out g, out s, out e))
          return $"SEQ{g}.{(e >= 0 ? e + 1 : s + 1)}";
      }

      return "Sheet_" + DateTime.Now.ToString("yyyyMMdd");
    }

    // ─── Chuẩn Hóa Sheet ─────────────────────────────────────────────────

    /// <summary>
    /// Duyệt qua tất cả sheet trong workbook, đặt zoom về <see cref="ColorSelectionSettings.SheetZoomPercent"/>%
    /// và chuyển ô focus về A1. Sau khi hoàn tất, sheet đang active trước đó được khôi phục.
    /// </summary>
    /// <param name="wb">Workbook đang hoạt động.</param>
    public static void NormalizeSheets(Excel.Workbook wb)
    {
      if (wb == null) throw new ArgumentNullException("wb");

      int zoom = ColorSelectionSettings.SheetZoomPercent;
      Excel.Worksheet originalSheet = wb.Application.ActiveSheet as Excel.Worksheet;
      int count = 0;

      foreach (Excel.Worksheet ws in wb.Worksheets)
      {
        ws.Activate();
        wb.Application.ActiveWindow.Zoom = zoom;
        ws.Range["A1"].Select();
        count++;
      }

      // Khôi phục sheet đang active trước đó
      if (originalSheet != null)
        originalSheet.Activate();

      MessageBox.Show(
          $"Đã chuẩn hóa {count} sheet: zoom {zoom}%, focus về A1.",
          "Chuẩn Hóa Sheet",
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
