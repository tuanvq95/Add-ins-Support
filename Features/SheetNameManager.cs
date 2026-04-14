using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace AddinsSupport.Features
{
    /// <summary>
    /// Quản lý tên sheet: đổi tên hàng loạt theo prefix/suffix hoặc theo nội dung cell A1.
    /// </summary>
    public static class SheetNameManager
    {
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
        /// Đổi tên sheet hiện tại theo format [yyyyMMdd]_[Nội dung ô A1].
        /// Nếu ô A1 trống thì chỉ dùng ngày. Áp dụng cho sheet đang active.
        /// </summary>
        /// <param name="wb">Workbook đang hoạt động.</param>
        public static void RenameSheetsByFormat(Excel.Workbook wb)
        {
            if (wb == null) throw new ArgumentNullException("wb");

            Excel.Worksheet activeSheet = wb.Application.ActiveSheet as Excel.Worksheet;
            if (activeSheet == null)
            {
                MessageBox.Show("Không tìm thấy sheet đang active.",
                    "Đặt Tên Theo Format", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Lấy nội dung ô A1 làm phần tên
            string cellValue = activeSheet.Range["A1"].Value2 as string;
            string datePart = DateTime.Now.ToString("yyyyMMdd");

            string rawName = string.IsNullOrWhiteSpace(cellValue)
                ? datePart
                : datePart + "_" + cellValue.Trim();

            string newName = SanitizeSheetName(rawName);
            newName = EnsureUnique(wb, activeSheet, newName);

            string oldName = activeSheet.Name;
            activeSheet.Name = newName;

            MessageBox.Show(
                $"Đã đổi tên sheet:\n• Trước: {oldName}\n• Sau:   {newName}",
                "Đặt Tên Theo Format",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        /// <summary>
        /// Sắp xếp tất cả sheet trong workbook theo thứ tự bảng chữ cái (A→Z)
        /// và loại bỏ khoảng trắng thừa ở đầu/cuối tên mỗi sheet.
        /// </summary>
        /// <param name="wb">Workbook đang hoạt động.</param>
        public static void SortAndNormalizeSheets(Excel.Workbook wb)
        {
            if (wb == null) throw new ArgumentNullException("wb");

            int sheetCount = wb.Worksheets.Count;

            // Bước 1: Chuẩn hóa tên — xóa khoảng trắng đầu/cuối
            int normalized = 0;
            foreach (Excel.Worksheet ws in wb.Worksheets)
            {
                string trimmed = ws.Name.Trim();
                if (trimmed != ws.Name)
                {
                    ws.Name = EnsureUnique(wb, ws, trimmed);
                    normalized++;
                }
            }

            // Bước 2: Sắp xếp theo alphabet bằng bubble sort (số sheet thường nhỏ)
            bool swapped = true;
            while (swapped)
            {
                swapped = false;
                for (int i = 1; i < sheetCount; i++)
                {
                    Excel.Worksheet a = wb.Worksheets[i] as Excel.Worksheet;
                    Excel.Worksheet b = wb.Worksheets[i + 1] as Excel.Worksheet;
                    if (a == null || b == null) continue;

                    if (string.Compare(a.Name, b.Name, StringComparison.OrdinalIgnoreCase) > 0)
                    {
                        // Di chuyển sheet b lên trước sheet a
                        b.Move(Before: a);
                        swapped = true;
                    }
                }
            }

            MessageBox.Show(
                $"Hoàn thành!\n• Đã sắp xếp {sheetCount} sheet theo thứ tự A→Z\n"
                + $"• Đã chuẩn hóa tên: {normalized} sheet",
                "Sắp Xếp &amp; Chuẩn Hóa Sheet",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        /// <summary>
        /// Nếu tên đã tồn tại, thêm số hậu tố để tạo tên duy nhất.
        /// </summary>
        public static string EnsureUnique(Excel.Workbook wb, Excel.Worksheet currentSheet, string name)
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
