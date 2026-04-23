using System;
using System.Collections.Generic;
using System.Globalization;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace AddinsSupport.Features
{
    // ═════════════════════════════════════════════════════════════════════════
    // Kết quả sau khi xử lý một vùng ô
    // ═════════════════════════════════════════════════════════════════════════

    public sealed class HyperlinkProcessResult
    {
        public int Added { get; set; }
        public int Skipped { get; set; }
    }

    // ═════════════════════════════════════════════════════════════════════════
    // Lớp cơ sở trừu tượng — mở rộng bằng cách kế thừa lớp này
    //
    //  ★ Cách thêm định dạng mới:
    //     1. Tạo lớp kế thừa SheetIdFormat (trong file này hoặc file riêng)
    //     2. Override ProcessRange() với logic nhận diện + tạo hyperlink
    //     3. Append instance vào HyperlinkManager.IdFormats
    //     → Dropdown Ribbon tự cập nhật, không cần sửa XML
    // ═════════════════════════════════════════════════════════════════════════

    public abstract class SheetIdFormat
    {
        /// <summary>Tên ngắn hiển thị trên dropdown Ribbon.</summary>
        public string Name { get; }

        /// <summary>Mô tả đầy đủ dùng làm supertip.</summary>
        public string Description { get; }

        /// <summary>
        /// true  = xử lý toàn bộ UsedRange (vd: URL Only).<br/>
        /// false = xử lý vùng ô đang được người dùng chọn (vd: Sheet ID modes).
        /// </summary>
        public virtual bool UseUsedRange { get { return false; } }

        protected SheetIdFormat(string name, string description)
        {
            Name = name;
            Description = description;
        }

        /// <summary>
        /// Điểm vào chính: xử lý <paramref name="range"/>, tạo hyperlink,
        /// định dạng ô và thêm back-link nếu cần.
        /// </summary>
        public abstract HyperlinkProcessResult ProcessRange(Excel.Range range, Excel.Worksheet ws);

        // ── Tiện ích dùng chung cho các lớp con ──────────────────────────────

        /// <summary>
        /// Chuyển chuỗi từ ô Excel thành số nguyên.
        /// Xử lý cả trường hợp Excel trả về "1.0" cho giá trị nguyên.
        /// </summary>
        protected static bool TryParseLong(string s, out long result)
        {
            result = 0;
            if (string.IsNullOrEmpty(s)) return false;
            if (long.TryParse(s, out result)) return true;

            // Excel đôi khi trả về "1.0" cho số nguyên khi Value2 là double
            if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out double d)
                && d == Math.Floor(d) && d > 0)
            { result = (long)d; return true; }

            return false;
        }

        /// <summary>
        /// Áp dụng viền 4 cạnh và căn giữa cho ô (theo hành vi VBA gốc).
        /// </summary>
        protected static void ApplyCellStyle(Excel.Range cell)
        {
            var b = cell.Borders;
            b[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            b[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            b[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            b[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            cell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
        }

        /// <summary>
        /// Khôi phục format ô về trạng thái bình thường sau khi xóa hyperlink:
        /// bỏ màu xanh, bỏ gạch chân, giữ nguyên viền và căn giữa (đối xứng với ApplyCellStyle).
        /// </summary>
        internal static void RestoreCellStyle(Excel.Range cell)
        {
            // Xóa màu font và gạch chân do hyperlink để lại
            cell.Font.ColorIndex = Excel.Constants.xlAutomatic;
            cell.Font.Underline = Excel.XlUnderlineStyle.xlUnderlineStyleNone;

            // Giữ lại viền 4 cạnh và căn giữa như ApplyCellStyle đã áp
            var b = cell.Borders;
            b[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            b[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            b[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            b[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            cell.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            cell.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
        }

        /// <summary>
        /// Thêm hyperlink "戻る" vào ô A1 của sheet đích, trỏ ngược về
        /// ô A1 của sheet nguồn (giữ nguyên hành vi VBA gốc).
        /// </summary>
        protected static void AddBackLink(
            Excel.Workbook wb,
            string targetSheetName,
            string sourceSheetName,
            int anchorRow,
            int anchorCol)
        {
            // "戻る" = tiếng Nhật "quay lại" — giữ nguyên theo VBA gốc
            const string BACK_TEXT = "戻る";
            // màu xanh dương RGB(0,102,204) ở định dạng OLE BGR
            const int BLUE_OLE = 13395456;

            try
            {
                Excel.Worksheet target = null;
                foreach (Excel.Worksheet sh in wb.Worksheets)
                    if (string.Equals(sh.Name, targetSheetName, StringComparison.OrdinalIgnoreCase))
                    { target = sh; break; }

                if (target == null) return;

                Excel.Range a1 = target.Cells[1, 1] as Excel.Range;
                if (a1 == null) return;

                a1.Hyperlinks.Delete();
                a1.ClearContents();

                target.Hyperlinks.Add(
                    Anchor: a1,
                    Address: string.Empty,
                    SubAddress: $"'{sourceSheetName}'!{ColAddress(anchorCol)}{anchorRow}",
                    TextToDisplay: BACK_TEXT);

                a1.Font.Color = BLUE_OLE;
                a1.Font.Underline = Excel.XlUnderlineStyle.xlUnderlineStyleSingle;
            }
            catch { /* sheet bị bảo vệ hoặc lỗi khác, bỏ qua */ }
        }

        /// <summary>
        /// Chuyển chỉ số cột (1-based) thành chữ cái Excel, vd: 1→"A", 27→"AA".
        /// </summary>
        private static string ColAddress(int col)
        {
            string s = string.Empty;
            while (col > 0)
            {
                col--;
                s = (char)('A' + col % 26) + s;
                col /= 26;
            }
            return s;
        }
    }

    // ═════════════════════════════════════════════════════════════════════════
    // Mode 0 — SEQ.xxx hoặc SEQ.xxx~yyy
    //
    //   Port từ VBA: AutoAddHyperlinksToSeqSheets (HYPERLINK_MODE = 0)
    //                ParseSeqMode0 / FindSheetForSeq0
    //
    //   Cách dùng:
    //     • Chọn vùng ô chứa số thứ tự (vd: 1, 5, 10)
    //     • Mỗi ô tự đọc giá trị số của chính nó
    //     • Tìm sheet tên "SEQ.{số}" hoặc range "SEQ.{x}~{y}" với số ∈ [x, y]
    //     • Tạo internal hyperlink + viền + căn giữa + back-link tại A1 sheet đích
    // ═════════════════════════════════════════════════════════════════════════

    public sealed class SeqDotFormat : SheetIdFormat
    {
        public SeqDotFormat()
            : base("SEQ.xxx",
                   "Chọn vùng ô chứa số thứ tự → tìm sheet SEQ.{số} hoặc SEQ.{x}~{y}.\n"
                   + "Ví dụ: ô chứa '5' → tìm sheet 'SEQ.5' hoặc 'SEQ.1~10'.\n"
                   + "Áp dụng viền, căn giữa và back-link 戻る tại A1 sheet đích.")
        { }

        public override HyperlinkProcessResult ProcessRange(Excel.Range range, Excel.Worksheet ws)
        {
            var result = new HyperlinkProcessResult();
            Excel.Workbook wb = ws.Parent as Excel.Workbook;
            if (wb == null) return result;

            foreach (Excel.Range cell in range.Cells)
            {
                object raw = cell.Value2;
                if (raw == null) { result.Skipped++; continue; }

                if (!TryParseLong(raw.ToString().Trim(), out long seqNum))
                { result.Skipped++; continue; }

                string sheetName = FindSheet(wb, seqNum);
                if (string.IsNullOrEmpty(sheetName)) { result.Skipped++; continue; }

                try
                {
                    // Xóa hyperlink cũ nếu có
                    if (cell.Hyperlinks.Count > 0) cell.Hyperlinks.Delete();

                    ws.Hyperlinks.Add(
                        Anchor: cell,
                        Address: string.Empty,
                        SubAddress: $"'{sheetName}'!A1",
                        TextToDisplay: seqNum.ToString());

                    ApplyCellStyle(cell);
                    AddBackLink(wb, sheetName, ws.Name, cell.Row, cell.Column);
                    result.Added++;
                }
                catch { result.Skipped++; }
            }

            return result;
        }

        private string FindSheet(Excel.Workbook wb, long num)
        {
            foreach (Excel.Worksheet sh in wb.Worksheets)
            {
                long s, e;
                if (ParseName(sh.Name, out s, out e))
                    if (e >= 0 ? (num >= s && num <= e) : num == s)
                        return sh.Name;
            }
            return null;
        }

        /// <summary>
        /// Parse tên sheet: "SEQ.xxx" hoặc "SEQ.xxx~yyy".<br/>
        /// outEnd = -1 nghĩa là không có range (số đơn).
        /// </summary>
        private static bool ParseName(string name, out long outStart, out long outEnd)
        {
            outStart = 0;
            outEnd = -1;

            // Phải bắt đầu bằng "SEQ." (4 ký tự)
            if (!name.StartsWith("SEQ.", StringComparison.OrdinalIgnoreCase)) return false;
            string rest = name.Substring(4);

            // Không cho phép thêm dấu chấm (tránh nhầm với SEQg.xxx)
            if (rest.IndexOf('.') >= 0) return false;

            int tilde = rest.IndexOf('~');
            if (tilde > 0)
            {
                long s, e;
                if (TryParseLong(rest.Substring(0, tilde), out s)
                    && TryParseLong(rest.Substring(tilde + 1), out e))
                { outStart = s; outEnd = e; return true; }
            }
            else
            {
                long s;
                if (TryParseLong(rest, out s))
                { outStart = s; return true; }
            }

            return false;
        }
    }

    // ═════════════════════════════════════════════════════════════════════════
    // Mode 1 — SEQg.xxx hoặc SEQg.xxx~yyy  (g = số nhóm cha)
    //
    //   Port từ VBA: AutoAddHyperlinksToSeqSheets (HYPERLINK_MODE = 1)
    //                ParseSeqMode1 / FindSheetForSeq1
    //
    //   Cách dùng:
    //     • Chọn vùng ô sẽ nhận hyperlink (cột STT con)
    //     • Khi giá trị ô = 1: bắt đầu nhóm mới
    //       → đọc số nhóm từ cột RefColumn (mặc định A=1) cùng hàng
    //       → nếu không đọc được: tự tăng số nhóm +1
    //     • Tìm sheet "SEQ{nhóm}.{STT}" hoặc range "SEQ{nhóm}.{x}~{y}"
    //     • Đánh lại STT con từ 1 trong mỗi nhóm (ghi đè giá trị ô)
    //     • Tạo internal hyperlink + viền + căn giữa + back-link tại A1 sheet đích
    // ═════════════════════════════════════════════════════════════════════════

    public sealed class SeqGroupFormat : SheetIdFormat
    {
        /// <summary>
        /// Cột tham chiếu để đọc số nhóm (1-based: 1=A, 2=B, …).<br/>
        /// Tương đương VBA: Const REF_COLUMN As Integer = 1
        /// </summary>
        public int RefColumn { get; set; } = 1;

        public SeqGroupFormat()
            : base("SEQg.xxx",
                   "Chọn vùng ô STT con → tìm sheet SEQ{nhóm}.{STT}.\n"
                   + "Đọc số nhóm từ cột A cùng hàng. Khi STT con = 1 → bắt đầu nhóm mới.\n"
                   + "Ô được đánh lại số từ 1 trong mỗi nhóm.\n"
                   + "Ví dụ: cột A=2, ô='3' → tìm 'SEQ2.3' hoặc 'SEQ2.1~5'.")
        { }

        public override HyperlinkProcessResult ProcessRange(Excel.Range range, Excel.Worksheet ws)
        {
            var result = new HyperlinkProcessResult();
            Excel.Workbook wb = ws.Parent as Excel.Workbook;
            if (wb == null) return result;

            long lastRefSeqNum = -1; // -1 = chưa xác định được nhóm nào
            int autoIndex = 0;

            foreach (Excel.Range cell in range.Cells)
            {
                object raw = cell.Value2;
                if (raw == null) { result.Skipped++; continue; }

                if (!TryParseLong(raw.ToString().Trim(), out long cellSeqNum))
                { result.Skipped++; continue; }

                // Khi bắt đầu nhóm mới (STT con = 1): cập nhật số nhóm
                if (cellSeqNum == 1)
                {
                    autoIndex = 0;

                    // Đọc số nhóm từ cột RefColumn cùng hàng
                    Excel.Range refCell = ws.Cells[cell.Row, RefColumn] as Excel.Range;
                    string refRaw = refCell?.Value2?.ToString()?.Trim() ?? string.Empty;

                    long parsed;
                    if (TryParseLong(refRaw, out parsed) && parsed > 0)
                    {
                        // Đọc được số nhóm hợp lệ từ ô tham chiếu
                        lastRefSeqNum = parsed;
                    }
                    else if (lastRefSeqNum >= 0)
                    {
                        // Không đọc được → tự tăng số nhóm lên +1 (hành vi VBA gốc)
                        lastRefSeqNum++;
                    }
                    else
                    {
                        // Chưa có nhóm nào trước đó, bỏ qua ô này
                        result.Skipped++;
                        continue;
                    }
                }

                // Chưa xác định được nhóm → bỏ qua
                if (lastRefSeqNum < 0) { result.Skipped++; continue; }

                string sheetName = FindSheet(wb, lastRefSeqNum, cellSeqNum);
                if (string.IsNullOrEmpty(sheetName)) { result.Skipped++; continue; }

                try
                {
                    if (cell.Hyperlinks.Count > 0) cell.Hyperlinks.Delete();

                    // Đánh lại STT con bắt đầu từ 1 trong mỗi nhóm (ghi đè giá trị ô)
                    autoIndex++;
                    cell.Value = autoIndex;

                    ws.Hyperlinks.Add(
                        Anchor: cell,
                        Address: string.Empty,
                        SubAddress: $"'{sheetName}'!A1",
                        TextToDisplay: autoIndex.ToString());

                    ApplyCellStyle(cell);
                    AddBackLink(wb, sheetName, ws.Name, cell.Row, cell.Column);
                    result.Added++;
                }
                catch { result.Skipped++; }
            }

            return result;
        }

        private string FindSheet(Excel.Workbook wb, long grp, long num)
        {
            foreach (Excel.Worksheet sh in wb.Worksheets)
            {
                long g, s, e;
                if (ParseName(sh.Name, out g, out s, out e) && g == grp)
                    if (e >= 0 ? (num >= s && num <= e) : num == s)
                        return sh.Name;
            }
            return null;
        }

        /// <summary>
        /// Parse tên sheet: "SEQg.xxx" hoặc "SEQg.xxx~yyy".<br/>
        /// Ví dụ: SEQ2.22 → group=2, start=22, end=-1<br/>
        ///        SEQ3.5~8 → group=3, start=5, end=8
        /// </summary>
        private static bool ParseName(
            string name,
            out long outGroup,
            out long outStart,
            out long outEnd)
        {
            outGroup = 0;
            outStart = 0;
            outEnd = -1;

            if (!name.StartsWith("SEQ", StringComparison.OrdinalIgnoreCase)) return false;

            string rest = name.Substring(3); // vd: "2.22" hoặc "2.22~25"

            int dot = rest.IndexOf('.');
            // dot <= 0: không có dấu chấm HOẶC bắt đầu bằng "SEQ." (Mode 0)
            if (dot <= 0) return false;

            string groupStr = rest.Substring(0, dot);
            long grp;
            if (!TryParseLong(groupStr, out grp)) return false;

            string numPart = rest.Substring(dot + 1);
            // Không cho phép thêm dấu chấm trong phần số
            if (numPart.IndexOf('.') >= 0) return false;

            outGroup = grp;

            int tilde = numPart.IndexOf('~');
            if (tilde > 0)
            {
                long s, e;
                if (TryParseLong(numPart.Substring(0, tilde), out s)
                    && TryParseLong(numPart.Substring(tilde + 1), out e))
                { outStart = s; outEnd = e; return true; }
            }
            else
            {
                long s;
                if (TryParseLong(numPart, out s))
                { outStart = s; return true; }
            }

            return false;
        }
    }

    // ═════════════════════════════════════════════════════════════════════════
    // HyperlinkManager — điểm vào chính, điều phối các định dạng
    // ═════════════════════════════════════════════════════════════════════════

    public static class HyperlinkManager
    {
        // ── Danh sách định dạng — THÊM ĐỊNH DẠNG MỚI TẠI ĐÂY ────────────────
        //
        //  ★ Quy trình thêm định dạng:
        //     1. Tạo lớp kế thừa SheetIdFormat (file này hoặc file riêng)
        //     2. Implement ProcessRange() với logic nhận diện + tạo hyperlink
        //     3. Append instance vào danh sách dưới đây
        //     → Dropdown Ribbon tự cập nhật, không cần sửa XML

        public static readonly List<SheetIdFormat> IdFormats = new List<SheetIdFormat>
        {
            new SeqDotFormat(),    // [0] Mode 0: SEQ.xxx / SEQ.xxx~yyy
            new SeqGroupFormat(),  // [1] Mode 1: SEQg.xxx / SEQg.xxx~yyy
            // new MyCustomFormat(),  ← thêm định dạng mới tại đây
        };

        /// <summary>
        /// Thực thi Auto Hyperlink theo định dạng <paramref name="modeIndex"/>
        /// trong <see cref="IdFormats"/>.
        /// </summary>
        /// <param name="ws">Sheet đang hoạt động.</param>
        /// <param name="modeIndex">Index trong IdFormats (mặc định 0 = URL Only).</param>
        public static void AutoAddHyperlinks(Excel.Worksheet ws, int modeIndex = 0)
        {
            if (ws == null) throw new ArgumentNullException("ws");

            if (modeIndex < 0 || modeIndex >= IdFormats.Count)
            {
                MessageBox.Show("Định dạng không hợp lệ. Vui lòng chọn lại.",
                    "Auto Hyperlink", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            SheetIdFormat fmt = IdFormats[modeIndex];

            // Xác định vùng xử lý: UsedRange hoặc vùng đang chọn
            Excel.Range range;
            if (fmt.UseUsedRange)
            {
                range = ws.UsedRange;
            }
            else
            {
                range = ws.Application.Selection as Excel.Range;
                if (range == null)
                {
                    MessageBox.Show(
                        $"Vui lòng chọn vùng ô trước khi thực hiện.\n"
                        + $"Chế độ '{fmt.Name}' yêu cầu đánh dấu vùng cần xử lý.",
                        "Auto Hyperlink",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }
            }

            HyperlinkProcessResult result = fmt.ProcessRange(range, ws);

            MessageBox.Show(
                $"Hoàn thành! Chế độ: {fmt.Name}\n"
                + $"• Đã thêm: {result.Added} hyperlink\n"
                + $"• Bỏ qua:  {result.Skipped} ô",
                "Auto Hyperlink — " + ws.Name,
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        // ── Xóa Hyperlink + Back-link ─────────────────────────────────────────

        /// <summary>
        /// Với mỗi ô trong vùng đang chọn có hyperlink nội bộ (SubAddress):
        ///   1. Xóa hyperlink tại ô đó, khôi phục nội dung gốc.
        ///   2. Tìm sheet đích từ SubAddress → xóa back-link 戻る tại A1.
        /// Hỗ trợ cả selection liên tục và nhiều vùng rời rạc (Ctrl+Click).
        /// </summary>
        public static void RemoveHyperlinks(Excel.Worksheet ws)
        {
            if (ws == null) throw new ArgumentNullException("ws");

            Excel.Workbook wb = ws.Parent as Excel.Workbook;
            if (wb == null) return;

            Excel.Range selection = ws.Application.Selection as Excel.Range;
            if (selection == null)
            {
                MessageBox.Show(
                    "Vui lòng chọn vùng ô cần xóa hyperlink trước khi thực hiện.",
                    "Xóa Hyperlink",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            int removed = 0;
            int backRemoved = 0;

            foreach (Excel.Range area in selection.Areas)
            {
                foreach (Excel.Range cell in area.Cells)
                {
                    if (cell.Hyperlinks.Count == 0) continue;

                    foreach (Excel.Hyperlink hl in cell.Hyperlinks)
                    {
                        // Chỉ xử lý hyperlink nội bộ (Address rỗng, có SubAddress)
                        string sub = hl.SubAddress;
                        if (!string.IsNullOrEmpty(sub))
                        {
                            // Parse tên sheet đích từ SubAddress: 'SheetName'!CellAddr
                            string targetSheetName = ParseSheetFromSubAddress(sub);
                            if (!string.IsNullOrEmpty(targetSheetName))
                                backRemoved += RemoveBackLinkAt(wb, targetSheetName);
                        }
                    }

                    // Lưu lại giá trị trước khi xóa để khôi phục nội dung
                    object savedValue = cell.Value2;
                    cell.Hyperlinks.Delete();
                    // Hyperlinks.Delete() có thể xóa nội dung ô — khôi phục lại
                    if (savedValue != null && (cell.Value2 == null
                        || cell.Value2.ToString() != savedValue.ToString()))
                        cell.Value = savedValue;

                    // Khôi phục format: bỏ màu xanh/gạch chân, giữ viền + căn giữa
                    SheetIdFormat.RestoreCellStyle(cell);

                    removed++;
                }
            }

            string msg = removed > 0
                ? $"Đã xóa {removed} hyperlink trong vùng chọn."
                    + (backRemoved > 0 ? $"\nĐã xóa {backRemoved} back-link 戻る ở sheet đích." : string.Empty)
                : "Không tìm thấy hyperlink nào trong vùng chọn.";

            MessageBox.Show(msg, "Xóa Hyperlink",
                MessageBoxButtons.OK,
                removed > 0 ? MessageBoxIcon.Information : MessageBoxIcon.Warning);
        }

        /// <summary>
        /// Parse tên sheet từ SubAddress dạng <c>'SheetName'!CellAddr</c>
        /// hoặc <c>SheetName!CellAddr</c>.
        /// </summary>
        private static string ParseSheetFromSubAddress(string sub)
        {
            if (string.IsNullOrEmpty(sub)) return null;
            int bang = sub.LastIndexOf('!');
            if (bang <= 0) return null;
            return sub.Substring(0, bang).Trim('\'', ' ');
        }

        /// <summary>
        /// Xóa back-link 戻る tại ô A1 của sheet <paramref name="sheetName"/>
        /// trong <paramref name="wb"/>. Trả về 1 nếu đã xóa, 0 nếu không tìm thấy.
        /// </summary>
        private static int RemoveBackLinkAt(Excel.Workbook wb, string sheetName)
        {
            const string BACK_TEXT = "戻る";
            try
            {
                foreach (Excel.Worksheet sh in wb.Worksheets)
                {
                    if (!string.Equals(sh.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                        continue;

                    Excel.Range a1 = sh.Cells[1, 1] as Excel.Range;
                    if (a1 == null) return 0;

                    bool found = false;
                    foreach (Excel.Hyperlink hl in a1.Hyperlinks)
                    {
                        if (string.Equals(hl.TextToDisplay, BACK_TEXT, StringComparison.Ordinal))
                        { found = true; break; }
                    }

                    if (!found) return 0;

                    a1.Hyperlinks.Delete();
                    a1.ClearContents();
                    return 1;
                }
            }
            catch { /* sheet bảo vệ hoặc lỗi khác — bỏ qua */ }
            return 0;
        }
    }
}

