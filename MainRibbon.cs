using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace AddinsSupport
{
    [ComVisible(true)]
    public class MainRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI _ribbon;

        // Truy cập Application lười (lazy) — đảm bảo không null sau khi AddIn khởi động xong.
        // Khắc phục lỗi: CreateRibbonExtensibilityObject() được gọi trong base.Initialize()
        // trước khi Globals.ThisAddIn.Application được gán, nên không thể inject qua constructor.
        private static Excel.Application App => Globals.ThisAddIn?.Application;

        /// <summary>Index định dạng Sheet ID đang chọn trên dropdown.
        /// Tương ứng với <see cref="Features.HyperlinkManager.IdFormats"/> index.</summary>
        private int _hyperlinkModeIndex = 0;

        /// <summary>Nội dung ô text 'tbCustomFormat'. Nếu không rỗng → ưu tiên hơn dropdown.</summary>
        private string _customFormatText = string.Empty;

        #region IRibbonExtensibility

        public MainRibbon() { }

        public string GetCustomUI(string ribbonID) => GetResourceText("AddinsSupport.MainRibbon.xml");

        #endregion

        #region Callbacks

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
        }

        // ══════════════════════════════════════════════════════════════════════
        // NHÓM 1 — CHỈNH SỬA SHEET
        // ══════════════════════════════════════════════════════════════════════

        /// <summary>Tô màu nền cho toàn dòng (hoặc dải cột) của vùng đang chọn.</summary>
        public void OnColorSelection(Office.IRibbonControl control)
        {
            if (App?.ActiveWorkbook == null) return;
            Excel.Worksheet ws = App.ActiveSheet as Excel.Worksheet;
            if (ws == null) return;
            Features.SheetEditingManager.ColorSelection(ws);
        }

        /// <summary>Mở dialog cài đặt tô màu (màu nền, giới hạn cột).</summary>
        public void OnColorSettings(Office.IRibbonControl control)
        {
            using (var form = new Forms.ColorSelectionSettingsForm())
                form.ShowDialog();
        }

        /// <summary>Tạo sheet mới sao chép định dạng từ sheet cuối trong workbook.
        /// Tên sheet mới dựa vào lựa chọn trên dropdown ddHyperlinkMode hoặc text ghi đè.</summary>
        public void OnAddSheetWithFormat(Office.IRibbonControl control)
        {
            Excel.Workbook wb = App?.ActiveWorkbook;
            if (wb == null) return;
            int idx = ResolveFormatIndex();
            if (idx == -2) idx = -1; // text không khớp → treat as None
            Features.SheetEditingManager.AddSheetWithFormat(wb, idx);
        }

        /// <summary>Chuẩn hóa tất cả sheet trong workbook: đặt zoom về X%, focus về A1.</summary>
        public void OnResizeImages(Office.IRibbonControl control)
        {
            Excel.Workbook wb = App?.ActiveWorkbook;
            if (wb == null) return;
            Features.SheetEditingManager.NormalizeSheets(wb);
        }

        /// <summary>Bỏ ẩn tất cả sheet đang bị ẩn trong workbook.</summary>
        public void OnUnhideAllSheets(Office.IRibbonControl control)
        {
            Excel.Workbook wb = App?.ActiveWorkbook;
            if (wb == null) return;
            Features.SheetEditingManager.UnhideAllSheets(wb);
        }

        // ══════════════════════════════════════════════════════════════════════
        // NHÓM 2 — TỰ ĐỘNG THỰC THI
        // ══════════════════════════════════════════════════════════════════════

        // ── Dropdown: Chọn định dạng Sheet ID ──────────────────────────────

        /// <summary>Trả về số lượng mục (IdFormats + 1 mục "--None--" ở đầu).</summary>
        public int GetHyperlinkModeCount(Office.IRibbonControl control)
            => Features.HyperlinkManager.IdFormats.Count + 1;

        /// <summary>Trả về nhãn hiển thị của mục thứ <paramref name="index"/> trong dropdown.</summary>
        public string GetHyperlinkModeLabel(Office.IRibbonControl control, int index)
            => index == 0 ? "-- None --" : Features.HyperlinkManager.IdFormats[index - 1].Name;

        /// <summary>Trả về ID duy nhất của mục thứ <paramref name="index"/>.</summary>
        public string GetHyperlinkModeID(Office.IRibbonControl control, int index)
            => index == 0 ? "hlMode_none" : "hlMode_" + index;

        /// <summary>Trả về index đang được chọn để Ribbon duy trì trạng thái UI.</summary>
        public int GetHyperlinkModeIndex(Office.IRibbonControl control)
            => _hyperlinkModeIndex;

        /// <summary>Lưu lại chế độ người dùng vừa chọn.</summary>
        public void OnHyperlinkModeChanged(Office.IRibbonControl control, string selectedId, int selectedIndex)
            => _hyperlinkModeIndex = selectedIndex;

        /// <summary>Lưu nội dung ô text ghi đè format.</summary>
        public void OnCustomFormatChanged(Office.IRibbonControl control, string text)
            => _customFormatText = text ?? string.Empty;

        /// <summary>
        /// Giải quyết index định dạng hiệu lực (0-based trong IdFormats):<br/>
        /// • Nếu <see cref="_customFormatText"/> có nội dung → tìm theo tên (ưu tiên).<br/>
        /// • Ngược lại → dùng <see cref="_hyperlinkModeIndex"/> - 1.<br/>
        /// Trả về -1 nếu None, -2 nếu text nhập không khớp format nào.
        /// </summary>
        private int ResolveFormatIndex()
        {
            if (!string.IsNullOrWhiteSpace(_customFormatText))
            {
                var formats = Features.HyperlinkManager.IdFormats;
                string key = _customFormatText.Trim();
                for (int i = 0; i < formats.Count; i++)
                    if (string.Equals(formats[i].Name, key, StringComparison.OrdinalIgnoreCase))
                        return i;
                return -2; // text không khớp format nào
            }
            return _hyperlinkModeIndex - 1;
        }

        // ── Nút thực thi ────────────────────────────────────────────────────

        /// <summary>Quét toàn bộ sheet hiện tại và tự động thêm hyperlink theo chế độ đã chọn.</summary>
        public void OnAutoHyperlink(Office.IRibbonControl control)
        {
            if (App?.ActiveWorkbook == null) return;
            Excel.Worksheet ws = App.ActiveSheet as Excel.Worksheet;
            if (ws == null) return;
            int idx = ResolveFormatIndex();
            if (idx == -2)
            {
                System.Windows.Forms.MessageBox.Show(
                    $"Định dạng '{_customFormatText.Trim()}' không khớp với bất kỳ định dạng nào.\n"
                    + "Vui lòng kiểm tra lại hoặc xóa text để dùng dropdown.",
                    "AutoHyperlink",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);
                return;
            }
            if (idx < 0)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Vui lòng chọn định dạng Sheet ID trên dropdown hoặc nhập vào ô Format.",
                    "AutoHyperlink",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);
                return;
            }
            Features.HyperlinkManager.AutoAddHyperlinks(ws, idx);
        }

        /// <summary>
        /// Xóa hyperlink tại từng ô trong vùng chọn và đồng thời xóa
        /// back-link 戻る tại A1 của sheet đích mà hyperlink đó trỏ đến.
        /// </summary>
        public void OnRemoveHyperlinks(Office.IRibbonControl control)
        {
            if (App?.ActiveWorkbook == null) return;
            Excel.Worksheet ws = App.ActiveSheet as Excel.Worksheet;
            if (ws == null) return;
            Features.HyperlinkManager.RemoveHyperlinks(ws);
        }

        /// <summary>
        /// Shift tên sheet SEQ từ vị trí active trở đi lên +1.
        /// Chế độ được lấy từ dropdown ddHyperlinkMode:
        ///   index 1 (SEQ.xxx)  → Mode 0 (toàn workbook)
        ///   index 2 (SEQg.xxx) → Mode 1 (chỉ trong nhóm)
        /// </summary>
        public void OnRenameByFormat(Office.IRibbonControl control)
        {
            Excel.Workbook wb = App?.ActiveWorkbook;
            if (wb == null) return;

            int seqMode = ResolveFormatIndex();
            if (seqMode == -2)
            {
                System.Windows.Forms.MessageBox.Show(
                    $"Định dạng '{_customFormatText.Trim()}' không khớp với bất kỳ định dạng nào.\n"
                    + "Vui lòng kiểm tra lại hoặc xóa text để dùng dropdown.",
                    "AutoRenameSeqSheets",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);
                return;
            }
            if (seqMode < 0 || seqMode > 1)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Vui lòng chọn chế độ SEQ.xxx hoặc SEQg.xxx trên dropdown 'Định dạng Sheet ID' trước khi thực hiện.",
                    "AutoRenameSeqSheets",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);
                return;
            }

            Features.SheetSeqRenamer.AutoRenameSeqSheets(wb, seqMode);
        }

        /// <summary>
        /// Kiểm tra và chỉnh lại thứ tự đánh số SEQ theo vị trí tab.
        /// Dùng cùng chế độ với dropdown ddHyperlinkMode.
        /// </summary>
        public void OnCheckFixSeqOrder(Office.IRibbonControl control)
        {
            Excel.Workbook wb = App?.ActiveWorkbook;
            if (wb == null) return;

            int seqMode = ResolveFormatIndex();
            if (seqMode == -2)
            {
                System.Windows.Forms.MessageBox.Show(
                    $"Định dạng '{_customFormatText.Trim()}' không khớp với bất kỳ định dạng nào.\n"
                    + "Vui lòng kiểm tra lại hoặc xóa text để dùng dropdown.",
                    "CheckAndFixSeqOrder",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);
                return;
            }
            if (seqMode < 0 || seqMode > 1)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Vui lòng chọn chế độ SEQ.xxx hoặc SEQg.xxx trên dropdown 'Định dạng Sheet ID' trước khi thực hiện.",
                    "CheckAndFixSeqOrder",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);
                return;
            }

            Features.SheetSeqRenamer.CheckAndFixSeqOrder(wb, seqMode);
        }

        /// <summary>Sắp xếp tất cả sheet theo alphabet và chuẩn hóa tên (xóa khoảng trắng thừa).</summary>
        public void OnSortNormalizeSheets(Office.IRibbonControl control)
        {
            Excel.Workbook wb = App?.ActiveWorkbook;
            if (wb == null) return;
            Features.SheetNameManager.SortAndNormalizeSheets(wb);
        }

        // ══════════════════════════════════════════════════════════════════════
        // NHÓM 3 — TIỆN ÍCH MỞ RỘNG (Đang Phát Triển)
        // ══════════════════════════════════════════════════════════════════════

        /// <summary>Hiển thị thông báo "Đang phát triển" cho các tính năng chưa hoàn thiện.</summary>
        public void OnExtensionComingSoon(Office.IRibbonControl control)
        {
            Features.ExtensionsManager.ShowComingSoon(control?.Id);
        }

        /// <summary>Trả về chuỗi phiên bản add-in, ví dụ "Phiên bản: v1.0.0".</summary>
        public string GetVersionLabel(Office.IRibbonControl control)
            => "Phiên bản: v" + Assembly.GetExecutingAssembly().GetName().Version.ToString(3);

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            foreach (string name in asm.GetManifestResourceNames())
            {
                if (string.Compare(resourceName, name, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader reader = new StreamReader(asm.GetManifestResourceStream(name)))
                        return reader.ReadToEnd();
                }
            }
            return null;
        }

        #endregion
    }
}
