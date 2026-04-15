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

        #region IRibbonExtensibility

        public MainRibbon() { }

        public string GetCustomUI(string ribbonID)
        {
            try
            {
                string xml = GetResourceText("AddinsSupport.MainRibbon.xml");
                if (xml == null)
                {
                    var names = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceNames();
                    System.Windows.Forms.MessageBox.Show(
                        "Không tìm thấy MainRibbon.xml.\n\nTài nguyên hiện có:\n"
                        + string.Join("\n", names),
                        "Ribbon Load Error",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Error);
                }
                return xml;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Lỗi khi tải Ribbon XML:\n" + ex.Message + "\n\n" + ex.StackTrace,
                    "Ribbon Load Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
                return null;
            }
        }

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
        /// Tên sheet mới dựa vào lựa chọn trên dropdown ddHyperlinkMode.</summary>
        public void OnAddSheetWithFormat(Office.IRibbonControl control)
        {
            Excel.Workbook wb = App?.ActiveWorkbook;
            if (wb == null) return;
            // _hyperlinkModeIndex: 0=None(→-1), 1=SEQ.xxx(→0), 2=SEQg.xxx(→1)
            Features.SheetEditingManager.AddSheetWithFormat(wb, _hyperlinkModeIndex - 1);
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

        // ── Nút thực thi ────────────────────────────────────────────────────

        /// <summary>Quét toàn bộ sheet hiện tại và tự động thêm hyperlink theo chế độ đã chọn.</summary>
        public void OnAutoHyperlink(Office.IRibbonControl control)
        {
            if (App?.ActiveWorkbook == null) return;
            Excel.Worksheet ws = App.ActiveSheet as Excel.Worksheet;
            if (ws == null) return;
            if (_hyperlinkModeIndex == 0)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Vui lòng chọn định dạng Sheet ID trên dropdown 'Sheet Name Format' trước khi thực hiện.",
                    "AutoHyperlink",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);
                return;
            }
            Features.HyperlinkManager.AutoAddHyperlinks(ws, _hyperlinkModeIndex - 1);
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

            // ddHyperlinkMode index: 0=None, 1=SEQ.xxx, 2=SEQg.xxx
            // Map sang VBA HYPERLINK_MODE: 1→0, 2→1
            int seqMode = _hyperlinkModeIndex - 1;
            if (seqMode < 0 || seqMode > 1)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Vui lòng chọn chế độ SEQ.xxx hoặc SEQg.xxx trên dropdown 'Sheet Name Format' trước khi thực hiện.",
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

            int seqMode = _hyperlinkModeIndex - 1;
            if (seqMode < 0 || seqMode > 1)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Vui lòng chọn chế độ SEQ.xxx hoặc SEQg.xxx trên dropdown 'Sheet Name Format' trước khi thực hiện.",
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

        /// <summary>Trả về chuỗi phiên bản add-in, ví dụ "v1.0.0".</summary>
        public string GetVersionLabel(Office.IRibbonControl control)
            => "v" + Assembly.GetExecutingAssembly().GetName().Version.ToString(3);

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
