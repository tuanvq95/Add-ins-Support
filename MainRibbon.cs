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

        private Excel.Application _app;

        private int _hyperlinkMode = 0;

        #region IRibbonExtensibility

        public MainRibbon(Excel.Application app)
        {
            _app = app;
        }

        public string GetCustomUI(string ribbonID) => GetResourceText("AddinsSupport.MainRibbon.xml");


        #endregion

        #region Callbacks

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            _ribbon = ribbonUI;
        }

        // ── Hyperlink ──────────────────────────────────────────────────────────

        public void OnHyperlinkModeChanged(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {
            _hyperlinkMode = selectedIndex; // 0 hoặc 1
        }

        /// <summary>Quét toàn bộ sheet hiện tại và tự động thêm hyperlink.</summary>
        public void OnAutoHyperlink(Office.IRibbonControl control)
        {
            Excel.Worksheet ws = _app.Application.ActiveSheet as Excel.Worksheet;
            if (ws == null) return;
            Features.HyperlinkManager.AutoAddHyperlinks(ws);
        }

        // ── Sheet Name ─────────────────────────────────────────────────────────

        /// <summary>Mở dialog đổi tên sheet hàng loạt.</summary>
        public void OnRenameSheets(Office.IRibbonControl control)
        {
            Excel.Workbook wb = _app.Application.ActiveWorkbook;
            if (wb == null) return;
            Features.SheetNameManager.ShowRenameDialog(wb);
        }

        /// <summary>Đổi tên sheet theo nội dung cell A1.</summary>
        public void OnRenameByCell(Office.IRibbonControl control)
        {
            Excel.Workbook wb = _app.Application.ActiveWorkbook;
            if (wb == null) return;
            Features.SheetNameManager.RenameSheetsByCell(wb);
        }

        // ── VBA Macro ──────────────────────────────────────────────────────────

        /// <summary>Mở dialog chọn macro để nhúng vào workbook.</summary>
        public void OnInjectVba(Office.IRibbonControl control)
        {
            Excel.Workbook wb = _app.Application.ActiveWorkbook;
            if (wb == null) return;
            Features.VbaMacroManager.ShowMacroSelector(wb);
        }

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
