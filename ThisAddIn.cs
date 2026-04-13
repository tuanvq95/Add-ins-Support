using System;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace AddinsSupport
{
    public partial class ThisAddIn
    {
        /// <summary>
        /// Đăng ký Ribbon XML với Excel.
        /// </summary>
        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject() => new MainRibbon(this.Application);

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            // Theo dõi sự kiện thay đổi cell để tự động nhận diện URL
            this.Application.SheetChange += Application_SheetChange;
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            this.Application.SheetChange -= Application_SheetChange;
        }

        /// <summary>
        /// Khi người dùng nhập URL vào một cell, tự động chuyển thành hyperlink.
        /// </summary>
        private void Application_SheetChange(object sh, Excel.Range target)
        {
            // Chỉ xử lý khi chỉnh sửa một cell duy nhất để tránh chậm
            if (target.Cells.Count != 1) return;

            string value = target.Value2 as string;
            if (string.IsNullOrEmpty(value)) return;

            string trimmed = value.Trim();
            if (Features.HyperlinkManager.IsUrl(trimmed) && target.Hyperlinks.Count == 0)
            {
                string url = trimmed.StartsWith("www.", StringComparison.OrdinalIgnoreCase)
                    ? "http://" + trimmed
                    : trimmed;

                try
                {
                    target.Worksheet.Hyperlinks.Add(
                        Anchor: target,
                        Address: url,
                        TextToDisplay: trimmed);
                }
                catch
                {
                    // Bỏ qua nếu không thể thêm hyperlink (ví dụ: cell protected)
                }
            }
        }

        #region VSTO generated code
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}
