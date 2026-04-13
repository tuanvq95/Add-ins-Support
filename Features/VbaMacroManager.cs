using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace AddinsSupport.Features
{
    /// <summary>
    /// Nhúng macro VBA vào workbook hiện tại.
    /// Yêu cầu: Excel → File → Options → Trust Center → Trust Center Settings
    ///          → Macro Settings → bật "Trust access to the VBA project object model".
    /// </summary>
    public static class VbaMacroManager
    {
        // ─── Danh sách macro sẵn có ───────────────────────────────────────────

        public static readonly IReadOnlyList<MacroDefinition> AvailableMacros = new List<MacroDefinition>
        {
            new MacroDefinition(
                "AutoFitAllColumns",
                "Tự động điều chỉnh độ rộng tất cả cột",
@"Sub AutoFitAllColumns()
    ActiveSheet.Cells.EntireColumn.AutoFit
    MsgBox ""Đã tự động điều chỉnh độ rộng tất cả cột!"", vbInformation
End Sub"),

            new MacroDefinition(
                "HighlightDuplicates",
                "Tô vàng các giá trị trùng lặp trong vùng chọn",
@"Sub HighlightDuplicates()
    Dim rng As Range, cell As Range, dict As Object
    Set dict = CreateObject(""Scripting.Dictionary"")
    Set rng = Selection

    ' Đếm lần xuất hiện
    For Each cell In rng
        Dim key As String
        key = CStr(cell.Value)
        If dict.Exists(key) Then
            dict(key) = dict(key) + 1
        Else
            dict.Add key, 1
        End If
    Next cell

    ' Tô màu
    For Each cell In rng
        If dict(CStr(cell.Value)) > 1 Then
            cell.Interior.Color = RGB(255, 255, 0)
        End If
    Next cell

    MsgBox ""Đã tô màu các giá trị trùng lặp."", vbInformation
End Sub"),

            new MacroDefinition(
                "ExportActiveSheetToCSV",
                "Xuất sheet hiện tại ra file CSV",
@"Sub ExportActiveSheetToCSV()
    Dim ws As Worksheet
    Dim savePath As String
    Set ws = ActiveSheet

    savePath = Application.GetSaveAsFilename( _
        InitialFileName:=ws.Name & "".csv"", _
        FileFilter:=""CSV Files (*.csv), *.csv"")

    If savePath = ""False"" Then Exit Sub

    Dim tempWb As Workbook
    ws.Copy
    Set tempWb = ActiveWorkbook
    Application.DisplayAlerts = False
    tempWb.SaveAs Filename:=savePath, FileFormat:=xlCSV, Local:=True
    tempWb.Close SaveChanges:=False
    Application.DisplayAlerts = True

    MsgBox ""Đã xuất ra: "" & savePath, vbInformation
End Sub"),

            new MacroDefinition(
                "CreateSheetIndex",
                "Tạo sheet mục lục với hyperlink đến tất cả sheet",
@"Sub CreateSheetIndex()
    Dim wb As Workbook
    Dim indexWs As Worksheet
    Dim ws As Worksheet
    Dim row As Long

    Set wb = ActiveWorkbook

    ' Xóa sheet Index cũ nếu có
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Sheets(""Mục Lục"").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' Tạo sheet Index mới ở đầu
    Set indexWs = wb.Sheets.Add(Before:=wb.Sheets(1))
    indexWs.Name = ""Mục Lục""

    indexWs.Range(""A1"").Value = ""Mục Lục Các Sheet""
    indexWs.Range(""A1"").Font.Bold = True
    indexWs.Range(""A1"").Font.Size = 14

    row = 3
    For Each ws In wb.Worksheets
        If ws.Name <> ""Mục Lục"" Then
            indexWs.Hyperlinks.Add _
                Anchor:=indexWs.Cells(row, 1), _
                Address:="""", _
                SubAddress:=""'"" & ws.Name & ""'!A1"", _
                TextToDisplay:=ws.Name
            row = row + 1
        End If
    Next ws

    indexWs.Columns(""A"").AutoFit
    MsgBox ""Đã tạo sheet Mục Lục với "" & (row - 3) & "" liên kết."", vbInformation
End Sub"),

            new MacroDefinition(
                "TrimAllCells",
                "Xóa khoảng trắng thừa trong tất cả cell có chứa text",
@"Sub TrimAllCells()
    Dim ws As Worksheet
    Dim cell As Range
    Dim count As Long

    Set ws = ActiveSheet
    count = 0

    Application.ScreenUpdating = False
    For Each cell In ws.UsedRange
        If cell.HasFormula = False And VarType(cell.Value) = vbString Then
            Dim cleaned As String
            cleaned = Application.WorksheetFunction.Trim(cell.Value)
            If cleaned <> cell.Value Then
                cell.Value = cleaned
                count = count + 1
            End If
        End If
    Next cell
    Application.ScreenUpdating = True

    MsgBox ""Đã làm sạch "" & count & "" cell."", vbInformation
End Sub"),

            new MacroDefinition(
                "AutoNumberColumn",
                "Đánh số thứ tự tự động cho cột A từ hàng 2",
@"Sub AutoNumberColumn()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, ""B"").End(xlUp).Row

    If lastRow < 2 Then
        MsgBox ""Không có dữ liệu để đánh số (cột B trống)."", vbExclamation
        Exit Sub
    End If

    For i = 2 To lastRow
        ws.Cells(i, 1).Value = i - 1
    Next i

    MsgBox ""Đã đánh số thứ tự cho "" & (lastRow - 1) & "" dòng."", vbInformation
End Sub")
        };

        // ─── Public API ───────────────────────────────────────────────────────

        /// <summary>
        /// Hiển thị dialog chọn macro, sau đó nhúng các macro được chọn vào workbook.
        /// </summary>
        public static void ShowMacroSelector(Excel.Workbook wb)
        {
            if (wb == null) throw new ArgumentNullException("wb");

            using (var form = new Forms.MacroSelectorForm())
            {
                if (form.ShowDialog() != DialogResult.OK) return;
                InjectMacros(wb, form.SelectedMacros);
            }
        }

        /// <summary>
        /// Nhúng danh sách macro đã chọn vào VBA project của workbook.
        /// </summary>
        public static void InjectMacros(Excel.Workbook wb, IList<MacroDefinition> macros)
        {
            if (macros == null || macros.Count == 0) return;

            // Kiểm tra quyền truy cập VBA Project Object Model
            bool canAccess;
            try
            {
                var _ = wb.VBProject;
                canAccess = true;
            }
            catch (Exception)
            {
                canAccess = false;
            }

            if (!canAccess)
            {
                MessageBox.Show(
                    "Không thể truy cập VBA Project.\n\n" +
                    "Vui lòng bật: Excel → File → Options → Trust Center → Trust Center Settings\n" +
                    "→ Macro Settings → Tích chọn \"Trust access to the VBA project object model\".",
                    "Quyền Truy Cập VBA",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            var vbaProject = wb.VBProject;
            int injected = 0;

            foreach (var macro in macros)
            {
                try
                {
                    // Tìm module tên "AddInSupport", tạo mới nếu chưa có
                    Microsoft.Vbe.Interop.VBComponent module = null;
                    foreach (Microsoft.Vbe.Interop.VBComponent comp in vbaProject.VBComponents)
                    {
                        if (comp.Name == "AddInSupport")
                        {
                            module = comp;
                            break;
                        }
                    }

                    if (module == null)
                    {
                        module = vbaProject.VBComponents.Add(
                            Microsoft.Vbe.Interop.vbext_ComponentType.vbext_ct_StdModule);
                        module.Name = "AddInSupport";
                    }

                    // Kiểm tra nếu macro đã tồn tại thì bỏ qua
                    string existingCode = module.CodeModule.Lines[1, module.CodeModule.CountOfLines];
                    if (existingCode.Contains("Sub " + macro.Name + "(")) continue;

                    // Thêm macro vào cuối module
                    int lineCount = module.CodeModule.CountOfLines;
                    module.CodeModule.InsertLines(lineCount + 1, Environment.NewLine + macro.Code);
                    injected++;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(
                        $"Lỗi khi nhúng macro '{macro.Name}': {ex.Message}",
                        "Lỗi",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
            }

            if (injected > 0)
            {
                MessageBox.Show(
                    $"Đã nhúng {injected} macro vào module 'AddInSupport'.\n" +
                    "Nhấn Alt+F8 trong Excel để chạy macro.",
                    "Nhúng Macro Thành Công",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
        }
    }

    // ─── Data Class ───────────────────────────────────────────────────────────

    public sealed class MacroDefinition
    {
        public string Name { get; }
        public string Description { get; }
        public string Code { get; }

        public MacroDefinition(string name, string description, string code)
        {
            Name = name;
            Description = description;
            Code = code;
        }

        public override string ToString() => Description;
    }
}
