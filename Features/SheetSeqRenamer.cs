using System;
using System.Collections.Generic;
using System.Globalization;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace AddinsSupport.Features
{
  // ═════════════════════════════════════════════════════════════════════════
  // Port từ VBA: AutoRenameSheet.bas
  //
  // AutoRenameSeqSheets — Shift tên sheet SEQ tại vị trí active lên +1
  //   Mode 0 (SeqDot)   → SEQ.xxx  / SEQ.xxx~yyy  (toàn bộ workbook)
  //   Mode 1 (SeqGroup) → SEQg.xxx / SEQg.xxx~yyy (chỉ trong nhóm g)
  //
  // CheckAndFixSeqOrder — Kiểm tra & chỉnh lại thứ tự đánh số theo vị trí tab
  // ═════════════════════════════════════════════════════════════════════════

  public static class SheetSeqRenamer
  {
    // ── Hằng đánh dấu không-có-range ────────────────────────────────────
    private const long NO_END = -1L;

    // ════════════════════════════════════════════════════════════════════
    // ĐIỂM VÀO CHÍNH
    // ════════════════════════════════════════════════════════════════════

    /// <summary>
    /// Shift tên sheet SEQ từ vị trí active sheet trở đi lên +1,
    /// tạo khoảng trống để chèn sheet mới tại vị trí active.
    /// <para>mode = 0 → SEQ.xxx (toàn workbook)</para>
    /// <para>mode = 1 → SEQg.xxx (chỉ trong nhóm của sheet active)</para>
    /// </summary>
    /// <param name="wb">Workbook đang làm việc.</param>
    /// <param name="mode">0 = SeqDot, 1 = SeqGroup — tương ứng VBA HYPERLINK_MODE.</param>
    public static void AutoRenameSeqSheets(Excel.Workbook wb, int mode)
    {
      if (wb == null) throw new ArgumentNullException("wb");

      if (mode == 0)
        RenameMode0(wb);
      else if (mode == 1)
        RenameMode1(wb);
      else
        MessageBox.Show("Mode không hợp lệ. Chọn SEQ.xxx hoặc SEQg.xxx trên dropdown.",
            "AutoRenameSeqSheets", MessageBoxButtons.OK, MessageBoxIcon.Warning);
    }

    /// <summary>
    /// Kiểm tra thứ tự đánh số SEQ theo vị trí tab.
    /// Nếu lệch → hỏi xác nhận rồi chỉnh lại tự động (bảo toàn độ rộng range).
    /// </summary>
    public static void CheckAndFixSeqOrder(Excel.Workbook wb, int mode)
    {
      if (wb == null) throw new ArgumentNullException("wb");

      if (mode == 0)
        FixOrderMode0(wb);
      else if (mode == 1)
        FixOrderMode1(wb);
      else
        MessageBox.Show("Mode không hợp lệ.",
            "CheckAndFixSeqOrder", MessageBoxButtons.OK, MessageBoxIcon.Warning);
    }

    // ════════════════════════════════════════════════════════════════════
    // MODE 0 — SEQ.xxx / SEQ.xxx~yyy  (port từ RenameMode0 VBA)
    // ════════════════════════════════════════════════════════════════════

    private static void RenameMode0(Excel.Workbook wb)
    {
      Excel.Worksheet activeSheet = wb.Application.ActiveSheet as Excel.Worksheet;
      if (activeSheet == null) return;

      string activeName = activeSheet.Name;

      // Lấy sheet ngay sau active (theo vị trí tab)
      Excel.Worksheet nextSheet = GetNextSheet(activeSheet);
      if (nextSheet == null)
      {
        MessageBox.Show("Không có sheet nào sau sheet active.",
            "AutoRenameSeqSheets (Mode 0)", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        return;
      }

      long nextStart, nextEnd;
      if (!ParseMode0(nextSheet.Name, out nextStart, out nextEnd))
      {
        MessageBox.Show(
            $"Sheet tiếp theo [{nextSheet.Name}] không đúng định dạng SEQ.xxx hoặc SEQ.xxx~yyy.",
            "AutoRenameSeqSheets (Mode 0)", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        return;
      }

      // Lưu tên gốc TRƯỚC khi shift — nextSheet.Name sẽ thay đổi sau vòng lặp
      string nextOriginalName = nextSheet.Name;

      // Thu thập tất cả sheet SEQ.x với x >= nextStart
      var entries = new List<SeqEntry>();
      foreach (Excel.Worksheet ws in wb.Worksheets)
      {
        long s, e;
        if (ParseMode0(ws.Name, out s, out e) && s >= nextStart)
          entries.Add(new SeqEntry { SheetName = ws.Name, Start = s, End = e });
      }

      if (entries.Count == 0) return;

      // Sắp xếp giảm dần để tránh xung đột tên
      entries.Sort((a, b) => b.Start.CompareTo(a.Start));

      // Shift +1 tất cả
      foreach (var entry in entries)
      {
        string newName = entry.End != NO_END
            ? $"SEQ.{entry.Start + 1}~{entry.End + 1}"
            : $"SEQ.{entry.Start + 1}";
        ((Excel.Worksheet)wb.Sheets[entry.SheetName]).Name = newName;
      }

      // Đặt tên active = tên GỐC của sheet tiếp theo (đã lưu trước khi shift)
      string activeNewName = StripRange(nextOriginalName);
      activeSheet.Name = activeNewName;

      MessageBox.Show(
          $"Hoàn thành! (Mode 0)\n"
          + $"• Sheet active [{activeName}] → [{activeNewName}]\n"
          + $"• Đã shift {entries.Count} sheet (+1)",
          "AutoRenameSeqSheets", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    // ════════════════════════════════════════════════════════════════════
    // MODE 1 — SEQg.xxx / SEQg.xxx~yyy  (port từ RenameMode1 VBA)
    // ════════════════════════════════════════════════════════════════════

    private static void RenameMode1(Excel.Workbook wb)
    {
      Excel.Worksheet activeSheet = wb.Application.ActiveSheet as Excel.Worksheet;
      if (activeSheet == null) return;

      string activeName = activeSheet.Name;

      Excel.Worksheet nextSheet = GetNextSheet(activeSheet);
      if (nextSheet == null)
      {
        MessageBox.Show("Không có sheet nào sau sheet active.",
            "AutoRenameSeqSheets (Mode 1)", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        return;
      }

      long activeGroup, nextStart, nextEnd;
      if (!ParseMode1(nextSheet.Name, out activeGroup, out nextStart, out nextEnd))
      {
        MessageBox.Show(
            $"Sheet tiếp theo [{nextSheet.Name}] không đúng định dạng SEQg.xxx hoặc SEQg.xxx~yyy.",
            "AutoRenameSeqSheets (Mode 1)", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        return;
      }

      // Lưu tên gốc TRƯỚC khi shift — nextSheet.Name sẽ thay đổi sau vòng lặp
      string nextOriginalName1 = nextSheet.Name;

      // Thu thập sheet trong cùng nhóm có start >= nextStart
      var entries = new List<SeqEntry>();
      foreach (Excel.Worksheet ws in wb.Worksheets)
      {
        long g, s, e;
        if (ParseMode1(ws.Name, out g, out s, out e) && g == activeGroup && s >= nextStart)
          entries.Add(new SeqEntry { SheetName = ws.Name, Group = g, Start = s, End = e });
      }

      if (entries.Count == 0) return;

      entries.Sort((a, b) => b.Start.CompareTo(a.Start));

      foreach (var entry in entries)
      {
        string newName = entry.End != NO_END
            ? $"SEQ{activeGroup}.{entry.Start + 1}~{entry.End + 1}"
            : $"SEQ{activeGroup}.{entry.Start + 1}";
        ((Excel.Worksheet)wb.Sheets[entry.SheetName]).Name = newName;
      }

      // Đặt tên active = tên GỐC của sheet tiếp theo (đã lưu trước khi shift)
      string activeNewName = StripRange(nextOriginalName1);
      activeSheet.Name = activeNewName;

      MessageBox.Show(
          $"Hoàn thành! (Mode 1 — Nhóm SEQ{activeGroup})\n"
          + $"• Sheet active [{activeName}] → [{activeNewName}]\n"
          + $"• Đã shift {entries.Count} sheet (+1)",
          "AutoRenameSeqSheets", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    // ════════════════════════════════════════════════════════════════════
    // FIX ORDER MODE 0  (port từ FixOrderMode0 VBA)
    // ════════════════════════════════════════════════════════════════════

    private static void FixOrderMode0(Excel.Workbook wb)
    {
      // Thu thập + sắp xếp theo vị trí tab
      var entries = new List<SeqEntry>();
      foreach (Excel.Worksheet ws in wb.Worksheets)
      {
        long s, e;
        if (ParseMode0(ws.Name, out s, out e))
          entries.Add(new SeqEntry { SheetName = ws.Name, TabIndex = ws.Index, Start = s, End = e });
      }

      if (entries.Count == 0)
      {
        MessageBox.Show("Không tìm thấy sheet nào dạng SEQ.xxx.",
            "CheckAndFixSeqOrder", MessageBoxButtons.OK, MessageBoxIcon.Information);
        return;
      }

      entries.Sort((a, b) => a.TabIndex.CompareTo(b.TabIndex));

      bool needFix = DetectOrderIssue(entries);
      if (!needFix)
      {
        MessageBox.Show("Thứ tự SEQ đã đúng, không cần chỉnh sửa.",
            "CheckAndFixSeqOrder", MessageBoxButtons.OK, MessageBoxIcon.Information);
        return;
      }

      if (MessageBox.Show(
          "Phát hiện số thứ tự bị lệch. Chỉnh lại tự động?\n(Thứ tự sẽ tính lại từ 1 theo vị trí tab)",
          "CheckAndFixSeqOrder", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
        return;

      // Tính tên mới, giữ nguyên độ rộng range
      var newNames = ComputeNewNames0(entries);

      // Đổi qua tên tạm để tránh xung đột
      for (int i = 0; i < entries.Count; i++)
        ((Excel.Worksheet)wb.Sheets[entries[i].SheetName]).Name = $"_TMP_{i}";
      for (int i = 0; i < entries.Count; i++)
        ((Excel.Worksheet)wb.Sheets[$"_TMP_{i}"]).Name = newNames[i];

      MessageBox.Show($"Hoàn thành! (Mode 0) Đã chỉnh lại {entries.Count} sheet.",
          "CheckAndFixSeqOrder", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    // ════════════════════════════════════════════════════════════════════
    // FIX ORDER MODE 1  (port từ FixOrderMode1 VBA)
    // ════════════════════════════════════════════════════════════════════

    private static void FixOrderMode1(Excel.Workbook wb)
    {
      // Thu thập tất cả sheet SEQg.xxx
      var all = new List<SeqEntry>();
      foreach (Excel.Worksheet ws in wb.Worksheets)
      {
        long g, s, e;
        if (ParseMode1(ws.Name, out g, out s, out e))
          all.Add(new SeqEntry { SheetName = ws.Name, Group = g, TabIndex = ws.Index, Start = s, End = e });
      }

      if (all.Count == 0)
      {
        MessageBox.Show("Không tìm thấy sheet nào dạng SEQg.xxx.",
            "CheckAndFixSeqOrder", MessageBoxButtons.OK, MessageBoxIcon.Information);
        return;
      }

      // Lấy danh sách nhóm duy nhất
      var groups = new HashSet<long>();
      foreach (var entry in all) groups.Add(entry.Group);

      bool needFix = false;
      foreach (long g in groups)
      {
        var grpEntries = all.FindAll(e => e.Group == g);
        grpEntries.Sort((a, b) => a.TabIndex.CompareTo(b.TabIndex));
        if (DetectOrderIssue(grpEntries)) { needFix = true; break; }
      }

      if (!needFix)
      {
        MessageBox.Show("Thứ tự tất cả các nhóm SEQ đã đúng.",
            "CheckAndFixSeqOrder", MessageBoxButtons.OK, MessageBoxIcon.Information);
        return;
      }

      if (MessageBox.Show(
          $"Phát hiện số thứ tự bị lệch. Chỉnh lại tất cả {groups.Count} nhóm?\n(Mỗi nhóm sẽ tính lại từ 1 theo vị trí tab)",
          "CheckAndFixSeqOrder", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
        return;

      // Đổi tất cả về tên tạm trước
      for (int i = 0; i < all.Count; i++)
        ((Excel.Worksheet)wb.Sheets[all[i].SheetName]).Name = $"_TMP_{i}";

      int fixCount = 0;
      foreach (long g in groups)
      {
        // Lấy lại bằng index tạm (tab index không đổi, tên đã đổi sang _TMP_i)
        var grpEntries = new List<SeqEntry>();
        for (int i = 0; i < all.Count; i++)
          if (all[i].Group == g)
            grpEntries.Add(new SeqEntry { SheetName = $"_TMP_{i}", TabIndex = all[i].TabIndex, Start = all[i].Start, End = all[i].End, Group = g });

        grpEntries.Sort((a, b) => a.TabIndex.CompareTo(b.TabIndex));

        long expectNum = 1;
        foreach (var entry in grpEntries)
        {
          string newName;
          if (entry.End != NO_END)
          {
            long width = entry.End - entry.Start + 1;
            newName = $"SEQ{g}.{expectNum}~{expectNum + width - 1}";
            expectNum += width;
          }
          else
          {
            newName = $"SEQ{g}.{expectNum}";
            expectNum++;
          }
          ((Excel.Worksheet)wb.Sheets[entry.SheetName]).Name = newName;
          fixCount++;
        }
      }

      MessageBox.Show($"Hoàn thành! Đã chỉnh lại {fixCount} sheet trong {groups.Count} nhóm.",
          "CheckAndFixSeqOrder", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    // ════════════════════════════════════════════════════════════════════
    // PARSE HELPERS  (port từ ParseMode0 / ParseMode1 VBA)
    // ════════════════════════════════════════════════════════════════════

    /// <summary>
    /// Parse "SEQ.xxx" hoặc "SEQ.xxx~yyy".<br/>
    /// outEnd = -1 nếu không có range.
    /// </summary>
    internal static bool ParseMode0(string name, out long outStart, out long outEnd)
    {
      outStart = 0;
      outEnd = NO_END;

      if (!name.StartsWith("SEQ.", StringComparison.OrdinalIgnoreCase)) return false;

      string rest = name.Substring(4);

      // Nếu còn dấu chấm → là SEQg.xxx (Mode 1), bỏ qua
      if (rest.IndexOf('.') >= 0) return false;

      return TryParseSeqNum(rest, ref outStart, ref outEnd);
    }

    /// <summary>
    /// Parse "SEQg.xxx" hoặc "SEQg.xxx~yyy".<br/>
    /// outEnd = -1 nếu không có range.
    /// </summary>
    internal static bool ParseMode1(string name, out long outGroup, out long outStart, out long outEnd)
    {
      outGroup = 0;
      outStart = 0;
      outEnd = NO_END;

      if (!name.StartsWith("SEQ", StringComparison.OrdinalIgnoreCase)) return false;

      string rest = name.Substring(3); // vd: "2.22" hoặc "2.22~25"

      int dot = rest.IndexOf('.');
      if (dot <= 0) return false;  // dot=0 → "SEQ.xxx" (Mode 0)

      long grp;
      if (!TryParseLong(rest.Substring(0, dot), out grp)) return false;

      string numPart = rest.Substring(dot + 1);
      if (numPart.IndexOf('.') >= 0) return false;  // không cho phép thêm dấu chấm

      outGroup = grp;
      return TryParseSeqNum(numPart, ref outStart, ref outEnd);
    }

    // ── TryParseSeqNum: phân tích "xxx" hoặc "xxx~yyy" ──────────────────
    private static bool TryParseSeqNum(string s, ref long outStart, ref long outEnd)
    {
      int tilde = s.IndexOf('~');
      if (tilde > 0)
      {
        long a, b;
        if (TryParseLong(s.Substring(0, tilde), out a) && TryParseLong(s.Substring(tilde + 1), out b))
        { outStart = a; outEnd = b; return true; }
      }
      else
      {
        long a;
        if (TryParseLong(s, out a))
        { outStart = a; return true; }
      }
      return false;
    }

    // ── Chuyển chuỗi Excel (có thể "1.0") thành long ────────────────────
    private static bool TryParseLong(string s, out long result)
    {
      result = 0;
      if (string.IsNullOrEmpty(s)) return false;
      if (long.TryParse(s, out result)) return true;
      if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out double d)
          && d == Math.Floor(d) && d > 0)
      { result = (long)d; return true; }
      return false;
    }

    // ════════════════════════════════════════════════════════════════════
    // MISC HELPERS
    // ════════════════════════════════════════════════════════════════════

    // Lấy sheet kế tiếp theo vị trí tab
    private static Excel.Worksheet GetNextSheet(Excel.Worksheet ws)
    {
      try { return ws.Next as Excel.Worksheet; }
      catch { return null; }
    }

    // Xóa phần "~yyy" khỏi tên nếu là range (chỉ giữ prefix)
    private static string StripRange(string name)
    {
      int tilde = name.IndexOf('~');
      return tilde > 0 ? name.Substring(0, tilde) : name;
    }

    // Kiểm tra xem danh sách entries (đã sắp theo tab) có đúng thứ tự số không
    private static bool DetectOrderIssue(List<SeqEntry> sorted)
    {
      long expect = 1;
      foreach (var e in sorted)
      {
        if (e.Start != expect) return true;
        expect = e.End != NO_END ? e.End + 1 : expect + 1;
      }
      return false;
    }

    // Tính tên mới cho Mode 0 (giữ nguyên độ rộng range)
    private static List<string> ComputeNewNames0(List<SeqEntry> sorted)
    {
      var result = new List<string>(sorted.Count);
      long expect = 1;
      foreach (var e in sorted)
      {
        if (e.End != NO_END)
        {
          long width = e.End - e.Start + 1;
          result.Add($"SEQ.{expect}~{expect + width - 1}");
          expect += width;
        }
        else
        {
          result.Add($"SEQ.{expect}");
          expect++;
        }
      }
      return result;
    }

    // ── Data class nội bộ ────────────────────────────────────────────────
    private sealed class SeqEntry
    {
      public string SheetName { get; set; }
      public long Group { get; set; }
      public long Start { get; set; }
      public long End { get; set; }  // NO_END (-1) = số đơn
      public int TabIndex { get; set; }
    }
  }
}
