using System.Drawing;

namespace AddinsSupport.Features
{
  /// <summary>
  /// Lưu trữ cài đặt tô màu cho tính năng Tô Màu Vùng Chọn.
  /// Tồn tại trong suốt phiên làm việc (session-level state).
  /// </summary>
  public static class ColorSelectionSettings
  {
    /// <summary>Giới hạn tô màu theo dải cột hay không.</summary>
    public static bool UseColumnRange { get; set; } = false;

    /// <summary>Cột bắt đầu (1-based). Chỉ dùng khi <see cref="UseColumnRange"/> = true.</summary>
    public static int ColFrom { get; set; } = 1;

    /// <summary>Cột kết thúc (1-based). Chỉ dùng khi <see cref="UseColumnRange"/> = true.</summary>
    public static int ColTo { get; set; } = 10;

    /// <summary>Màu nền (System.Drawing). Mặc định: vàng nhạt #FFFFCC.</summary>
    public static Color FillColor { get; set; } = Color.FromArgb(0xFF, 0xFF, 0xCC);

    /// <summary>Phần trăm zoom áp dụng khi chuẩn hóa sheet. Mặc định: 100%.</summary>
    public static int SheetZoomPercent { get; set; } = 100;

    /// <summary>
    /// Chuyển <see cref="FillColor"/> sang giá trị BGR dùng cho Excel COM
    /// (<c>Interior.Color</c>).
    /// Công thức: R | (G &lt;&lt; 8) | (B &lt;&lt; 16) — tương đương VBA RGB().
    /// </summary>
    public static int FillColorBgr
        => FillColor.R | (FillColor.G << 8) | (FillColor.B << 16);
  }
}
