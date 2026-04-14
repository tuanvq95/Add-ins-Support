using System.Windows.Forms;

namespace AddinsSupport.Features
{
  /// <summary>
  /// Quản lý nhóm tiện ích mở rộng.
  /// Hiện tại các tính năng đang được phát triển; chỉ hiển thị thông báo chờ.
  /// </summary>
  public static class ExtensionsManager
  {
    /// <summary>
    /// Hiển thị thông báo "Đang phát triển" cho các tính năng chưa hoàn thiện.
    /// </summary>
    /// <param name="featureName">Tên tính năng (tuỳ chọn, để hiển thị trong thông báo).</param>
    public static void ShowComingSoon(string featureName = null)
    {
      string name = string.IsNullOrWhiteSpace(featureName) ? "Tính năng này" : featureName;

      MessageBox.Show(
          $"{name} đang được xây dựng.\n\n"
          + "Vui lòng chờ phiên bản cập nhật tiếp theo.\n"
          + "Cảm ơn bạn đã sử dụng Add-in Hỗ Trợ!",
          "Đang Phát Triển",
          MessageBoxButtons.OK,
          MessageBoxIcon.Information);
    }
  }
}
