using ExcelDna.Integration;

namespace AddinsSupport
{
    public partial class ThisAddIn
    {
        private void InternalStartup() { }
    }

    public class AddInEntryPoint : IExcelAddIn
    {
        public void AutoOpen() { }
        public void AutoClose() { }
    }
}
