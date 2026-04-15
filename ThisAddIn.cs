using Office = Microsoft.Office.Core;

namespace AddinsSupport
{
    public partial class ThisAddIn
    {
        private void InternalStartup() { }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new MainRibbon();
        }
    }
}
