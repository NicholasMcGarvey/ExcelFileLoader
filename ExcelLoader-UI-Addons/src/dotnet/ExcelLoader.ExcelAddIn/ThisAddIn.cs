
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace ExcelLoader.ExcelAddIn
{
    public partial class ThisAddIn
    {
        private Ribbon _ribbon;

        private void ThisAddIn_Startup(object sender, System.EventArgs e) { }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e) { }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            _ribbon = new Ribbon();
            return _ribbon;
        }

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
    }
}
