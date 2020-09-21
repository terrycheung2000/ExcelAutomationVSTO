using Excel = Microsoft.Office.Interop.Excel;

namespace Data_Load
{
    public partial class ThisAddIn
    {
        Excel.Workbook main;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            main = this.Application.ActiveWorkbook;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion

    }



}
