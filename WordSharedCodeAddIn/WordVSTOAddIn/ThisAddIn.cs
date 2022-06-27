namespace WordVSTOAddIn
{
    public partial class ThisAddIn
    {
        private WordDocumentAnalyzer control;
        private Microsoft.Office.Tools.CustomTaskPane taskpane;


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.control = new WordDocumentAnalyzer();
            this.taskpane = this.CustomTaskPanes.Add(this.control, "Word Analyzer");
            this.taskpane.Visible = true;
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
