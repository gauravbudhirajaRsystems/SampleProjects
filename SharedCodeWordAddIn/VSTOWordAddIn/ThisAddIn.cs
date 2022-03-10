namespace VSTOWordAddIn
{
    public partial class ThisAddIn
    {
        private WordDocumentAnalyzer wordDocumentAnalyzer;
        private Microsoft.Office.Tools.CustomTaskPane customTaskPane;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            wordDocumentAnalyzer = new WordDocumentAnalyzer();
            customTaskPane = CustomTaskPanes.Add(wordDocumentAnalyzer, "Word Analyzer");
            customTaskPane.Visible = true;
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
            Startup += new System.EventHandler(ThisAddIn_Startup);
            Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
