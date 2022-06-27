using System;
using System.Windows.Forms;

namespace WordVSTOAddIn
{
    public partial class WordDocumentAnalyzer : UserControl
    {
        public WordDocumentAnalyzer()
        {
            InitializeComponent();
        }

        private void btnShowText_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Word.Document document;

            document = Globals.ThisAddIn.Application.ActiveDocument;

            if (document != null)
            {
                //int startIndex = document.ActiveWindow.Selection.Range.Start;
                //int endIndex = document.ActiveWindow.Selection.Range.End;

                //string highlightedText = document.Range(startIndex, endIndex).Text;
                string selectedText = document.ActiveWindow.Selection.Text;

                //document.Range(startIndex, endIndex).Font.Size = 14;
                //document.Range(startIndex, endIndex).Font.Name = "Arial";

                txtResult.Text = WordSharedCodeLibrary.WordOperations.GetHighlightedTextUnicode(selectedText);
            }
        }
    }
}
