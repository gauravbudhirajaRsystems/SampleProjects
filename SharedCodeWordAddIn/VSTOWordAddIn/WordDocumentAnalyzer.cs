using System.Windows.Forms;

namespace VSTOWordAddIn
{
    public partial class WordDocumentAnalyzer : UserControl
    {
        public WordDocumentAnalyzer()
        {
            InitializeComponent();
        }

        Microsoft.Office.Interop.Word.Document document;

        private void btnUnicode_Click(object sender, System.EventArgs e)
        {
            document = Globals.ThisAddIn.Application.ActiveDocument;

            if (document != null)
            {
                string selectedText = document.ActiveWindow.Selection.Text;
                txtUnicode.Text = SharedCodeWordLibrary.WordOperations.GetUnicode(selectedText);
            }
        }

        private void btnCharCount_Click(object sender, System.EventArgs e)
        {
            document = Globals.ThisAddIn.Application.ActiveDocument;

            if (document != null)
            {
                string selectedText = document.ActiveWindow.Selection.Text;
                txtCharCount.Text = SharedCodeWordLibrary.WordOperations.GetCharCount(selectedText);
            }
        }

        private void btnWordCount_Click(object sender, System.EventArgs e)
        {

            document = Globals.ThisAddIn.Application.ActiveDocument;

            if (document != null)
            {
                string selectedText = document.ActiveWindow.Selection.Text;
                txtWordCount.Text = SharedCodeWordLibrary.WordOperations.GetWordCount(selectedText);
            }
        }
    }
}
