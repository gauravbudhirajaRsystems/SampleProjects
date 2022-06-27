using System.Windows.Forms;

namespace VSTOWordAddIn
{
    public partial class WordDocumentAnalyzer : UserControl
    {
        private readonly Microsoft.Office.Interop.Word.Document _document;

        public WordDocumentAnalyzer()
        {
            InitializeComponent();
            _document = Globals.ThisAddIn.Application.ActiveDocument;

        }

        private void Unicode_Click(object sender, System.EventArgs e)
        {
            if (_document != null)
            {
                string selectedText = _document.ActiveWindow.Selection.Text;
                txtUnicode.Text = SharedCodeWordLibrary.WordOperations.GetUnicode(selectedText);
            }
        }

        private void CharCount_Click(object sender, System.EventArgs e)
        {
            if (_document != null)
            {
                string selectedText = _document.ActiveWindow.Selection.Text;
                txtCharCount.Text = SharedCodeWordLibrary.WordOperations.GetCharCount(selectedText);
            }
        }

        private void WordCount_Click(object sender, System.EventArgs e)
        {
            if (_document != null)
            {
                string selectedText = _document.ActiveWindow.Selection.Text;
                txtWordCount.Text = SharedCodeWordLibrary.WordOperations.GetWordCount(selectedText);
            }
        }
    }
}
