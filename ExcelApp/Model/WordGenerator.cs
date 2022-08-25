using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Word = Microsoft.Office.Interop.Word;

namespace ExcelApp.Model
{
    class WordGenerator
    {
        string _wordText;

        public WordGenerator(string wordtext)
        {
            _wordText = wordtext;
        }
        public void Generate()
        {
            Word.Document doc = null;
            try
            {
                string path = "c:\\1\\1.docx";
                Word.Application application = new Word.Application();
                doc = application.Documents.Open(path);
                doc.Activate();

                Word.Bookmarks bookmarks = doc.Bookmarks;
                Word.Range range;
                foreach (Word.Bookmark bookmark in bookmarks)
                {
                    range = bookmark.Range;
                    range.Text = _wordText;
                }
                doc.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                doc.Close();
            }
        }
    }
}
