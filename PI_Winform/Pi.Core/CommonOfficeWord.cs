using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace Pi.Core
{
    /// <summary>
    /// File này chứa các hàm liên quan tới file word
    /// </summary>
    public static class CommonOfficeWord
    {
        //public static void WriteFunctionHere()
        //{

        public static void FindAndReplace(Word.Application wordApp, object findText, object replaceWithText)
        {
            object matchCase = true;
            object matchWholeWord = true;
            object matchWillCards = false;
            object matchSoundLike = false;
            object matchAllForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiactitics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;

            wordApp.Selection.Find.Execute(ref findText,
                ref matchCase, ref matchWholeWord,
                ref matchWillCards, ref matchSoundLike,
                ref matchAllForms, ref forward, ref wrap,
                ref format, ref replaceWithText,
                ref replace, ref matchKashida,
                ref matchDiactitics, ref matchAlefHamza,
                ref matchControl);
        }

        public static void CreateWordDocument(object fileName, object saveAs, string a, string b, string c)
        {
            Word.Application wordApp = new Word.Application();
            object missing = Missing.Value;
            Word.Document aDoc = null;

            if (File.Exists((string)fileName))
            {
                object readOnly = false;
                object isVisible = false;
                wordApp.Visible = false;

                aDoc = wordApp.Documents.Open(ref fileName, ref missing, ref readOnly, 
                                    ref missing, ref missing, ref missing, ref missing, 
                                    ref missing, ref missing, ref missing, ref missing, 
                                    ref missing, ref missing, ref missing, ref missing, ref missing);

                aDoc.Activate();

                //Find and Replace
                FindAndReplace(wordApp, "[PI_PARAM1]", a);
                FindAndReplace(wordApp, "[PI_PARAM2]", b);
                FindAndReplace(wordApp, "[PI_PARAM3]", c);
            }
            else
            {
                MessageBox.Show("File not found!");
            }

            aDoc.SaveAs2(ref saveAs, ref missing, ref missing, ref missing, ref missing, 
                        ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

            aDoc.Close(ref missing, ref missing, ref missing);
            wordApp.Quit();
            MessageBox.Show("File created!");
        }
    }
}
