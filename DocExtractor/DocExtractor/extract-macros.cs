using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using VBIDE;

namespace DocExtractor
{
    public class extract_macros
    {
        public static Document GetDoc(string docName)
        {
            var wordApp = new Microsoft.Office.Interop.Word.Application();

            // 2. Specify the path to your Word document.

            // 3. Open the document.
            return wordApp.Documents.Open(docName);
        }

        public static void CloseDoc(string docName)
        {
            var wordApp = new Microsoft.Office.Interop.Word.Application();
            wordApp.Documents[docName].Close();
        }

        public static List<string> GetMacrosFromDoc(string docName)
        {
            Document doc = GetDoc(docName);
            
            List<string> macros = new List<string>();

            VBProject prj;
            CodeModule code;
            string composedFile;

            prj = doc.VBProject;
            foreach (VBComponent comp in prj.VBComponents)
            {
                code = comp.CodeModule;

                // Put the name of the code module at the top
                composedFile = comp.Name + Environment.NewLine;

                // Loop through the (1-indexed) lines
                for (int i = 0; i < code.CountOfLines; i++)
                {
                    composedFile += code.get_Lines(i + 1, 1) + Environment.NewLine;
                }

                // Add the macro to the list
                macros.Add(composedFile);
            }
            CloseDoc(docName);
            return macros;
        }
    }
}
