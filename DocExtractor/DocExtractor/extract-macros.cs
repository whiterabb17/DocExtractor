using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
#if !DEBUG && !RELEASE
using Microsoft.Vbe.Interop;
#endif
namespace DocExtractor
{
    public class extract_macros
    {
        public static Document GetDoc(string docName)
        {
            var wordApp = new Microsoft.Office.Interop.Word.Application();
            return wordApp.Documents.Open(docName);
        }

        public static void CloseDoc(string docName)
        {
            var wordApp = new Microsoft.Office.Interop.Word.Application();
            wordApp.Documents[docName].Close();
        }

//        public static void RemoveMacrosAndSave(string filename)
//        {
//#if !DEBUG && !RELEASE
//            Spire.Doc.Document document = new Spire.Doc.Document();
//            //Load the Word document
//            document.LoadFromFile("Input.docm");

//            bool hasMacros = false;
//            //Detect if the document contains macros
//            hasMacros = document.IsContainMacro;

//            //If the result is true, remove the macros from the document
//            if (hasMacros)
//            {
                
//                document.ClearMacros();
//            }

//            //Save the document
//            document.SaveToFile($"macros_removed-{filename.Split("\\").Last()}", Spire.Doc.FileFormat.Docm);
//#else
//            Console.WriteLine("[!] DOES NOT WORK WITH THIS BUILD! PLEASE USE A WINDOWS BUILD (DebugWin/ReleaseWin)");
//            throw new NotImplementedException("This method is not implemented in debug mode");  
//#endif
//        }

        public static List<string> GetMacrosFromDoc(string docName)
        {
            List<string> macros = new List<string>();
#if !DEBUG && !RELEASE
            string tempdoc = $"temp{docName}";
            Document doc = GetDoc(tempdoc);
            
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
            //RemoveMacrosAndSave(docName);
#endif
            return macros;
        }
    }
}
