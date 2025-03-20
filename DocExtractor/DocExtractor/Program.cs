using Aspose.Words;
using System.Collections;

namespace DocExtractor
{ 
    static class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("No file was provided...\nAttempting to Extract all '.docx' files in current directory");
                var __argsx = System.IO.Directory.GetFiles(System.IO.Directory.GetCurrentDirectory(), "*.docx");
                string[] argsx = { };
                var __args = System.IO.Directory.GetFiles(System.IO.Directory.GetCurrentDirectory(), "*.doc");
                foreach (string _argx in __argsx)
                {
                    if (!_argx.Contains("extracted"))
                        ExtractBetweenParagraphs(new string[] { _argx });
                }
                foreach (string _args in __args)
                { 
                    if (!_args.Contains("extracted"))
                        ExtractBetweenParagraphs(new string[] { _args });
                }
            }
            ExtractBetweenParagraphs(args);
            Console.WriteLine("\n\n [!] Extraction complete!");
            Console.ReadLine();
        }

        private static void ExtractBetweenParagraphs(string[] files, int starting = 0)
        {
            foreach (var file in files)
            {

                //string tempdoc = $"temp{file.Split("\\").Last()}";
                //File.Copy(file, tempdoc);
                Console.WriteLine($"Extracting content from {file}...\n\n");
                Document doc = new Document(file);

                // Gather the nodes (the GetChild method uses 0-based index)
                Paragraph startPara = (Paragraph)doc.FirstSection.Body.GetChild(NodeType.Paragraph, starting, true);
                Paragraph endPara = (Paragraph)doc.FirstSection.Body.GetChild(NodeType.Paragraph, (doc.FirstSection.Body.GetChildNodes(NodeType.Paragraph, true).Count - 1), true);

                // Extract the content between these nodes in the document. Include these markers in the extraction.
                ArrayList extractedNodes = extract_text.ExtractContent(startPara, endPara, true);

                // Insert the content into a new document and save it to disk.
                Document dstDoc = text_extraction_helper.GenerateDocument(doc, extractedNodes);

                Console.WriteLine("Would you like to save to a new document? (y/n)");
                var response = Console.ReadLine();
                if (response.ToLower() == "y")
                    dstDoc.Save($"extracted-{file.Split("\\").Last()}");
                dstDoc.Cleanup();
#if !DEBUG && !RELEAE
                Console.WriteLine("Would you like to attempt to remove macros? (y/n)");
                var ans = Console.ReadLine();
                if (ans.ToLower() == "y")
                { 
                    try
                    {
                        //var tempfile = File.ReadAllBytes(file);
                        //File.WriteAllBytes($"tempmacro.{file.Split(".").Last()}", tempfile);
                        var macros = extract_macros.GetMacrosFromDoc(tempdoc);
                        if (macros.Count > 0)
                        {
                            File.WriteAllLines($"extracted-macros-{file.Split("\\").Last().Replace("docx", "txt").Replace("doc", "txt")}", macros);
                            Console.WriteLine($"Extracted Macros:\nWritten to: extracted-{file.Split("\\").Last().Replace("docx", "txt").Replace("doc", "txt")}\n\nContent:\n{macros}");
                        }
                        //File.Delete($"tempmacro.{file.Split(".").Last()}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"ERROR: {ex.Message}");
                        File.Delete($"tempmacro.{file.Split(".").Last()}");
                    }
                }
#endif
                //File.Delete(tempdoc);
            }
            Console.WriteLine("Extraction Processes Complete");
            Console.ReadLine();
        }
    }
}