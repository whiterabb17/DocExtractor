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
                args = System.IO.Directory.GetFiles(System.IO.Directory.GetCurrentDirectory(), "*.docx");
                args += System.IO.Directory.GetFiles(System.IO.Directory.GetCurrentDirectory(), "*.doc");
            }
            ExtractBetweenParagraphs(args);
            Console.WriteLine("\n\n [!] Extraction complete!");
            Console.ReadLine();
        }

        private static void ExtractBetweenParagraphs(string[] files, int starting = 0)
        {
            foreach (var file in files)
            {
                Console.WriteLine($"Extracting content from {file}...\n\n");
                Document doc = new Document(file);

                // Gather the nodes (the GetChild method uses 0-based index)
                Paragraph startPara = (Paragraph)doc.FirstSection.Body.GetChild(NodeType.Paragraph, starting, true);
                Paragraph endPara = (Paragraph)doc.FirstSection.Body.GetChild(NodeType.Paragraph, (doc.FirstSection.Body.GetChildNodes(NodeType.Paragraph, true).Count - 1), true);

                // Extract the content between these nodes in the document. Include these markers in the extraction.
                ArrayList extractedNodes = extract_text.ExtractContent(startPara, endPara, true);

                // Insert the content into a new document and save it to disk.
                Document dstDoc = text_extraction_helper.GenerateDocument(doc, extractedNodes);
                dstDoc.Save($"extracted-{file.Split("\\").Last()}");
            }
            
        }
    }
}