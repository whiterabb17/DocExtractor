using Aspose.Words;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocExtractor
{
    class extract_text
    {
        public static ArrayList ExtractContent(Node startNode, Node endNode, bool isInclusive)
        {
            // First check that the nodes passed to this method are valid for use.
            text_extraction_helper.VerifyParameterNodes(startNode, endNode);

            // Create a list to store the extracted nodes.
            ArrayList nodes = new ArrayList();

            // Keep a record of the original nodes passed to this method so we can split marker nodes if needed.
            Node originalStartNode = startNode;
            Node originalEndNode = endNode;

            // Extract content based on block level nodes (paragraphs and tables). Traverse through parent nodes to find them.
            // We will split the content of first and last nodes depending if the marker nodes are inline
            while (startNode.ParentNode.NodeType != NodeType.Body)
                startNode = startNode.ParentNode;

            while (endNode.ParentNode.NodeType != NodeType.Body)
                endNode = endNode.ParentNode;

            bool isExtracting = true;
            bool isStartingNode = true;
            bool isEndingNode = false;
            // The current node we are extracting from the document.
            Node currNode = startNode;

            // Begin extracting content. Process all block level nodes and specifically split the first and last nodes when needed so paragraph formatting is retained.
            // Method is little more complex than a regular extractor as we need to factor in extracting using inline nodes, fields, bookmarks etc as to make it really useful.
            while (isExtracting)
            {
                // Clone the current node and its children to obtain a copy.
                Node cloneNode = currNode.Clone(true);
                isEndingNode = currNode.Equals(endNode);

                if ((isStartingNode || isEndingNode) && cloneNode.IsComposite)
                {
                    // We need to process each marker separately so pass it off to a separate method instead.
                    if (isStartingNode)
                    {
                        text_extraction_helper.ProcessMarker((CompositeNode)cloneNode, nodes, originalStartNode, isInclusive, isStartingNode, isEndingNode);
                        isStartingNode = false;
                    }

                    // Conditional needs to be separate as the block level start and end markers maybe the same node.
                    if (isEndingNode)
                    {
                        text_extraction_helper.ProcessMarker((CompositeNode)cloneNode, nodes, originalEndNode, isInclusive, isStartingNode, isEndingNode);
                        isExtracting = false;
                    }
                }
                else
                    // Node is not a start or end marker, simply add the copy to the list.
                    nodes.Add(cloneNode);

                // Move to the next node and extract it. If next node is null that means the rest of the content is found in a different section.
                if (currNode.NextSibling == null && isExtracting)
                {
                    // Move to the next section.
                    Section nextSection = (Section)currNode.GetAncestor(NodeType.Section).NextSibling;
                    currNode = nextSection.Body.FirstChild;
                }
                else
                {
                    // Move to the next node in the body.
                    currNode = currNode.NextSibling;
                }
            }

            // Return the nodes between the node markers.
            return nodes;
        }
    }
}
