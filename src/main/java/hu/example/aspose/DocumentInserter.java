package hu.example.aspose;

import com.aspose.words.CompositeNode;
import com.aspose.words.Document;
import com.aspose.words.Node;
import com.aspose.words.NodeImporter;
import com.aspose.words.NodeType;
import com.aspose.words.Paragraph;
import com.aspose.words.Section;

public class DocumentInserter
{    
    /**
     * Inserts content of the external document after the specified node.
     * Section breaks and section formatting of the inserted document are
     * ignored.
     *
     * @param insertionDestination Node in the destination document after which the content
     *                        should be inserted. This node should be a block level node
     *                        (paragraph or table).
     * @param docToInsert     The document to insert.
     */
    public static void insertDocument(Node insertionDestination, Document docToInsert, int importFormatMode)
    {
        // Make sure that the node is either a paragraph or table.
        if (((insertionDestination.getNodeType()) == (NodeType.PARAGRAPH)) || ((insertionDestination.getNodeType()) == (NodeType.TABLE)))
        {
            // We will be inserting into the parent of the destination paragraph.
            CompositeNode<?> dstStory = insertionDestination.getParentNode();

            // This object will be translating styles and lists during the import.
            NodeImporter importer = new NodeImporter(docToInsert, insertionDestination.getDocument(), importFormatMode);
           
            // Loop through all block level nodes in the body of the section
            for (Section srcSection : docToInsert.getSections().toArray())
                for (Node srcNode : srcSection.getBody())
                {
                    // Let's skip the node if it is a last empty paragraph in a section
                    if (((srcNode.getNodeType()) == (NodeType.PARAGRAPH)))
                    {
                        Paragraph para = (Paragraph)srcNode;
                        if (para.isEndOfSection() && !para.hasChildNodes())
                            continue;
                    }

                    // This creates a clone of the node, suitable for insertion into the destination document.
                    Node newNode = importer.importNode(srcNode, true);

                    // Insert new node after the reference node.
                    dstStory.insertAfter(newNode, insertionDestination);
                    insertionDestination = newNode;
                }
        }
        else
        {
            throw new IllegalArgumentException("The destination node should be either a paragraph or table.");
        }
    }
}
