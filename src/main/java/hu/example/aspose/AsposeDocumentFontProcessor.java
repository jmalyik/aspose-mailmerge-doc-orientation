package hu.example.aspose;

import java.util.Iterator;

import org.apache.log4j.Logger;

import com.aspose.words.Document;
import com.aspose.words.Font;
import com.aspose.words.NodeCollection;
import com.aspose.words.NodeType;
import com.aspose.words.Paragraph;
import com.aspose.words.Run;

/**
 * when we would like to use the destination styles, but we have a 'Arial Narrow' font, that will be kept :(
 * so we need to replace it to Arial Narrow in order to get rid of it in the resulted document
 */
public class AsposeDocumentFontProcessor
{
    private static final Logger LOGGER = Logger.getLogger(AsposeDocumentFontProcessor.class);
    
    public static void removeUnknownFonts(Document documentToMerge)
    {
        if(documentToMerge == null){
            return;
        }
        
        NodeCollection<?> nodes = documentToMerge.getChildNodes(NodeType.PARAGRAPH, true);
        Iterator<?> it = nodes.iterator();
        while (it.hasNext())
        {
            Paragraph paragraph = (Paragraph) it.next();
            for(Run run : paragraph.getRuns()){
                Font font = run.getFont();
                String fontName = font.getName();
                if(fontName.startsWith("'") && fontName.endsWith("'")){
                    String fixedFontName = fontName.substring(1, fontName.length() - 1);
                    LOGGER.debug("Replacing " + fontName + " to  "+ fixedFontName);
                    font.setName(fixedFontName);
                    font.setNameAscii(fixedFontName);
                    font.setNameBi(fixedFontName);
                    font.setNameFarEast(fixedFontName);
                    font.setNameOther(fixedFontName);
                }
            }
        }
    }
}
