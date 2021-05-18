package hu.example.aspose;

import java.io.InputStream;

import org.junit.Test;

import com.aspose.words.BreakType;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.FieldMergingArgs;
import com.aspose.words.IFieldMergingCallback;
import com.aspose.words.ImageFieldMergingArgs;
import com.aspose.words.ImportFormatMode;
import com.aspose.words.License;
import com.aspose.words.Orientation;
import com.aspose.words.PageSetup;
import com.aspose.words.Section;


public class SuggestedMergeTest
{
    private static final String ASPOSE_WORDS_JAVA_LIC_FILE = "Aspose.Words.Java.lic";

    static
    {
        ClassLoader cl = DocumentMailMergeAspose.class.getClassLoader();
        try (InputStream stream = cl.getResourceAsStream(ASPOSE_WORDS_JAVA_LIC_FILE))
        {
            License lic = new License();
            lic.setLicense(stream);
        } catch (Exception e)
        {
            e.printStackTrace();
        }
    }

    private class HandleMergeField implements IFieldMergingCallback
    {

        private final int importFormatMode;

        public HandleMergeField(int importFormatMode)
        {
            this.importFormatMode = importFormatMode;
        }

        @Override
        public void fieldMerging(FieldMergingArgs args) throws Exception
        {
            if (args.getFieldName().equals("Attachment[0]#") || args.getFieldName().equals("Attachment[1]#"))
            {
                DocumentBuilder builder = new DocumentBuilder(args.getDocument());
                System.out.println("Template current section width: " + builder.getCurrentSection().getPageSetup().getPageWidth());
                builder.moveToMergeField(args.getFieldName());
                Document document = (Document) args.getFieldValue();
                System.out.println(args.getFieldName() + " orientation: " + Orientation.getName(document.getFirstSection().getPageSetup().getOrientation()));
                System.out.println(args.getFieldName() + " width: " + document.getFirstSection().getPageSetup().getPageWidth());
                builder.insertDocument(document, importFormatMode);
            }
        }

        @Override
        public void imageFieldMerging(ImageFieldMergingArgs args)
        {

        }
    }

    @Test
    public void testMergeKeepDiff() throws Exception
    {
        System.out.println("KEEP_DIFFERENT_STYLES");
        Document doc = new Document("src/test/resources/template.docx");
        Document subDoc = new Document("src/test/resources/attachment1.docx");
        Document subDoc2 = new Document("src/test/resources/attachment2.docx");

        doc.getMailMerge().setFieldMergingCallback(new HandleMergeField(ImportFormatMode.KEEP_DIFFERENT_STYLES));
        doc.getMailMerge().execute(new String[] { "Attachment[0]#", "Attachment[1]#" }, new Object[] { subDoc, subDoc2 });

        doc.save("target/awjava-keep-diff-styles.docx");

    }

    @Test
    public void testMergeUseDest() throws Exception
    {
        System.out.println("USE_DESTINATION_STYLES");
        Document doc = new Document("src/test/resources/template.docx");
        Document subDoc = new Document("src/test/resources/attachment1.docx");
        Document subDoc2 = new Document("src/test/resources/attachment2.docx");

        doc.getMailMerge().setFieldMergingCallback(new HandleMergeField(ImportFormatMode.USE_DESTINATION_STYLES));
        doc.getMailMerge().execute(new String[] { "Attachment[0]#", "Attachment[1]#" }, new Object[] { subDoc, subDoc2 });

        doc.save("target/awjava-use-dest-styles.docx");

    }

    @Test
    public void testMergeKeepSourceFormat() throws Exception
    {
        System.out.println("KEEP_SOURCE_FORMATTING");
        Document doc = new Document("src/test/resources/template.docx");
        Document subDoc = new Document("src/test/resources/attachment1.docx");
        Document subDoc2 = new Document("src/test/resources/attachment2.docx");

        doc.getMailMerge().setFieldMergingCallback(new HandleMergeField(ImportFormatMode.KEEP_SOURCE_FORMATTING));
        doc.getMailMerge().execute(new String[] { "Attachment[0]#", "Attachment[1]#" }, new Object[] { subDoc, subDoc2 });

        doc.save("target/awjava-keep-sourceform.docx");

    }

    @Test
    public void testMergeKeepSourceFormatAndAtt1LandscapeOriented() throws Exception
    {
        System.out.println("KEEP_SOURCE_FORMATTING + landscape Attachment1");
        Document doc = new Document("src/test/resources/template.docx");
        Document subDoc = new Document("src/test/resources/attachment1_landscape.docx");
        Document subDoc2 = new Document("src/test/resources/attachment2.docx");

        doc.getMailMerge().setFieldMergingCallback(new HandleMergeField(ImportFormatMode.KEEP_SOURCE_FORMATTING));
        doc.getMailMerge().execute(new String[] { "Attachment[0]#", "Attachment[1]#" }, new Object[] { subDoc, subDoc2 });

        doc.save("target/awjava-keep-sourceform-att1-landscaped.docx");

    }
    
    @Test
    public void testMergeKeepSourceFormatAndAtt1LandscapeOriented2() throws Exception
    {
        System.out.println("KEEP_SOURCE_FORMATTING + landscape Attachment1");
        Document doc = new Document("src/test/resources/template.docx");
        Document subDoc = new Document("src/test/resources/attachment1_landscape.docx");
        Document subDoc2 = new Document("src/test/resources/attachment2.docx");

        doc.getMailMerge().setFieldMergingCallback(new HandleMergeField2(ImportFormatMode.KEEP_SOURCE_FORMATTING));
        doc.getMailMerge().execute(new String[] { "Attachment[0]#", "Attachment[1]#" }, new Object[] { subDoc, subDoc2 });

        doc.save("target/awjava-keep-sourceform-att1-landscaped2.docx");

    }
    
    private class HandleMergeField2 implements IFieldMergingCallback
    {

        private final int importFormatMode;

        public HandleMergeField2(int importFormatMode)
        {
            this.importFormatMode = importFormatMode;
        }

        @Override
        public void fieldMerging(FieldMergingArgs args) throws Exception
        {
            if (args.getFieldName().equals("Attachment[0]#") || args.getFieldName().equals("Attachment[1]#"))
            {
                DocumentBuilder builder = new DocumentBuilder(args.getDocument());
                System.out.println("Template current section width: " + builder.getCurrentSection().getPageSetup().getPageWidth());
                builder.moveToMergeField(args.getFieldName());
                Document document = (Document) args.getFieldValue();
                System.out.println(args.getFieldName() + " orientation: " + Orientation.getName(document.getFirstSection().getPageSetup().getOrientation()));
                System.out.println(args.getFieldName() + " width: " + document.getFirstSection().getPageSetup().getPageWidth());
                // if there is a difference, then store it as we will modify the template
                boolean  hasDiff = hasDifferentOrientationOrWidth(builder.getCurrentSection(), document.getFirstSection());
                // copy page setup (width & orientation)
                copyPageSetup(builder, document);
                // insert document
                builder.insertDocument(document, importFormatMode);
                // if there was a diff, insert a section break too
                if(hasDiff){
                    System.out.println("inserting section page break");
                    builder.insertBreak(BreakType.SECTION_BREAK_NEW_PAGE); 
                }
            }
        }
        
        private void copyPageSetup(DocumentBuilder builder, Document documentToMerge)
        {
            PageSetup documentToMergePageSetup = documentToMerge.getFirstSection().getPageSetup();
            PageSetup templateCurrentSectionPageSetup = builder.getCurrentSection().getPageSetup();
            templateCurrentSectionPageSetup.setOrientation(documentToMergePageSetup.getOrientation());
            System.out.println("Set orientation: " + Orientation.getName(documentToMergePageSetup.getOrientation()));
            templateCurrentSectionPageSetup.setPageHeight(documentToMergePageSetup.getPageHeight());
            System.out.println("Set page height: " + documentToMergePageSetup.getPageHeight());
            templateCurrentSectionPageSetup.setPageWidth(documentToMergePageSetup.getPageWidth());
            System.out.println("Set page width: " + documentToMergePageSetup.getPageWidth());
        }

        private boolean hasDifferentOrientationOrWidth(Section templateSection, Section documentToMergeSection)
        {
            if (templateSection.getPageSetup().getOrientation() != documentToMergeSection.getPageSetup().getOrientation())
            {
                System.out.println("Different orientation");
                return true;
            }
            if (templateSection.getPageSetup().getPageWidth() != documentToMergeSection.getPageSetup().getPageWidth())
            {
                System.out.println("Different page width");
                return true;
            }
            return false;
        }

        @Override
        public void imageFieldMerging(ImageFieldMergingArgs args)
        {

        }
    }
}
