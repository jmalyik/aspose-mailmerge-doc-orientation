package hu.example.aspose;

import org.junit.Assert;
import org.junit.Test;

public class DocumentMailMergeAsposeTest
{    
    /**
     * here we expects that the subject.docx will be merged with ImportFormatMode.USE_DESTINATION_STYLES
     * (Using destination formatting) -> the expected font should be the default Arial
     * 
     * This test can be used to test whether the first attachment mergefield (Attachment[0]) orientation is kept or not
     */
    @Test
    public void testSubjectHasStar(){
        try
        {
            System.setProperty(DocumentMailMergeAspose.JVM_ARG_CONFIGFILE, "src/test/resources/testConfig.properties");
            DocumentMailMergeAspose merger = new DocumentMailMergeAspose();
            merger.merge();
        } catch (Exception e)
        {
            e.printStackTrace();
            Assert.fail(e.getMessage());
        }
    }
    
    @Test
    public void testSubjectPlain(){
        try
        {
            System.setProperty(DocumentMailMergeAspose.JVM_ARG_CONFIGFILE, "src/test/resources/testConfig_plain.properties");
            DocumentMailMergeAspose merger = new DocumentMailMergeAspose();
            merger.merge();
        } catch (Exception e)
        {
            e.printStackTrace();
            Assert.fail(e.getMessage());
        }
    }
    
    @Test
    public void testSubjectHashmarked(){
        try
        {
            System.setProperty(DocumentMailMergeAspose.JVM_ARG_CONFIGFILE, "src/test/resources/testConfig_hashmarked.properties");
            DocumentMailMergeAspose merger = new DocumentMailMergeAspose();
            merger.merge();
        } catch (Exception e)
        {
            e.printStackTrace();
            Assert.fail(e.getMessage());
        }
    }
}
