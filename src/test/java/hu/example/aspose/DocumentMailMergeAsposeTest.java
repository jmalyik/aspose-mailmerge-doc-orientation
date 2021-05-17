package hu.example.aspose;

import org.junit.Assert;
import org.junit.Test;

public class DocumentMailMergeAsposeTest
{    
    /**
     * the code throws here (but this is not desired)
     * org.apache.xmlbeans.impl.values.XmlValueDisconnectedException
     * at org.apache.xmlbeans.impl.values.XmlObjectBase.check_orphaned(XmlObjectBase.java:1258)
     * because the template contains w:fldSimple instead of w:fldChar and w:instrText
     */
    @Test
    public void testOrientation(){
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
}
