package hu.example.aspose;

import java.io.File;
import java.io.FilenameFilter;

public enum SourceFileNameFilter implements FilenameFilter
{
    INSTANCE;

    @Override
    public boolean accept(File dir, String fileName)
    {
        return fileName.endsWith("docx")||
                fileName.endsWith("docx*") ||
                fileName.endsWith("docx#") ||
                fileName.startsWith("Attachment");
    }

    public String getFieldNameWithoutMarks(String fieldName)
    {
        if(fieldName.endsWith("*") || fieldName.endsWith("#")){
            return fieldName.substring(0, fieldName.length() - 1);
        }
        return fieldName;
    }
}
