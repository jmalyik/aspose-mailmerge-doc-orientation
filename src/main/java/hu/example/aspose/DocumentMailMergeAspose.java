package hu.example.aspose;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.Reader;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Enumeration;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import java.util.Set;
import java.util.TreeMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.log4j.Logger;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.FieldMergingArgs;
import com.aspose.words.IFieldMergingCallback;
import com.aspose.words.IMailMergeDataSource;
import com.aspose.words.ImageFieldMergingArgs;
import com.aspose.words.ImportFormatMode;
import com.aspose.words.License;
import com.aspose.words.MailMergeCleanupOptions;
import com.aspose.words.NodeCollection;
import com.aspose.words.NodeType;
import com.aspose.words.ref.Ref;

public class DocumentMailMergeAspose
{
    private static final Logger LOGGER = Logger.getLogger(DocumentMailMergeAspose.class);
    
    public static final String JVM_ARG_CONFIGFILE = "mergeConfig";
    public static final String JVM_ARG_ASPOSELICENSEFILE = "asposeLicence";

    public static class MergeContext {
        private final Map<String, Object> variables;
        private Map<String, String> files;
        // default merge options below
        private boolean removeComments = true;
        private boolean mergeDuplicatedRegions = false;
        private boolean updateFields = true;
        private boolean acceptAllRevisions = true;
        private boolean removeUnusedFields = true;

        /**
         * I had to create factory methods because of the type erasure in constructors
         * @param variables
         */
        private MergeContext(Map<String, Object> variables)
        {
            this.variables = variables;
        }
        
        public static MergeContext createMergeContextWithFiles(Map<String, Object> variables, Map<String, String> files)
        {
            MergeContext context = new MergeContext(variables);
            context.files = files;
            return context;
        }
    }
    
    private static String getConfigFileName()
    {
        String configFileName = System.getProperty(JVM_ARG_CONFIGFILE);
        if (configFileName != null)
        {
            File f = new File(configFileName);
            if (f.exists() && f.isFile())
            {
                LOGGER.info("Using config " + configFileName);
                return configFileName;
            } else
            {
                LOGGER.info("JVM ARG " + JVM_ARG_CONFIGFILE + " is defined, but the '" + configFileName + "' file does not exist or not a file!");
            }
        }
        LOGGER.info("Using default config config.props");
        return "config.props";
    }
    
    public DocumentMailMergeAspose() throws Exception
    {
        String licenseFileName = System.getProperty(JVM_ARG_ASPOSELICENSEFILE);
        if(licenseFileName != null){
            File licenseFile = new File(licenseFileName);
            if(licenseFile.exists() && licenseFile.isFile()){
                try(InputStream stream = new FileInputStream(licenseFile)){
                    License lic = new License();
                    lic.setLicense(stream);     
                }   
            }else{
                throw new IOException("License file " + licenseFileName + " does not exists or it is not file!");
            }
        }else{
            try(InputStream stream = DocumentMailMergeAspose.class.getClassLoader().getResourceAsStream("Aspose.Words.Java.lic")){
                License lic = new License();
                lic.setLicense(stream);     
            }
        }
    };


    public static class DocumentMergingCallback implements IFieldMergingCallback{

        private final MergeContext mergeContext;
        
        public DocumentMergingCallback(MergeContext mergeContext){
            this.mergeContext = mergeContext;
        }
        
        @Override
        public void fieldMerging(FieldMergingArgs fieldMergingArgs) throws Exception
        {
            String fieldName = fieldMergingArgs.getFieldName();
            
            if (fieldName.endsWith("docx")||
                fieldName.endsWith("docx*") ||
                fieldName.endsWith("docx#") ||
                fieldName.startsWith("Attachment")) {
                
                DocumentBuilder builder = new DocumentBuilder(fieldMergingArgs.getDocument());
                builder.moveToMergeField(fieldName);
                String fileName = fieldName;
                String importFormatString = "Keeping source formatting";
                int importFormatMode = ImportFormatMode.KEEP_SOURCE_FORMATTING;
                if(fieldName.endsWith("*")){
                    fileName = fileName.substring(0, fileName.length() - 1);
                    importFormatMode = ImportFormatMode.USE_DESTINATION_STYLES;
                    importFormatString = "Using destination formatting";
                }
                if(fieldName.endsWith("#")){
                    fileName = fileName.substring(0, fileName.length() - 1);
                    importFormatMode = ImportFormatMode.KEEP_DIFFERENT_STYLES;
                    importFormatString = "Keeping different styles";
                }
                String realFile = mergeContext.files.get(fileName);
                if(realFile != null){
                    File file = new File(realFile);
                    if(file.exists() && file.isFile()){
                        LOGGER.debug("Using " + realFile + " for " + fieldName + ".\t" + importFormatString);
                        Document documentToMerge = new Document(realFile);  
                        builder.insertDocument(documentToMerge, importFormatMode);
                    }else{
                        LOGGER.warn("Attachment " + fieldName + " in config points to " + realFile + " that is either does not exist or not a file!");
                    }
                }else{
                    LOGGER.warn("Attachment " + fieldName + " missing from configuration!");
                } 
            }else{
                LOGGER.trace("FieldMerging ignored for " + fieldName + " (field is not mapped to an attachment)");
            }
        }

        @Override
        public void imageFieldMerging(ImageFieldMergingArgs imageFieldMergingArgs) throws Exception
        {
        }
    }
    
    public void merge() throws Exception{
        String configFileName = getConfigFileName();
        Properties props = new Properties();
        try (Reader reader = new InputStreamReader(new FileInputStream(configFileName), StandardCharsets.UTF_8))
        {
            props.load(reader);
        }
        String template = props.getProperty("mainTemplate");
        String targetDocx = props.getProperty("outDocxPath");
        String targetPDF = props.getProperty("outPdfPath");
        Map<String, String> files = new TreeMap<>();
        Map<String, Object> variables = new TreeMap<>();// treeMap because it is sorted
        Enumeration<Object> keys = props.keys();
        while (keys.hasMoreElements())
        {
            String key = (String) keys.nextElement();
            if(key.startsWith("textVariables")){
                variables.put(key.substring("textVariables.".length()), props.getProperty(key));
            }else if(key.startsWith("attachments")){
                files.put(key.substring("attachments.".length()), props.getProperty(key));
            }
        }
        
        MergeContext mergeContext = MergeContext.createMergeContextWithFiles(variables, files);
        
        if(targetDocx != null){
            try (FileOutputStream fos = new FileOutputStream(targetDocx))
            {
                LOGGER.debug("Writing " + targetDocx);
                byte[] content = merge(template, DocumentFormat.DOCX, mergeContext);
                fos.write(content);
                fos.flush();
                LOGGER.info("Done writing " + targetDocx);
            }
        }
        if(targetPDF != null){
            try (FileOutputStream fos = new FileOutputStream(targetPDF))
            {
                LOGGER.debug("Writing " + targetPDF);
                byte[] content = merge(template, DocumentFormat.PDF, mergeContext);
                fos.write(content);
                fos.flush();
                LOGGER.info("Done writing " + targetPDF);
            }
        }
    }
    
    public byte[] merge(String template, DocumentFormat targetFormat, MergeContext mergeContext) throws Exception{
        LOGGER.info("Using template " + template);
        File templateFile = new File(template);
        try(FileInputStream fis = new FileInputStream(templateFile)){
            return mergeInternal(fis, targetFormat, mergeContext);
        }
    }

    public byte[] mergeInternal(InputStream templateStream, DocumentFormat targetFormat, MergeContext mergeContext) throws Exception{
        
        if(LOGGER.isDebugEnabled()){
            LOGGER.debug("");
            LOGGER.debug("variables: ");
            mergeContext.variables.keySet().forEach(key -> LOGGER.debug(key + "=" + mergeContext.variables.get(key)));
            LOGGER.debug("");
            LOGGER.debug("attachments: ");
            mergeContext.files.keySet().forEach(key -> LOGGER.debug(key + "=" + mergeContext.files.get(key)));
            LOGGER.debug("");
        }
        
        Document outputDoc = new Document(templateStream);

        // this row below is needed to be able to remove the template (first) row that will be empty after the merge 
        outputDoc.getMailMerge().setCleanupOptions(MailMergeCleanupOptions.REMOVE_EMPTY_TABLE_ROWS);
        if(mergeContext.removeUnusedFields ){
            LOGGER.debug("Unused fields will be removed.");
            int currentCleanupOptions = outputDoc.getMailMerge().getCleanupOptions();
            outputDoc.getMailMerge().setCleanupOptions(currentCleanupOptions
                    | MailMergeCleanupOptions.REMOVE_UNUSED_FIELDS);
        }

        /**
         * updating settings before the document will be rendered 
         */
        if(mergeContext.mergeDuplicatedRegions){
            LOGGER.debug("Duplicated regions will be merged.");
        }
        outputDoc.getMailMerge().setMergeDuplicateRegions(mergeContext.mergeDuplicatedRegions);
        
        DocumentMergingCallback fieldMergingCallback = new DocumentMergingCallback(mergeContext);
        outputDoc.getMailMerge().setFieldMergingCallback(fieldMergingCallback);
        List<IMailMergeDataSource> tableDataSources = createMailMergeDataSources(mergeContext);
        
        for(IMailMergeDataSource ds : tableDataSources){
            outputDoc.getMailMerge().executeWithRegions(ds);
        }
        outputDoc.getMailMerge().execute(
                mergeContext.variables.keySet().toArray(new String[mergeContext.variables.size()]),
                mergeContext.variables.values().toArray(new Object[mergeContext.variables.size()])
                );
        
        if(mergeContext.removeComments){
            LOGGER.debug("Comments will be removed.");
            NodeCollection<?> comments = outputDoc.getChildNodes(NodeType.COMMENT, true);
            comments.clear();
        }
        
        if(mergeContext.acceptAllRevisions){
            LOGGER.debug("All revisions will be accepted.");
            outputDoc.acceptAllRevisions();
        }
        
        if(mergeContext.updateFields){
            LOGGER.debug("Updating fields.");
            outputDoc.updateFields();
        }
        
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        if (DocumentFormat.DOCX.equals(targetFormat))
        {
            outputDoc.save(bos,new com.aspose.words.OoxmlSaveOptions());
        } else
        {
            outputDoc.save(bos,new com.aspose.words.PdfSaveOptions());
        }
        return bos.toByteArray();        
    }

    private final Pattern arrayVarablePatternForDefaultDS = Pattern.compile("^(\\w+)\\.(\\d+)$");
    private final Pattern arrayVarablePatternForNamedDS = Pattern.compile("^(\\w+)\\.(\\w+)\\.(\\d+)$");

    /**
     * Via this method, we support multiple tables in the same document
     * @param mergeContext
     * @return
     */
    private List<IMailMergeDataSource> createMailMergeDataSources(MergeContext mergeContext)
    {
        Map<String, Object> variables = mergeContext.variables;
        Set<String> tableVariableNames = new HashSet<>();
        Map<String, Set<String>> namedTableVariableNames = new HashMap<>();
        Map<String,Map<Integer, Object>> tableData = new HashMap<>();
        Map<String, Map<String,Map<Integer, Object>>> namedTableData = new HashMap<>();
        for(String key : variables.keySet()){
            Matcher m = arrayVarablePatternForDefaultDS.matcher(key);
            if(m.matches()){
                String variableName = m.group(1);
                tableVariableNames.add(variableName);
                int index = Integer.parseInt(m.group(2));
                tableData.putIfAbsent(variableName, new HashMap<>());
                tableData.get(variableName).put(Integer.valueOf(index), variables.get(key));
            }
            Matcher m2 = arrayVarablePatternForNamedDS.matcher(key);
            if(m2.matches()){
                String tableName = m2.group(1);
                String variableName = m2.group(2);
                namedTableVariableNames.putIfAbsent(tableName, new HashSet<>());
                namedTableVariableNames.get(tableName).add(variableName);
                int index = Integer.parseInt(m2.group(3));
                namedTableData.putIfAbsent(tableName, new HashMap<>());
                namedTableData.get(tableName).putIfAbsent(variableName, new HashMap<>());
                namedTableData.get(tableName).get(variableName).put(Integer.valueOf(index), variables.get(key));
            }
        }

        List<IMailMergeDataSource> dataSources = new ArrayList<IMailMergeDataSource>();
        
        if(!tableVariableNames.isEmpty()){
            IMailMergeDataSource defaultDS = createMailMergeDS("Default", tableVariableNames, tableData);
            dataSources.add(defaultDS);
        }
        
        namedTableData.keySet().forEach(tableName -> {
            IMailMergeDataSource ds = createMailMergeDS(tableName, 
                    namedTableVariableNames.get(tableName), 
                    namedTableData.get(tableName));
            dataSources.add(ds);
        });
        
        return dataSources;
    }

    /**
     * Do not forget!:
     * In order to get it work, you have to use "aspose row notation":
     * In the first cell of the row, you have to place a mergefield using this name: TableStart:<tableName>
     * and, in the same row, in the last cell TableEnd:<tableName>
     * (This was the simplest example. Basically, you have to put these start and end mergefield to mark the rows/cells 
     * that you would like to copy for every record in this datasource)
     * 
     * @see {@link https://github.com/aspose-words/Aspose.Words-for-Java/blob/master/Examples/src/main/resources/MailMerge/NestedMailMerge.CustomDataSource.doc}
     * 
     * @param tableName
     * @param tableVariableNames
     * @param tableData
     * @return
     */
    private IMailMergeDataSource createMailMergeDS(String tableName, Set<String> tableVariableNames, Map<String, Map<Integer, Object>> tableData)
    {
        if(LOGGER.isDebugEnabled()){
            printTableData(tableName, tableData);
        }   
        // It is possible to not to provide values to some cells in the table so the tableData.size()
        // is not the right value. 
        // Therefore we have to find the biggest index within this data structure. 
        // the value will be used in hasMoreRecords
        int size = tableData.values().stream() // the values are Map<Integer, Object>
                .map(map -> Collections.max(map.keySet())) // map it to the stream of the maximum of the keys
                .mapToInt(Integer::intValue) // map Integer to int
                .max() // grab the maximum
                .orElse(0); // if the stream is empty, return zero
        
        return new IMailMergeDataSource()
        {
            // When the data source is initialized, it must be positioned before the first record.
            private int rowIndex = -1;
            
            @Override
            public boolean moveNext() throws Exception
            {
                boolean flag = hasMoreRecords();
                if (flag){ 
                    rowIndex++;
                }
                return flag;
            }
            
            private boolean hasMoreRecords() {
                return rowIndex < size;
            }
            
            @Override
            public boolean getValue(String fieldName, Ref<Object> fieldValue) throws Exception
            {
                if (tableVariableNames.contains(fieldName)) {
                    Map<Integer, Object> variableValues = tableData.get(fieldName);
                    if(variableValues.containsKey(rowIndex)){
                        fieldValue.set(variableValues.get(rowIndex));
                        return true;    
                    }
                }
                // A field with this name was not found,
                // return false to the Aspose.Words mail merge engine
                fieldValue.set(null);
                return false;
            }
            
            @Override
            public String getTableName() throws Exception
            {
                return tableName;  
            }
            
            @Override
            public IMailMergeDataSource getChildDataSource(String var1) throws Exception
            {
                return null;
            }
        };
    }

    private void printTableData(String tableName, Map<String, Map<Integer, Object>> tableData)
    {
        LOGGER.debug("The parsed '" + tableName + "' table data from config:");
        for(String key : tableData.keySet()){
            LOGGER.debug(key + ": " + tableData.get(key));
        }
        LOGGER.debug("");
    }
}
