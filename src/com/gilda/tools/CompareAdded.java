package com.gilda.tools;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.dom4j.DocumentException;
import org.dom4j.Element;
import org.dom4j.io.SAXReader;
import org.w3c.dom.Document;
import org.w3c.dom.NamedNodeMap;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

public class CompareAdded {

    private static final String XMLFILEPREFIX = "strings";
    private static final int EXCEL_COLUMN_0 =0;
    private static final int EXCEL_COLUMN_1 =1;
    private static final int EXCEL_COLUMN_2 =2;
    private static final int EXCEL_COLUMN_3 =3;
    /**all start point method*/
    public static void main(String[] args) {
        CompareAdded manThis = new CompareAdded();
        manThis.doConvert();
        //manThis.doCombine();
        manThis.docompare();
    }
   
    private void convertToExcel(File file)throws Exception{
        File excelFile = createExcelFile(file);
        FileOutputStream outputStream = new FileOutputStream(excelFile);
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        factory.setIgnoringComments(true);//
        DocumentBuilder builder = factory.newDocumentBuilder();
        Document doc = builder.parse(file);
        NodeList listChild = doc.getElementsByTagName("string");//this is the resources node
        NodeList arrayStringNode = doc.getElementsByTagName("string-array");
        NodeList pluralsStringNode = doc.getElementsByTagName("plurals");
        //System.out.println("length : "+listChild.getLength() );
        Workbook stringBook = new HSSFWorkbook();
        Sheet stringSheet = stringBook.createSheet();
        //int startRowNumber = stringSheet.getLastRowNum();
        int arrayIndexTotal = 1;
        for (int i = 0; i < listChild.getLength(); i++){
            Node singleNode = listChild.item(i);
            NamedNodeMap nodeMap = singleNode.getAttributes();
            Node product = nodeMap.getNamedItem("product");
            if(!isTranslatable(nodeMap)){
            	continue;
            }
            String id = nodeMap.getNamedItem("name").getNodeValue();
            if( product != null){
                String productString = product.getNodeValue();
                if(productString != null){
                    id = id + "_" + productString;
                    //System.out.println(id +" : " + productString);
                }
            }
            Row newRow = stringSheet.createRow(arrayIndexTotal++);
            newRow.createCell(EXCEL_COLUMN_0).setCellValue(id);
            String value = singleNode.getTextContent();
            newRow.createCell(EXCEL_COLUMN_1).setCellValue(value);
            //System.out.println(id +" : "+value );
        }
       
        for(int i = 0;i< arrayStringNode.getLength(); i++){
            Node singleNode = arrayStringNode.item(i);
            NamedNodeMap nodeMap = singleNode.getAttributes();
            if(!isTranslatable(nodeMap)){
            	continue;
            }
            String id = nodeMap.getNamedItem("name").getNodeValue();
            NodeList arrayItemsList = singleNode.getChildNodes();
            int itemIndex = 0;
            for(int j = 0;j < arrayItemsList.getLength(); j++){
                String value = arrayItemsList.item(j).getTextContent();
                if(value.trim().equals("")) continue;
                if(arrayItemsList.item(j).getNodeName().equals("item")){
                    Row newRow = stringSheet.createRow(arrayIndexTotal++);
                    newRow.createCell(EXCEL_COLUMN_0).setCellValue(id + "_" + itemIndex++);
                    newRow.createCell(EXCEL_COLUMN_1).setCellValue(value);
                    //System.out.println(id +" : "+value );
                }
            }
        }
        for(int i = 0;i< pluralsStringNode.getLength(); i++){
            Node singleNode = pluralsStringNode.item(i);
            NamedNodeMap nodeMap = singleNode.getAttributes();
            if(!isTranslatable(nodeMap)){
            	continue;
            }
            String id = nodeMap.getNamedItem("name").getNodeValue();
            NodeList arrayItemsList = singleNode.getChildNodes();
            for(int j = 0;j < arrayItemsList.getLength(); j++){
                String value = arrayItemsList.item(j).getTextContent();
                if(value.trim().equals("")) continue;
               
                NamedNodeMap quantityAttr = arrayItemsList.item(j).getAttributes();
                String quantity = "";
                if(quantityAttr != null){
                    quantity = quantityAttr.getNamedItem("quantity").getNodeValue();
                } else {
                    System.err.println(id +" : "+value );
                }
                if(arrayItemsList.item(j).getNodeName().equals("item")){
                    Row newRow = stringSheet.createRow(arrayIndexTotal++);
                    newRow.createCell(EXCEL_COLUMN_0).setCellValue(id + "_" + quantity);
                    newRow.createCell(EXCEL_COLUMN_1).setCellValue(value);
                }
            }
        }
        stringBook.write(outputStream);
        outputStream.flush();
        outputStream.close();
    }
    
    private boolean isTranslatable(NamedNodeMap nodeMap){
    	Node transNode = nodeMap.getNamedItem("translatable");
        if( transNode != null){
            String isTranslate = transNode.getNodeValue();
            if((isTranslate != null) && isTranslate.equalsIgnoreCase("false")){
                return false;//do not add to excel
            }
        }
        return true;
    }
    
    private void compareTwoExcel(File English, File other){
        HSSFWorkbook EnglishWorkbook ;
        HSSFWorkbook OtherWorkbook ;
        try{
            EnglishWorkbook = new HSSFWorkbook(new FileInputStream(English));
            OtherWorkbook = new HSSFWorkbook(new FileInputStream(other));
        } catch (IOException ioe){
            System.err.println(ioe);
            return;
        }
        if((EnglishWorkbook != null) && (OtherWorkbook != null)){
            Map<String, String> englishMap = getExcelKeyValueMap(EnglishWorkbook);
            Map<String, String> otherMap = getExcelKeyValueMap(OtherWorkbook);
            Sheet otherSheet = OtherWorkbook.getSheetAt(0);
            otherSheet.setColumnWidth(0, 30*256);
            otherSheet.setColumnWidth(1, 30*256);
            otherSheet.setColumnWidth(2, 30*256);
            otherSheet.setColumnHidden(EXCEL_COLUMN_3, true);
            //set the column title
            Row rowtitle = otherSheet.createRow(0);
            rowtitle.createCell(EXCEL_COLUMN_0).setCellValue("string_id");
            rowtitle.createCell(EXCEL_COLUMN_1).setCellValue("English");
            String filePath = other.getParent();
            rowtitle.createCell(EXCEL_COLUMN_2).setCellValue(filePath.substring(filePath.indexOf("values")));
            rowtitle.createCell(EXCEL_COLUMN_3).setCellValue("string_id");
            Set<String> ids= englishMap.keySet();
            int index = 1;
            for (String id : ids) {
                String otherValue = otherMap.get(id);
                Row row = otherSheet.createRow(index++);
                row.createCell(EXCEL_COLUMN_0).setCellValue(id);
                row.createCell(EXCEL_COLUMN_1).setCellValue(englishMap.get(id));
                if((otherValue == null) || otherValue.isEmpty()){
                    row.createCell(EXCEL_COLUMN_2).setCellValue("");
                    row.createCell(EXCEL_COLUMN_3).setCellValue("");
                } else {
                    row.createCell(EXCEL_COLUMN_2).setCellValue(otherMap.get(id));
                    row.createCell(EXCEL_COLUMN_3).setCellValue(id);
                    otherMap.remove(id);
                }
            }
            if(otherMap.size()>0){//to add the value left in otherMap
                Set<String> leftIds = otherMap.keySet();
                for (String id : leftIds) {
                    Row row = otherSheet.createRow(index++);
                    row.createCell(EXCEL_COLUMN_2).setCellValue(otherMap.get(id));
                    row.createCell(EXCEL_COLUMN_3).setCellValue(id);
                }
            }
            try {
                FileOutputStream output = new FileOutputStream(other);
                OtherWorkbook.write(output);
                output.flush();
                output.close();
            } catch (FileNotFoundException e) {
                // TODO Auto-generated catch block
                e.printStackTrace();
            } catch (IOException e) {
                // TODO Auto-generated catch block
                e.printStackTrace();
            }
        }
    }
   
    private File createExcelFile(File xmlFile){
        String fullPath = xmlFile.getAbsolutePath();
        //System.out.println(fullPath);
        if(fullPath.endsWith("xml")){
            fullPath = fullPath.replace("xml", "xls");
        }
        File excelFile = new File(fullPath);
       
        return excelFile;
    }
   
    private void doConvert(){
        File [] files = getFiles("D:\\string_compare");
        for (int i = 0; i < files.length; i++) {
            try {
                convertToExcel(files[i]);
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }
   
    private void doCombine(){
        List<File> files = listFileByType(new File("D:\\string_compare"),
                "xls", new ArrayList<File>());
        for (File file : files) {
            File fileInValues = getExcelInValues(file);
            if(fileInValues == null || !fileInValues.exists()) continue;
            compareTwoExcel(fileInValues, file);
        }
    }
    
    private void docompare(){
    	 List<File> files = listFileByType(new File("D:\\string_compare\\star_os"),
                 "xls", new ArrayList<File>());
    	 for (File file : files) {
             File fileInValues = getOriginStringFile(file);
             if(fileInValues == null || !fileInValues.exists()) continue;
             extraAddedString(file, fileInValues);
         }
    }
    
    /**
     * @Todo get the file of English of excel
     *       for example: give excel file "D:\string_compare\Browser\values-af\strings.xls"
     *       it will return file "D:\string_compare\Browser\values\strings.xls"
     * @param excel
     * @return
     */
    private File getExcelInValues(File excel){
        final String value = "values";
        String fullPath = excel.getAbsolutePath();
        int indexOfValue = fullPath.indexOf(value);
        int nextSlash = fullPath.indexOf("\\", indexOfValue);
        String subString = fullPath.substring(indexOfValue, nextSlash);
        //System.out.println(indexOfValue +" : " + subString );
        if(value.equals(subString)){
            return null;
        } else {
            String valuesString = fullPath.replace(subString, value);
            System.out.println(valuesString);
            return new File(valuesString);
        }
    }
    
    /**
     * @Todo get the file of English of excel
     *       for example: give excel file "D:\string_compare\Browser\values-af\strings.xls"
     *       it will return file "D:\string_compare\Browser\values\strings.xls"
     * @param excel
     * @return
     */
    private File getOriginStringFile(File excel){
        final String origin = "origin";
        final String star_os = "star_os";
        String fullPath = excel.getAbsolutePath();
        String target_path = fullPath.replace(star_os, origin);
        //System.out.println("getOriginStringFile:" + target_path);
        return new File(target_path);
    }
   
    private File[] getFiles(String fileRoot){
        File parent = new File(fileRoot);
        List<File> listFile = listFileByType(parent, "xml", new ArrayList<File>());
        return listFile.toArray(new File[listFile.size()]);
    }
   
    private List<File> listFileByType(File file, String suffix, List<File> list){
        File [] xmlInDirectory;
        if(file.isDirectory()){
            xmlInDirectory = file.listFiles();
            for (File file2 : xmlInDirectory) {
                listFileByType(file2, suffix, list);
            }
        } else {
            String nameString = file.getName();
            if(nameString.contains(XMLFILEPREFIX)&&nameString.endsWith(suffix)){
                list.add(file);
            } else if(nameString.contains("array")&&nameString.endsWith(suffix)){
                list.add(file);
            }
        }
        return list;
    }
   
    /**
     * @Todo the other file will fill string of english in column 3 and 4
     * @param English
     * @param other
     */
    private void extraAddedString(File English, File other){
        HSSFWorkbook EnglishWorkbook ;
        HSSFWorkbook OtherWorkbook ;
        try{
            EnglishWorkbook = new HSSFWorkbook(new FileInputStream(English));
            OtherWorkbook = new HSSFWorkbook(new FileInputStream(other));
        } catch (IOException ioe){
            System.err.println(ioe);
            return;
        }
        Map<String, String> englishMap = getExcelKeyValueMap(EnglishWorkbook);
        Map<String, String> otherMap = getExcelKeyValueMap(OtherWorkbook);

        Set<String> ids= englishMap.keySet();
        Set<String> other_ids = otherMap.keySet();
        //int index = 1;
		for (String id : other_ids) {
			ids.remove(id);
			//System.out.println(id);
		}
    	HSSFSheet sheet = EnglishWorkbook.createSheet();
    	EnglishWorkbook.removeSheetAt(0);
        if(ids.size()>0){//to add the value left in otherMap
        	int i=0;
        	sheet.setColumnWidth(0, 30*256);
        	sheet.setColumnWidth(1, 40*256);
        	Row row_first = sheet.createRow(i++);
        	row_first.createCell(EXCEL_COLUMN_0).setCellValue("id");
        	row_first.createCell(EXCEL_COLUMN_1).setCellValue("value-en");
            for (String id : ids) {
            	//System.out.println("write:"+id);
                Row row = sheet.createRow(i++);
                row.createCell(EXCEL_COLUMN_0).setCellValue(id);
                row.createCell(EXCEL_COLUMN_1).setCellValue(englishMap.get(id));
            }
        } else {
        	English.delete();
        	return;
        }
        try {
            FileOutputStream output = new FileOutputStream(English);
            EnglishWorkbook.write(output);
            output.flush();
            output.close();
        } catch (FileNotFoundException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }
   
    private Map<String, String> getExcelKeyValueMap(HSSFWorkbook workbook){
        Map<String, String> key_value = new HashMap<String, String>();
        HSSFSheet sheet = workbook.getSheetAt(0);
        for (Row row : sheet) {
            String ID = row.getCell(EXCEL_COLUMN_0).getStringCellValue();
            String value = row.getCell(EXCEL_COLUMN_1).getStringCellValue();
            key_value.put(ID, value);
        }
        return key_value;
    }
}