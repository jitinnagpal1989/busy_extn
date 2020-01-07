package busy_extn;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.Map;

import javax.swing.JFileChooser;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Controller {
	private static final String inputExcelSheetName = "B2B Invoices - 4A, 4B, 4C, 6B, ";
	private static int rowNumToWrite = 0; 
	private static CellStyle dateCellStyle;
	public static void main(String[] args) {
		JFileChooser chooser = new JFileChooser();
		FileNameExtensionFilter filter = new FileNameExtensionFilter(
                "XLSX & CSV files", "xlsx", "csv");
        chooser.setFileFilter(filter);
        int returnVal = chooser.showOpenDialog(null);
        if(returnVal == JFileChooser.CANCEL_OPTION) {
            return;
        }else if(returnVal != JFileChooser.APPROVE_OPTION){
        	throw new RuntimeException("Invalid file format");
        }
        
        Map<String, InvoiceData> map = readFromExcel(chooser.getSelectedFile().getAbsolutePath());
        if (map == null){
        	System.out.println("Null/blank file encountered, exiting.");
        }
        map.remove("");
        for(String key : map.keySet()){
        	InvoiceData invoiceData = map.get(key);
        	print(invoiceData);
        }
        
        writeToExcel(map);
	}

	private static void writeToExcel(Map<String, InvoiceData> map) {
		XSSFWorkbook excelBook = new XSSFWorkbook();
		dateCellStyle = excelBook.createCellStyle();
		CreationHelper createHelper = excelBook.getCreationHelper();
		dateCellStyle.setDataFormat(
		    createHelper.createDataFormat().getFormat("dd-mm-yyyy"));

		XSSFSheet excelSheet = excelBook.createSheet();
		for (String key : map.keySet()){
			writeRow(excelSheet, map.get(key));
		}
		try {
			excelBook.write(new FileOutputStream("./output.xlsx"));
		} catch(FileNotFoundException fnfe){
			System.out.println("File not found, exiting." + fnfe.getStackTrace());
		}catch(IOException ioe){
			System.out.println("An exception occurred while writing the file, exiting." + ioe.getStackTrace());
		}
	}

	private static void writeRow(XSSFSheet excelSheet, InvoiceData invoiceData) {
		for (TaxData taxData : invoiceData.taxes){
			Row row = excelSheet.createRow(rowNumToWrite++);
			int cellNum = 0;
			addCell(cellNum++, invoiceData.gstNumber, row, "String");
			addCell(cellNum++, invoiceData.receiverName, row, "String");
			addCell(cellNum++, invoiceData.invoiceNumber, row, "String");
			addCell(cellNum++, invoiceData.invoiceDate, row, "Date");
			addCell(cellNum++, taxData.invoiceValue, row, "Double");
			addCell(cellNum++, invoiceData.placeOfSupply, row, "String");
			addCell(cellNum++, invoiceData.reverseCharge, row, "String");
			addCell(cellNum++, "", row, "String");
			addCell(cellNum++, invoiceData.invoiceType, row, "String");
			addCell(cellNum++, "", row, "String");
			addCell(cellNum++, taxData.rate, row, "Double");
			addCell(cellNum++, taxData.taxableVal, row, "Double");
			addCell(cellNum++, taxData.totalTax, row, "Double");
		}
	}

	private static void addCell(int cellNum, Object data, Row row, String type) {
		Cell cell = row.createCell(cellNum);
		if (type.equalsIgnoreCase("String"))
			cell.setCellValue((String)data);
		else if (type.equalsIgnoreCase("Double"))
			cell.setCellValue((Double)data);
		else if (type.equalsIgnoreCase("Date")){
		    cell.setCellStyle(dateCellStyle);			
			cell.setCellValue((Date)data);
		}
	}

	private static void print(InvoiceData invoiceData) {
		StringBuilder data = new StringBuilder();
		data.append("invoice_no:" + invoiceData.invoiceNumber);
		data.append(" gst_no:" + invoiceData.gstNumber);
		data.append(" receiver:" + invoiceData.receiverName);
		data.append(" invoice_date:" + invoiceData.invoiceDate);
		data.append(" place:" + invoiceData.placeOfSupply);
		data.append(" reverse_charge:" + invoiceData.reverseCharge);
		data.append(" invoice_type:" + invoiceData.invoiceType);	
		for(TaxData taxData : invoiceData.taxes){
			data.append('{');
			data.append(" invoice_val:" + taxData.invoiceValue);
			data.append(" rate:" + taxData.rate);
			data.append(" taxableVal:" + taxData.taxableVal);
			data.append(" totalTax:" + taxData.totalTax);
			data.append('}');
		}
		
		System.out.println(data);
	}

	/*
	 * Reads invoice related data from excel.*/
	private static Map<String, InvoiceData> readFromExcel(String fileName){
		System.out.println(fileName);
		XSSFWorkbook excelBook = null;
		try{
			excelBook = new XSSFWorkbook(new FileInputStream(fileName));
		}catch(FileNotFoundException fnfe){
			System.out.println("File not found, exiting." + fnfe.getStackTrace());
			return null;
		}catch(IOException ioe){
			System.out.println("An exception occurred while reading the file, exiting." + ioe.getStackTrace());
			return null;
		}
		Map<String, InvoiceData> map = new LinkedHashMap();
		XSSFSheet excelSheet = excelBook.getSheetAt(0);
		Iterator<Row> iterator = excelSheet.iterator();
		int count = -1;
		while (iterator.hasNext()) {
            Row currentRow = iterator.next();
            count++;
            if(count < 5)
            	continue;
            try{
            	readRowData(currentRow, map, iterator, excelSheet);
            }catch(NullPointerException npe){
            	System.out.println("Reached EOF");
            	break;
            }   
		}
		return map;
	}

	private static void readRowData(Row row, Map<String, InvoiceData> map, Iterator<Row> iterator, XSSFSheet excelSheet) {
		InvoiceData invoiceData = new InvoiceData();
		invoiceData.gstNumber = row.getCell(1).getStringCellValue();
		invoiceData.receiverName = row.getCell(2).getStringCellValue();
		invoiceData.invoiceNumber = row.getCell(3).getStringCellValue();
		invoiceData.invoiceDate = row.getCell(4).getDateCellValue();
		invoiceData.placeOfSupply = row.getCell(6).getStringCellValue();
		invoiceData.reverseCharge = row.getCell(7).getStringCellValue();
		invoiceData.invoiceType = row.getCell(9).getStringCellValue();
		
		invoiceData.taxes = new ArrayList<TaxData>();
		TaxData taxData = readTaxData(row);
		
		invoiceData.taxes.add(taxData);
		// read 5% data
		if(iterator.hasNext()){
			row = excelSheet.getRow(row.getRowNum() + 1);
			try{
				if(row.getCell(2).getNumericCellValue() == 0){
					TaxData taxData2 = readTaxData(row);
					taxData2.invoiceValue = taxData2.taxableVal + taxData2.totalTax;
					taxData.invoiceValue -= taxData2.invoiceValue;
					invoiceData.taxes.add(taxData2);
					iterator.next();
				}
			}catch(Exception e){
				// not a 12% tax row, do nothing.
			}
		}
		map.put(invoiceData.invoiceNumber, invoiceData);
//		System.out.println(invoiceData.invoiceNumber + ", ");
	}

	private static TaxData readTaxData(Row row) {
		TaxData taxData = new TaxData();
		
		
		taxData.invoiceValue = row.getCell(5).getNumericCellValue();
		taxData.rate = row.getCell(11).getNumericCellValue()*100;
		taxData.taxableVal = row.getCell(12).getNumericCellValue();
		taxData.totalTax = row.getCell(17).getNumericCellValue();
		return taxData;
	}

}
