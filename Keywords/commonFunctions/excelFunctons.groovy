package commonFunctions

import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import static com.kms.katalon.core.testobject.ObjectRepository.findWindowsObject

import com.kms.katalon.core.annotation.Keyword
import com.kms.katalon.core.checkpoint.Checkpoint
import com.kms.katalon.core.cucumber.keyword.CucumberBuiltinKeywords as CucumberKW
import com.kms.katalon.core.mobile.keyword.MobileBuiltInKeywords as Mobile
import com.kms.katalon.core.model.FailureHandling
import com.kms.katalon.core.testcase.TestCase
import com.kms.katalon.core.testdata.TestData
import com.kms.katalon.core.testng.keyword.TestNGBuiltinKeywords as TestNGKW
import com.kms.katalon.core.testobject.TestObject
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import com.kms.katalon.core.windows.keyword.WindowsBuiltinKeywords as Windows

import internal.GlobalVariable

//excel includes
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;



public class excelFunctons {
	@Keyword
	def readExcel (String filePath, String sheetName, int column, int row) {
		'open file'
		FileInputStream file = new FileInputStream (new File(filePath));
		HSSFWorkbook workbook = new HSSFWorkbook(file);
		//XSSFWorkbook workbook = new XSSFWorkbook(file);

		'select sheetName'
		//XSSFSheet sheet = workbook.getSheetAt(sheetIndex); to select sheet using sheet index
		HSSFSheet sheet = workbook.getSheet(sheetName)
		//XSSFSheet sheet = workbook.getSheet(sheetName)

		'Read data from excel'
		def Data_fromCell;
		HSSFRow rowExcel;
		HSSFCell cell;
		//XSSFRow rowExcel;
		//XSSFCell cell;

		rowExcel=sheet.getRow(row); //returns the logical row
		cell=rowExcel.getCell(column,org.apache.poi.ss.usermodel.Row.MissingCellPolicy.CREATE_NULL_AS_BLANK); //getting the cell representing the given column

		Data_fromCell = cell.toString()

		//String Data_fromCell;
		//Data_fromCell=sheet.getRow(row).getCell(column).getStringCellValue();
		//Data_fromCell=sheet.getRow(row).getCell(column).getRawValue()
		//println("data from cell : "+Data_fromCell)
		return Data_fromCell;

		file.close();
	}

	@Keyword
	def writeExcel (String filePath, String sheetName, int column, int row, String value) {
		'open file'
		FileInputStream file = new FileInputStream (new File(filePath))
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		'select sheetName'
		XSSFSheet sheet = workbook.getSheet(sheetName)

		'Write data to excel'
		sheet.getRow(row).createCell(column).setCellValue(value);

		file.close();
		FileOutputStream outFile =new FileOutputStream(new File(filePath));
		workbook.write(outFile);
		outFile.close();
	}

	@Keyword
	def getSheetLastRow (String filePath, String sheetName) {
		'open file'
		FileInputStream file = new FileInputStream (new File(filePath))
		HSSFWorkbook workbook = new HSSFWorkbook(file);
		//XSSFWorkbook workbook = new XSSFWorkbook(file);
		'select sheetName'
		HSSFSheet sheet = workbook.getSheet(sheetName)
		//XSSFSheet sheet = workbook.getSheet(sheetName)

		'get last row of selected sheet'
		return sheet.getLastRowNum()

		file.close();

	}

	@Keyword
	def getAllDetail (String filePath, String sheetName, int PKIndex, int SKColumnIndex, int targetColumnIndex) {
		//String temp = getSheetLastRow(filePath, sheetName)
		//int jmlRow = Integer.parseInt(temp)
		int jmlRow = getSheetLastRow(filePath, sheetName)
		String returnValue = ""
		def returnList = []
		int j=0
		def SKSearchIndex
		for (int i = 1; i<= jmlRow;i++) {
			SKSearchIndex = Integer.parseInt(readExcel(filePath, sheetName, SKColumnIndex, i))
			if (SKSearchIndex >> PKIndex) {
				//WebUI.delay(1)
				break;
			}
			if (SKSearchIndex == PKIndex) {
				returnValue = readExcel(filePath, sheetName, targetColumnIndex, i)
				returnList[j] = returnValue
				//println("list dari "+sheetName+" index ke-"+j+" adalah "+returnList[j]+" main index : "+PKIndex)
				j++
			}
		}

		return returnList
	}
}
