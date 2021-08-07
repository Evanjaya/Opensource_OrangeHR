import static com.kms.katalon.core.checkpoint.CheckpointFactory.findCheckpoint
import static com.kms.katalon.core.testcase.TestCaseFactory.findTestCase
import static com.kms.katalon.core.testdata.TestDataFactory.findTestData
import static com.kms.katalon.core.testobject.ObjectRepository.findTestObject
import static com.kms.katalon.core.testobject.ObjectRepository.findWindowsObject
import com.kms.katalon.core.checkpoint.Checkpoint as Checkpoint
import com.kms.katalon.core.cucumber.keyword.CucumberBuiltinKeywords as CucumberKW
import com.kms.katalon.core.mobile.keyword.MobileBuiltInKeywords as Mobile
import com.kms.katalon.core.model.FailureHandling as FailureHandling
import com.kms.katalon.core.testcase.TestCase as TestCase
import com.kms.katalon.core.testdata.TestData as TestData
import com.kms.katalon.core.testng.keyword.TestNGBuiltinKeywords as TestNGKW
import com.kms.katalon.core.testobject.TestObject as TestObject
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.webui.keyword.WebUiBuiltInKeywords as WebUI
import com.kms.katalon.core.windows.keyword.WindowsBuiltinKeywords as Windows
import internal.GlobalVariable as GlobalVariable

ArrayList attcFileList = CustomKeywords.'commonFunctions.excelFunctons.getAllDetail'(GlobalVariable.sheetDir, "Attachment", mainIndex, 1, 2)
ArrayList attcComment = CustomKeywords.'commonFunctions.excelFunctons.getAllDetail'(GlobalVariable.sheetDir, "Attachment", mainIndex, 1, 3)
int arraySize = attcFileList.size()
String comment

for(int i=1;i<=arraySize;i++) {
	WebUI.click(findTestObject('Object Repository/orangehrmlive/My_Info/btnAddAttachment'))
	
	WebUI.uploadFile(findTestObject('Object Repository/orangehrmlive/Attachments/uplFileAttc'), 
		GlobalVariable.attachDir+attcFileList[i-1])
	
	comment = attcComment[i-1]
	if (comment != "-") {
		WebUI.setText(findTestObject('Object Repository/orangehrmlive/Attachments/txtComment'),comment)
	}
	
	WebUI.click(findTestObject('Object Repository/orangehrmlive/Attachments/btnUpload'))
}

/*
WebUI.click(findTestObject('Object Repository/orangehrmlive/My_Info/btnAddAttachment'))

String attachmentFileName = CustomKeywords.'commonFunctions.excelFunctons.readExcel'(GlobalVariable.sheetDir, "Attachment", 2, mainIndex)
WebUI.uploadFile(findTestObject('Object Repository/orangehrmlive/Attachments/uplFileAttc'), GlobalVariable.attachDir+attachmentFileName)

String comment = CustomKeywords.'commonFunctions.excelFunctons.readExcel'(GlobalVariable.sheetDir, "Attachment", 3, mainIndex)
if (comment != "-") {
	WebUI.setText(findTestObject('Object Repository/orangehrmlive/Attachments/txtComment'),comment)
}
*/

//WebUI.click(findTestObject('Object Repository/orangehrmlive/Attachments/btnUpload'))