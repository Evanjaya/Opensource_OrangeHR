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
import org.openqa.selenium.Keys as Keys

WebUI.click(findTestObject('Object Repository/orangehrmlive/My_Info/btnEditPersonalDetail'))

String firstName = CustomKeywords.'commonFunctions.excelFunctons.readExcel'(GlobalVariable.sheetDir, "Personal", 1, mainIndex)
WebUI.setText(findTestObject('Object Repository/orangehrmlive/Personal_Details/txtFirstName'), firstName)

String middleName = CustomKeywords.'commonFunctions.excelFunctons.readExcel'(GlobalVariable.sheetDir, "Personal", 2, mainIndex)
if (middleName != "-") {
	WebUI.setText(findTestObject('Object Repository/orangehrmlive/Personal_Details/txtMidName'), middleName)
}

String lastName = CustomKeywords.'commonFunctions.excelFunctons.readExcel'(GlobalVariable.sheetDir, "Personal", 3, mainIndex)
WebUI.setText(findTestObject('Object Repository/orangehrmlive/Personal_Details/txtLastName'),lastName)

String employeeID = CustomKeywords.'commonFunctions.excelFunctons.readExcel'(GlobalVariable.sheetDir, "Personal", 4, mainIndex)
if (employeeID != "-") {
	WebUI.setText(findTestObject('Object Repository/orangehrmlive/Personal_Details/txtEmployeeID'),employeeID)
}

String otherID = CustomKeywords.'commonFunctions.excelFunctons.readExcel'(GlobalVariable.sheetDir, "Personal", 5, mainIndex)
if (otherID != "-") {
	WebUI.setText(findTestObject('Object Repository/orangehrmlive/Personal_Details/txtOtherID'), otherID)
}

String SIMNo = CustomKeywords.'commonFunctions.excelFunctons.readExcel'(GlobalVariable.sheetDir, "Personal", 6, mainIndex)
if (SIMNo != "-") {
	WebUI.setText(findTestObject('Object Repository/orangehrmlive/Personal_Details/txtSIMno'), SIMNo)
}

String SIMExpDate = CustomKeywords.'commonFunctions.excelFunctons.readExcel'(GlobalVariable.sheetDir, "Personal", 7, mainIndex)
if (SIMExpDate != "-") {
	WebUI.setText(findTestObject('Object Repository/orangehrmlive/Personal_Details/txtSIMExpDate'), SIMExpDate)
	//WebUI.sendKeys(findTestObject(findTestObject('Object Repository/orangehrmlive/Personal_Details/txtSIMExpDate')), Keys.chord(Keys.ENTER))
	WebUI.click(findTestObject('Object Repository/orangehrmlive/Personal_Details/txtSIMno'))
}

String SSNNo = CustomKeywords.'commonFunctions.excelFunctons.readExcel'(GlobalVariable.sheetDir, "Personal", 8, mainIndex)
if (SSNNo != "-") {
	WebUI.setText(findTestObject('Object Repository/orangehrmlive/Personal_Details/txtSSNno'), SSNNo)
}

String SINNo = CustomKeywords.'commonFunctions.excelFunctons.readExcel'(GlobalVariable.sheetDir, "Personal", 9, mainIndex)
if (SINNo != "-") {
	WebUI.setText(findTestObject('Object Repository/orangehrmlive/Personal_Details/txtSINno'), SINNo)
}

String gender = CustomKeywords.'commonFunctions.excelFunctons.readExcel'(GlobalVariable.sheetDir, "Personal", 10, mainIndex)
boolean genderCheck
if (gender == "Male") {
	genderCheck = WebUI.verifyElementChecked(findTestObject('Object Repository/orangehrmlive/Personal_Details/rdrGenderMale'), 
		GlobalVariable.waitShort)
	if (!genderCheck) {
		WebUI.check(findTestObject('Object Repository/orangehrmlive/Personal_Details/rdrGenderMale'))
	}
} else if(gender == "Female") {
	WebUI.check(findTestObject('Object Repository/orangehrmlive/Personal_Details/rdrGenderFemale'))
}

String matrialStatus = CustomKeywords.'commonFunctions.excelFunctons.readExcel'(GlobalVariable.sheetDir, "Personal", 11, mainIndex)
WebUI.selectOptionByLabel(findTestObject('Object Repository/orangehrmlive/Personal_Details/selMatrialStatus'), matrialStatus, false)

String nationality = CustomKeywords.'commonFunctions.excelFunctons.readExcel'(GlobalVariable.sheetDir, "Personal", 12, mainIndex)
WebUI.selectOptionByLabel(findTestObject('Object Repository/orangehrmlive/Personal_Details/selNationality'), nationality, false)

String birthDate = CustomKeywords.'commonFunctions.excelFunctons.readExcel'(GlobalVariable.sheetDir, "Personal", 13, mainIndex)
if (birthDate != "-") {
	WebUI.setText(findTestObject('Object Repository/orangehrmlive/Personal_Details/txtDateBirth'), birthDate)
	WebUI.click(findTestObject('Object Repository/orangehrmlive/Personal_Details/txtSINno'))
}

String NickName = CustomKeywords.'commonFunctions.excelFunctons.readExcel'(GlobalVariable.sheetDir, "Personal", 14, mainIndex)
if (NickName != "-") {
	WebUI.setText(findTestObject('Object Repository/orangehrmlive/Personal_Details/txtNickName'), NickName)
}

String militaryService = CustomKeywords.'commonFunctions.excelFunctons.readExcel'(GlobalVariable.sheetDir, "Personal", 15, mainIndex)
if (militaryService != "-") {
	WebUI.setText(findTestObject('Object Repository/orangehrmlive/Personal_Details/txtMilitaryService'), militaryService)
}

String smoker = CustomKeywords.'commonFunctions.excelFunctons.readExcel'(GlobalVariable.sheetDir, "Personal", 16, mainIndex)
boolean hasChecked = false
if (smoker == "Yes") {
	hasChecked = WebUI.verifyElementChecked(findTestObject('Object Repository/orangehrmlive/Personal_Details/chkSmoker'), 
		GlobalVariable.waitShort, FailureHandling.OPTIONAL)
	if (!hasChecked) {
		WebUI.check(findTestObject('Object Repository/orangehrmlive/Personal_Details/chkSmoker'))
	}
} else {
	hasChecked = WebUI.verifyElementChecked(findTestObject('Object Repository/orangehrmlive/Personal_Details/chkSmoker'),
		GlobalVariable.waitShort)
	if (hasChecked) {
		WebUI.uncheck(findTestObject('Object Repository/orangehrmlive/Personal_Details/chkSmoker'))
	}	
}

WebUI.click(findTestObject('Object Repository/orangehrmlive/Personal_Details/btnSavePersonalDetails'))