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

GlobalVariable.sheetDir = (GlobalVariable.xlsDir + 'EDMI_Entrance_Test_Datafile.xls')

String StartDataRaw = CustomKeywords.'commonFunctions.excelFunctons.readExcel'(GlobalVariable.sheetDir, 'StartEnd', 0, 1)

String EndDataRaw = CustomKeywords.'commonFunctions.excelFunctons.readExcel'(GlobalVariable.sheetDir, 'StartEnd', 1, 1)

int StartData

int EndData

if (StartDataRaw == '-') {
    StartData = 1
} else {
    StartData = Integer.parseInt(StartDataRaw)
}

if (EndData == '-') {
    EndData = CustomKeywords.'commonFunctions.excelFunctons.getSheetLastRow'(GlobalVariable.sheetDir, 'Personal')
} else {
    EndData = Integer.parseInt(EndDataRaw)
}

for (int mainLoop = StartData; mainLoop <= EndData; mainLoop++) {
    WebUI.callTestCase(findTestCase('Pages/Open_Browser'), [('WebURL') : GlobalVariable.siteURL], FailureHandling.STOP_ON_FAILURE)

    WebUI.callTestCase(findTestCase('Pages/Login'), [('ADMusername') : 'admin', ('ADMpassword') : 'admin123'], FailureHandling.STOP_ON_FAILURE)

    WebUI.click(findTestObject('Object Repository/orangehrmlive/Dashboard/lnkMyInfo'))

    WebUI.callTestCase(findTestCase('Pages/Edit_Personal_Details'), [('mainIndex') : mainLoop], FailureHandling.STOP_ON_FAILURE)

    WebUI.callTestCase(findTestCase('Pages/Edit_Custom_Fields'), [('mainIndex') : mainLoop], FailureHandling.STOP_ON_FAILURE)

    WebUI.callTestCase(findTestCase('Pages/Upload_Attachment'), [('mainIndex') : mainLoop], FailureHandling.STOP_ON_FAILURE)

    WebUI.callTestCase(findTestCase('Pages/Logout'), [:], FailureHandling.STOP_ON_FAILURE)
}



