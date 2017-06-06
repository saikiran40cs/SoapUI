package Kiran.ExcelReader

// IMPORT THE LIBRARIES WE NEED
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;

import com.eviware.soapui.support.XmlHolder

// DECLARE THE VARIABLES

def myTestCase = context.testCase //myTestCase contains the test case
def counter,next,previous,size //Variables used to handle the loop and to move inside the file
def path = /C:\KiranTest\soapUI_TestData\dataFile.xlsx/; //file containing the data

FileInputStream inputStream = new FileInputStream(path);
XSSFWorkbook workbook1 = new XSSFWorkbook(inputStream);//create instance of workbook
XSSFSheet sheet1 = workbook1.getSheetAt(0) //save the first sheet in sheet1
size= sheet1.getLastRowNum()+1//get the number of rows, each row is a data set

propTestStep = myTestCase.getTestStepByName("VarProperties") // get the Property TestStep object
propTestStep.setPropertyValue("Total", size.toString())     //set the total number of times to execute the script
counter = propTestStep.getPropertyValue("Count").toString() //counter variable contains iteration number
counter = counter.toInteger() //Parse counter to integer
next = (counter > size-2? 0: counter+1) //set the next value
 
// OBTAINING THE DATA YOU NEED
XSSFRow ro = sheet1.getRow(counter) //Based on counter get the row value
XSSFCell uDataFromXL = ro.getCell(0) // getCell(column,row) //obtains details from column 1

workbook1.close() //close the file

//Get contents from XL and Store in a variable
usrSearch = uDataFromXL.toString()
//set the property value in Groovy step
propTestStep.setPropertyValue("TextToSearch", usrSearch) //the value is saved in the property
propTestStep.setPropertyValue("Count", next.toString()) //increase Count value

next++ //increase next value
propTestStep.setPropertyValue("Next", next.toString()) //set Next value on the properties step

//Decide if the test has to be run again or not
if (counter == size-1)
{
	propTestStep.setPropertyValue("StopLoop", "T")
	log.info "Setting the stoploop property now..."
}
else if (counter==0)
{
	def runner = new com.eviware.soapui.impl.wsdl.testcase.WsdlTestCaseRunner(testRunner.testCase, null)
	propTestStep.setPropertyValue("StopLoop", "F")
}
else
{
	propTestStep.setPropertyValue("StopLoop", "F")
}