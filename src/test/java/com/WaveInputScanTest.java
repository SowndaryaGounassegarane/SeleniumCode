package com;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.TimeoutException;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import com.aventstack.extentreports.ExtentTest;

public class WaveInputScanTest extends TestClass{
	private static ExtentTest test;
	public static String getCellValueAsString(Cell cell)
	{
		if(cell==null)
		{
			return"";
	}
		switch(cell.getCellType())
		{
		case STRING:
			return cell.getStringCellValue();
		 case NUMERIC:
	            return String.valueOf(cell.getNumericCellValue());
	        case BOOLEAN:
	            return String.valueOf(cell.getBooleanCellValue());
	        case FORMULA:
	            return cell.getCellFormula();
	        case BLANK:
	            return "";
	        default:
	            return cell.toString();
	    }
	}

	@SuppressWarnings("resource")
    @Test
	public void waveInput() throws IOException, InterruptedException
	{
		//String path = System.getProperty("user.dir")+"/Data.xlsx";
		String filePath = System.getProperty("user.dir") + "/" + ConfigClass.getProperty("INPUT_FILE");
		File file = new File(filePath);
		if(!file.exists())
		{
			System.out.println("Create file:"+file);
		}
		FileInputStream stream = new FileInputStream(file);
		XSSFWorkbook workbook = new XSSFWorkbook(stream);
		int sheetIndex = Integer.parseInt(ConfigClass.getProperty("WAVEINPUTSCAN_SHEET"));
		XSSFSheet sheet = workbook.getSheetAt(sheetIndex);
		//XSSFSheet sheet = workbook.getSheetAt(5);
		for(int i=1;i<sheet.getPhysicalNumberOfRows();i++)
		{
			XSSFRow row =sheet.getRow(i);
			String workOrder = getCellValueAsString(row.getCell(0));
			String waveInput = getCellValueAsString(row.getCell(1));
			String expectedType = getCellValueAsString(row.getCell(2));
			String timestamp = new SimpleDateFormat("yyyyMMdd_HHmmss_SSS").format(new Date());
			String screenshotName = expectedType+"Wave Input"+timestamp;
			test = extent.createTest("Wave Input -"+expectedType.toUpperCase());
			//String excelReportPath = System.getProperty("user.dir")+"/report/Wave_Input_Report.xlsx";
			String excelReportPath = System.getProperty("user.dir") + "/" + ConfigClass.getProperty("SFCS_REPORT");
			String timeStamp = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(new Date());
			String testName = "Wave Input Test -"+expectedType.toUpperCase();
			 try {
				 Connection con = DriverManager.getConnection( ConfigClass.getProperty("DB_URL"), ConfigClass.getProperty("DB_USERNAME"), ConfigClass.getProperty("DB_PASSWORD"));
	    			//Connection con = DriverManager.getConnection("jdbc:mysql://10.1.71.113:3306/qa_sfcs?useSSL=false&allowPublicKeyRetrieval=true","root", "Vepl@123");
	    			if(con.isClosed())
	    			{
	    				System.out.println("Database is not connected");
	    				
	    			}
	    			else
	    			{
	    				System.out.println("Database is connected");
	    			}
	    			 String deleteExistWaveInputNumber = waveInput; 
	                 String deleteQuery = "DELETE FROM wave_reports WHERE wave_pcb_sl_no = ?";
	                 PreparedStatement prepStat = con.prepareStatement(deleteQuery);
	                 prepStat.setString(1, deleteExistWaveInputNumber);
	                 int deleteDuplicateRows = prepStat.executeUpdate();

	                 if (deleteDuplicateRows > 0) 
	                 {
	                     System.out.println("WAVE Input Number deleted Successfully -"+ deleteExistWaveInputNumber);
	                 } else
	                 {
	                     System.out.println("WAVE Input Number not found in DB -"+ deleteExistWaveInputNumber);
	                 }

	                 prepStat.close();
	                 con.close();
	             
                  Thread.sleep(1000);
	         } 
	            catch (SQLException e) 
	            {
	             e.printStackTrace();
	         }
			 test.assignCategory("Wave IP Scan");
	         test.assignAuthor("Automation Testing");
	         test.assignDevice("Chrome");
			driver.get(ConfigClass.getProperty("BASEURL")+"/wave-ip-scan/");
			driver.manage().window().maximize();
			//test.info("<span style='color:black'><b> Navigated to Wave Input Scan Page</b>");
			WebDriverWait wait = new WebDriverWait(driver,Duration.ofSeconds(20));
			WebElement icon = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("(//a[@class='nav-link text-secondary'])[1]")));
		    icon.click();
		    //test.info(" <span style='color:black'><b>Click Icon button</b>");
            Thread.sleep(500);
			WebElement clickWaveInputScan= wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//p[text()='Wave IP Scan']")));
			clickWaveInputScan.click();
			test.info("<span style='color:black'><b>Click Wave Input scan</b>");
			 Thread.sleep(500);
		        WebElement dropdown = wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector(".css-13cymwt-control")));
		        dropdown.click();
		        test.info("<b>Click Work Order dropdown</b>");
		        Thread.sleep(500);
		        String optionXpath1 = "//div[contains(@class,'-option') and contains(normalize-space(), '" + workOrder + "')]";
		        WebElement woOption = wait.until(ExpectedConditions.elementToBeClickable(By.xpath(optionXpath1)));
		        woOption.click();
		        test.info("<b>Selected Work Order: </b><span style='color:blue'><b>" + workOrder + "</b></span>");
//			WebElement selctWorkOrder =wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector(".css-13cymwt-control")));
//			selctWorkOrder.click();
//			test.info("<span style='color:black'><b>Click the Work Order for Wave</b>");
//			test.info("<span style='color:black'><b> Select the Work Order for Wave :<span style='color:blue'><b>" +workOrder+"</b></span>");
			WebElement waveInputScan = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[@placeholder='Scan or enter PCB Serial Number']")));
			JavascriptExecutor js = (JavascriptExecutor) driver;
			js.executeScript("arguments[0].scrollIntoView(true);", waveInputScan);
		    js.executeScript("arguments[0].setAttribute('value','')", waveInputScan);
		    Thread.sleep(500);
		    waveInputScan.sendKeys(waveInput);
		    waveInputScan.sendKeys(Keys.ENTER);
		    test.info("<span style='color:black'><b>Scan the Pass AOI PCB serial number in Wave Input scan:<span style='color:blue'><b> " + waveInput+"</b></span>");
		    test.info("<b style='color:green;'> PCB Number Entered Successfully</b>");
		    Thread.sleep(500);
		    try 
		    {
		        WebDriverWait wait1 = new WebDriverWait(driver, Duration.ofSeconds(20));
		        WebElement successMsg = wait1.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'successfully entered') or contains(text(),'Barcode invalid') or contains(text(),'AOI Failed') or contains(text(),'PCB already has been scanned')]")));
		        String messageText = successMsg.getText().trim().toLowerCase();
		        System.out.println(messageText);
		        if (messageText.contains("successfully entered")) 
		        {
		            passScreenshot(test, "<b style='color:green;'>Pass AOI Input to WaveInputScan</b>", screenshotName);
		            test.pass("<b style='color:green;'>PCB entered successfully</b>");
		            writeExtentSFCSReport(excelReportPath, "Wave Input", testName, "PASS", timeStamp, "AOI Pass PCB","Wave PCB message:"+waveInput+" - "+messageText);
		        }//- PCB Entered Successfully"
		        else if (messageText.contains("Barcode Invalid")) 
		        {
		            failScreenshot(test, "<b style='color:red;'>PCB Failed - Barcode Invalid not passed from previous stages</b>", screenshotName);
		            test.fail("<b style='color:red;'>PCB Barcode is Invalid or not passed from previous stages</b>");
		            writeExtentSFCSReport(excelReportPath, "Wave Input", testName, "FAIL", timeStamp, "Invalid AOI PCB To Wave Input Scan", "Barcode Invalid");
		        } 
		        else if (messageText.contains("AOI Failed"))
		        {
		            failScreenshot(test, "<b style='color:red;'>PCB Failed - AOI Failed record cannot be Inserted</b>", screenshotName);
		            test.fail("<b style='color:red;'>AOI Failed - this record cannot be inserted</b>");
		            writeExtentSFCSReport(excelReportPath, "Wave Input", testName, "FAIL", timeStamp, "AOI Fail PCB To Wave Input Scan", "AOI Failed");
		        }
		        else if (messageText.contains("PCB already scanned"))
		        {
		            failScreenshot(test, "<b style='color:red;'>PCB Failed - PCB Number already been scanned</b>", screenshotName);
		            test.fail("<b style='color:red;'>AOI Failed - this record cannot be inserted</b>");
		            writeExtentSFCSReport(excelReportPath, "Wave Input", testName, "FAIL", timeStamp, "AOI Fail PCB To Wave Input Scan", "PCB already scanned");
		        }
		        else 
		        {
		            failScreenshot(test, "<b style='color:red;'>Fail message : " + messageText, screenshotName);
		            test.fail("Failed PCB message: " + messageText);
		            writeExtentSFCSReport(excelReportPath, "Wave Input", testName, "FAIL", timeStamp,"Pass AOI PCB to wave", "Wave PCB message: " +waveInput +" - "+messageText);
		        
		        }
		    } 
		    catch (TimeoutException e)
		    {	       
		    	test.pass("<b style='color:green;'>AOI Pass PCB</b>");
		    	passScreenshot(test, "<b style='color:green;'>Pass PCB of AOI entered into Wave </b>", screenshotName);
		    	writeExtentSFCSReport(excelReportPath, "Rework Report", testName, "PASS", timeStamp,"Wave PCB", "PCB entered successfully");
		    }
		}
	} 

}
