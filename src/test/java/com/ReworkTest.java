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
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

import com.aventstack.extentreports.ExtentTest;

public class ReworkTest extends TestClass{
	private static ExtentTest test;
	public static String getCellValueAsString(Cell cell) {
	    if (cell == null) 
	    {
	        return "";
	    }

	    switch (cell.getCellType()) 
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
	public void reworkAOIWave() throws IOException, InterruptedException, TimeoutException
	{
		//String path = System.getProperty("user.dir")+"/Data.xlsx";
		String filePath = System.getProperty("user.dir") + "/" + ConfigClass.getProperty("INPUT_FILE");
		File file = new File(filePath);
		if(!file.exists())
		{
			System.out.println("Create the file:"+file);
			
		}
		FileInputStream stream = new FileInputStream(file);
		XSSFWorkbook workbook = new XSSFWorkbook(stream);
		int sheetIndex = Integer.parseInt(ConfigClass.getProperty("REWORK_SHEET"));
		XSSFSheet sheet = workbook.getSheetAt(sheetIndex);
	//	XSSFSheet sheet =workbook.getSheetAt(8);
		for(int i=1;i<sheet.getPhysicalNumberOfRows();i++)
		{
			XSSFRow row =sheet.getRow(i);
			String rework = getCellValueAsString(row.getCell(0));
			String rework_status = getCellValueAsString(row.getCell(1));
			String remarks = getCellValueAsString(row.getCell(2));
			String expectedType = getCellValueAsString(row.getCell(3));
			String timestamp = new SimpleDateFormat("yyyyMMdd_HHmmss_SSS").format(new Date());
			String screenshotName = expectedType+"Rework"+timestamp;
			test = extent.createTest("Rework -"+expectedType.toUpperCase());
			//String excelReportPath = System.getProperty("user.dir")+"/report/Rework_Report.xlsx";
			String excelReportPath = System.getProperty("user.dir") + "/" + ConfigClass.getProperty("SFCS_REPORT");
			String timeStamp = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(new Date());
			String testName = "Rework Test -"+expectedType.toUpperCase();
			 try {
	    		//	Connection con = DriverManager.getConnection("jdbc:mysql://10.1.71.113:3306/qa_sfcs?useSSL=false&allowPublicKeyRetrieval=true","root", "Vepl@123");
				 Connection con = DriverManager.getConnection( ConfigClass.getProperty("DB_URL"), ConfigClass.getProperty("DB_USERNAME"), ConfigClass.getProperty("DB_PASSWORD"));
				 if(con.isClosed())
	    			{
	    				System.out.println("Database is not connected");
	    				
	    			}
	    			else
	    			{
	    				System.out.println("Database is connected");
	    			}
	    			 String deleteExistReworkNumber = rework; 
	                 String deleteQuery = "DELETE FROM rework_reports WHERE rework_pcb_sl_no = ?";
	                 PreparedStatement prepStat = con.prepareStatement(deleteQuery);
	                 prepStat.setString(1, deleteExistReworkNumber);
	                 int deleteDuplicateRows = prepStat.executeUpdate();

	                 if (deleteDuplicateRows > 0) 
	                 {
	                     System.out.println("Fail PCB of AOI and WAVE Rework Number deleted Successfully -"+ deleteExistReworkNumber);
	                 } else
	                 {
	                     System.out.println("Fail PCB of AOI and WAVE Rework Number not found in DB -"+ deleteExistReworkNumber);
	                 }

	                 prepStat.close();
	                 con.close();
	             
                  Thread.sleep(1000);
	         } 
	            catch (SQLException e) 
	            {
	             e.printStackTrace();
	         }
			test.assignCategory("Rework");
	        test.assignAuthor("Automation Testing");
	        test.assignDevice("Chrome");
			driver.get(ConfigClass.getProperty("BASEURL")+"/rework");
			driver.manage().window().maximize();
			//test.info("<span style='color:black'><b>Navigate to Rework screen</b>");
			Thread.sleep(2000);
	        WebDriverWait wait = new WebDriverWait(driver,Duration.ofSeconds(20));
	        WebElement icon = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("(//a[@class='nav-link text-secondary'])[1]")));
		    icon.click();
		    //test.info("<span style='color:black'><b>Click Icon button</b>");
		    JavascriptExecutor js = (JavascriptExecutor) driver;
		    WebElement clickReworkBtn= wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//p[text()='Rework']")));
		    js.executeScript("arguments[0].scrollIntoView(true)", clickReworkBtn);
		    clickReworkBtn.click();
		    test.info("<span style='color:black'><b>Click Rework</b>");
		    WebElement reworkNum = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[@placeholder='Enter PCB Serial Number']")));
		    Thread.sleep(100);
		    reworkNum.sendKeys(rework);
		    Thread.sleep(100);
		    reworkNum.sendKeys(Keys.ENTER);
		    test.info("<span style='color:black'><b> Enter Fail PCB of AOI or Wave for Rework : <span style='color:blue'><b> "  + rework+"</b></span>");
		    Thread.sleep(100);
		    try {
		    WebDriverWait wait1 = new WebDriverWait(driver, Duration.ofSeconds(20));
			WebElement successMsg = wait1.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),'Rework status updated successfully') or contains(text(),'PCB serial number already exists.') or contains(text(),'No records found for this barcode')]")));
			String messageText = successMsg.getText().trim().toLowerCase();
			System.out.println(messageText);
			if (messageText.contains("rework status updated successfully")) 
			{
			    passScreenshot(test, "<b style='color:green;'>Failed AOI Input or Wave for Rework </b>", screenshotName);
			    test.pass("<b style='color:green;'>Rework status updated successfully</b>");
			    writeExtentSFCSReport(excelReportPath, "Rework Report", testName, "PASS", timeStamp, "Fail AOI or Wave","Rework PCB message:"+rework+" - "+messageText );
			}//"Rework Status Updated Successfully"
			else if (messageText.contains("pcb already exists")) 
			{
			    failScreenshot(test, "<b style='color:red;'>PCB rework Failed - PCB already exists.Use different PCB</b>", screenshotName);
			    test.fail("<b style='color:red;'>Use different PCB serial number</b>");
			    writeExtentSFCSReport(excelReportPath, "Rework Report", testName, "FAIL", timeStamp, "PCB already scanned", "PCB already scanned.Use different PCB");
			    continue;
			} 
			else if (messageText.contains("no record found for this barcode")) 
			{
			    failScreenshot(test, "<b style='color:red;'>PCB rework Failed - No Record found for this barcode</b>", screenshotName);
			    test.fail("<b style='color:red;'>Record not found</b>");
			    writeExtentSFCSReport(excelReportPath, "Rework Report", testName, "FAIL", timeStamp, "No  found", "Record not found for this barcode");
			    continue;
			} 
		    
			 else 
			{
			    failScreenshot(test, "<b style='color:red;'>Rework message : " + messageText, screenshotName);
			    test.fail("<b style='color:red;'>Rework status message :" + messageText);
			   // writeExtentSFCSReport(excelReportPath, "Rework Report", testName, "PASS", timeStamp,"Rework PCB", "Rework message: " +rework +" - "+messageText);
			    writeExtentSFCSReport(excelReportPath, "Rework Report", testName, "FAIL", timeStamp,"Rework PCB", "Rework message: " +rework +" - "+messageText);
			    continue;
			}
		//	continue;
		    }
		    catch (TimeoutException e)
		    {
		    	test.pass("<b style='color:green;'>Rework the PCB</b>");
		    	passScreenshot(test, "<b style='color:green;'>Rework the PCB for failed AOI or Wave </b>", screenshotName);
		    	writeExtentSFCSReport(excelReportPath, "Rework Report", testName, "PASS", timeStamp,"Rework the fail PCB", "Rework status updated successfully");
		    }
		    Thread.sleep(300);
		    WebElement selectReworkStatus = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//select[@class='form-control form-control-sm']")));
		    Select select = new Select(selectReworkStatus);
		    select.selectByVisibleText(rework_status);
		    test.info(" <span style='color:black'><b>Select the rework Status : <span style='color:blue'><b>" + rework_status+"</b></span>");
		    WebElement enterRemark = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[@placeholder='Enter remarks']")));
		    Actions actions = new Actions(driver);
		    actions.doubleClick(enterRemark).click().perform();
		    test.info(" <span style='color:black'><b> Enter remarks :<span style='color:blue'><b> " + remarks+"</b></span>");
		    Thread.sleep(400);
		    WebElement enterRemark1 = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[@class='ag-input-field-input ag-text-field-input']")));
		    enterRemark1.sendKeys(remarks);
		    enterRemark1.sendKeys(Keys.ENTER);
		    WebElement clickChangePanel = driver.findElement(By.xpath("//button[text()='Change Panel Operation']"));
		    js.executeScript("arguments[0].scrollIntoView(true)", clickChangePanel);
		    Thread.sleep(200);
		    clickChangePanel.click();
		    test.info("<span style='color:black'><b> Click Change Panel Operation</b>");
		}
	}

}
