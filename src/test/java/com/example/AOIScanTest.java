package com.example;

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
import org.openqa.selenium.support.ui.WebDriverWait;

import com.ConfigClass;
import com.TestClass;
import com.aventstack.extentreports.ExtentTest;

public class AOIScanTest extends TestClass{
	@SuppressWarnings("unused")
    private static ExtentTest test;
	public static String getCellValueAsString(Cell cell) {
	    if (cell == null) {
	        return "";
	    }

	    switch (cell.getCellType()) {
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
	public void AOIScreen() throws IOException, InterruptedException
	{
		
	//String path =System.getProperty("user.dir")+"/Data.xlsx";
   String filePath = System.getProperty("user.dir") + "/" + ConfigClass.getProperty("INPUT_FILE");
	File file = new File(filePath);
	if(!file.exists())
	{
		System.out.println("create file:"+file);
	}
	FileInputStream stream = new FileInputStream(file);
	XSSFWorkbook workbook = new XSSFWorkbook(stream);
	int sheetIndex = Integer.parseInt(ConfigClass.getProperty("AOISCAN_SHEET"));
	XSSFSheet sheet = workbook.getSheetAt(sheetIndex);
	//XSSFSheet sheet=workbook.getSheetAt(4);
	for(int i=1;i<sheet.getPhysicalNumberOfRows();i++)
	{
		XSSFRow row =sheet.getRow(i);
		String workOrder = getCellValueAsString(row.getCell(0));
		String AOIPass =getCellValueAsString(row.getCell(1));
		String AOIFail =getCellValueAsString(row.getCell(2));
		String AOI_pass = getCellValueAsString(row.getCell(3));
		String defectType =getCellValueAsString(row.getCell(4));
		String defectLocation=getCellValueAsString(row.getCell(5));
		String expectedType = getCellValueAsString(row.getCell(6));
		String timestamp = new SimpleDateFormat("yyyyMMdd_HHmmss_SSS").format(new Date());
		//String screenshotName = expectedType+"AOI"+timestamp;
		String passScreenshotName = expectedType + "_AOIPASS_" + timestamp;
		String failScreenshotName = expectedType + "_AOIFAIL_" + timestamp;
		ExtentTest testPass = extent.createTest("AOI Report - PASS: " + AOIPass);
		ExtentTest testFail = extent.createTest("AOI Report - FAIL: " + AOIFail);
		//test = extent.createTest("AOI Report - "+expectedType.toUpperCase());
		//String excelReportPath = System.getProperty("user.dir")+"/report/AOI_Report.xlsx";
		String excelReportPath = System.getProperty("user.dir") + "/" + ConfigClass.getProperty("SFCS_REPORT");
		String timeStamp = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(new Date());
		String testName = "AOI Report -"+expectedType.toUpperCase();
		try 
		{

		  //  Connection con = DriverManager.getConnection("jdbc:mysql://10.1.71.113:3306/qa_sfcs?useSSL=false&allowPublicKeyRetrieval=true","root", "Vepl@123");
			Connection con = DriverManager.getConnection( ConfigClass.getProperty("DB_URL"), ConfigClass.getProperty("DB_USERNAME"), ConfigClass.getProperty("DB_PASSWORD"));
		    if (con.isClosed())
		    {
		        System.out.println("Database is not connected");
		    } else
		    {
		        System.out.println("Database is connected");
		    }
		    String deleteQuery = "DELETE FROM aoi_reports WHERE aoi_pcb_sl_no = ?";
		    PreparedStatement prepStat = con.prepareStatement(deleteQuery);
		    prepStat.setString(1, AOIPass);
		    int deletedPass = prepStat.executeUpdate();
		    prepStat.setString(1, AOIFail);
		    int deletedFail = prepStat.executeUpdate();
		    if (deletedPass > 0 || deletedFail > 0) 
		    {
		        System.out.println("AOI PCB number deleted successfully : " +(deletedPass > 0 ? AOIPass : "") + " " +(deletedFail > 0 ? AOIFail : ""));
		    } else 
		    {
		        System.out.println("AOI PCB numbers not found in DB : " + AOIPass + ", " + AOIFail);
		    }
            prepStat.close();
		    con.close();
            Thread.sleep(1000);

		} 
		catch (SQLException | InterruptedException e)
		{
		    e.printStackTrace();
		}
		testPass.assignCategory("AOI Scan");
		testPass.assignAuthor("PASS PCB");
		testPass.assignDevice("Chrome");
		testFail.assignCategory("AOI Scan");
		testFail.assignAuthor("FAIL PCB");
		testFail.assignDevice("Chrome");
		System.out.println("AOIPASS:"+AOIPass);
		System.out.println("AOIFAIL:"+AOIFail);
		driver.manage().window().maximize();
        driver.get(ConfigClass.getProperty("BASEURL") + "/aoi-scan/");  
        testPass.info("<span style='color:black'><b> Navigate to the AOI scan page</b>");
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(15));
        JavascriptExecutor js = (JavascriptExecutor) driver;
        Actions actions = new Actions(driver);
        Thread.sleep(5000);
        WebElement dropdown = wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector(".css-13cymwt-control")));
        dropdown.click();
        testPass.info("<b>Click Work Order dropdown</b>");
        Thread.sleep(10000);
        String optionXpath1 = "//div[contains(@class,'-option') and contains(normalize-space(), '" + workOrder + "')]";
        WebElement woOption = wait.until(ExpectedConditions.elementToBeClickable(By.xpath(optionXpath1)));
        woOption.click();
        testPass.info("<b>Selected Work Order: </b><span style='color:blue'><b>" + workOrder + "</b></span>");
        System.out.println(workOrder);

//        WebElement clickWorkOrder = wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector(".css-13cymwt-control")));
//        clickWorkOrder.click();
//        testPass.info("<span style='color:black'><b> Select the Work Order for AOI</b>");
//        testPass.info("<span style='color:black'><b> Click Work Order</b>");
        WebElement enterPassPcb = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("(//input[@placeholder='Scan PCB Serial Number'])[1]")));
        js.executeScript("arguments[0].scrollIntoView(true);", enterPassPcb);
      //  js.executeScript("arguments[0].value='';", enterPassPcb); 
       // js.executeScript("arguments[0].setAttribute('value','')", enterPassPcb);
        js.executeScript("arguments[0].scrollIntoView({block: 'center', inline: 'nearest'});", enterPassPcb);
        Thread.sleep(2000);
        enterPassPcb.sendKeys(AOIPass);
        enterPassPcb.sendKeys(Keys.ENTER);
        testPass.info("<span style='color:black'><b> Enter PCB serial number for PASS Boards : <span style='color:blue'><b>   " + AOIPass+"</b></span>");
        testPass.info("<span style='color:green'><b> PCB marked as passed</b>");
        Thread.sleep(2000);
        System.out.println("Enter AOI Pass: " + enterPassPcb.getAttribute("value"));
        WebElement enterFailPcb = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("(//input[@placeholder='Scan PCB Serial Number'])[2]")));
        js.executeScript("arguments[0].scrollIntoView(true);", enterFailPcb);
      //  js.executeScript("arguments[0].value='';", enterFailPcb);
       // js.executeScript("arguments[0].setAttribute('value','')", enterFailPcb);
        js.executeScript("arguments[0].scrollIntoView({block: 'center', inline: 'nearest'});", enterFailPcb);
        Thread.sleep(2000);
        enterFailPcb.sendKeys(AOIFail);
        testFail.info("<span style='color:black'><b> Enter the PCB scan for FAIL boards :<span style='color:blue'><b>  "+AOIFail+"</b></span>");
        WebElement defectTypeDropdown = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//div[label[contains(text(),'Defect Type')]]//div[contains(@class, 'css-13cymwt-control')]")));
        js.executeScript("arguments[0].scrollIntoView(true);", defectTypeDropdown);
        defectTypeDropdown.click();
        Thread.sleep(2000);  
        String optionXpath = "//div[contains(@class,'-menu')]//div[contains(@class,'-option') and normalize-space()='" + defectType + "']";
        WebElement defectOption = driver.findElement(By.xpath(optionXpath));
        js.executeScript("arguments[0].scrollIntoView(true);", defectOption);
        js.executeScript("arguments[0].scrollIntoView({block: 'center', inline: 'nearest'});", defectOption);
        Thread.sleep(2000);
        js.executeScript("arguments[0].click();", defectOption);
        Thread.sleep(2000);
        System.out.println("Select defect type: " + defectType);
        testFail.info("<span style='color:black'><b> Select Defect Type : <span style='color:blue'><b>  " + defectType+"</b></span>");
        WebElement enterDefectLocation = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[@placeholder='Enter Defect Location']")));
        js.executeScript("arguments[0].scrollIntoView(true);", enterDefectLocation);
        js.executeScript("arguments[0].scrollIntoView({block: 'center', inline: 'nearest'});", enterDefectLocation);
        enterDefectLocation.clear();
        Thread.sleep(2000);
        enterDefectLocation.sendKeys(defectLocation);
        Thread.sleep(2000);
        System.out.println("Enter Defect Location: " + defectLocation);
        testFail.info(" <span style='color:black'><b>Enter Defect Location : <span style='color:blue'><b>  " + defectLocation+"</b></span>");
        testFail.info("<span style='color:red'><b>PCB marked as failed </b>");
        actions.moveToElement(enterFailPcb).perform();
        Thread.sleep(2000);
        enterFailPcb.sendKeys(Keys.ENTER);
        Thread.sleep(2000);
        enterPassPcb.sendKeys(AOI_pass);
        enterPassPcb.sendKeys(Keys.ENTER); 
        WebElement clickRadioStatus = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//div[@class='form-check me-4']")));
        clickRadioStatus.click();
        WebElement clickUpdate = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//button[text()=' Update']")));
        clickUpdate.click();
        try {
            WebDriverWait wait1 = new WebDriverWait(driver, Duration.ofSeconds(30));
            WebElement successMsg = wait1.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(),' marked as passed') or contains(text(),' marked as failed') or contains(text(),'PCB updated to pass') or contains(text(),'updated as passed')]")));            
            String msgText = successMsg.getText().trim();
            System.out.println("AOI Message: " + msgText);
            if (msgText.contains("marked as passed") || msgText.contains("updated as passed") || msgText.contains("PCB updated to pass")) {
                Thread.sleep(1000);
                passScreenshot(testPass, "<b style='color:green;'>PCB marked as passed</b>", passScreenshotName);
                testPass.pass("<b style='color:green;'>PCB marked passed</b>");                
                System.out.println("Writing PASS to Excel for: " + AOIPass);
                writeExtentSFCSReport(excelReportPath, "AOI Scan", testName, "PASS", timeStamp,"Scan Pass PCB", "PCB marked as passed: " + AOIPass);
            }             
            else if (msgText.contains("marked as failed"))
            {
                failScreenshot(testFail, "<b style='color:red;'>PCB Failed - Select Defect Type and Defect Location</b>", failScreenshotName);
                testFail.fail("<b style='color:red;'>PCB marked as failed</b>");  
                System.out.println("Writing FAIL to Excel for: " + AOIFail);
                writeExtentSFCSReport(excelReportPath, "AOI Scan", testName, "FAIL", timeStamp,"Scan PCB as Fail", "PCB marked as failed: " + AOIFail);
                continue;
            }             
            else 
            {
                failScreenshot(testFail, "<b style='color:red;'>Fail message : " + msgText + "</b>", failScreenshotName);
                testFail.fail("Failed PCB message: " + msgText);
                writeExtentSFCSReport(excelReportPath, "AOI Scan", testName, "FAIL", timeStamp,"Scan AOI Fail", "PCB Fail");
                continue;
            }
        } 
        catch (TimeoutException e) {
            passScreenshot(testPass, "<b style='color:green;'>PCB marked as passed (Timeout)</b>", passScreenshotName);
            testPass.pass("<b style='color:green;'>PCB marked as passed (Timeout)</b>");
            writeExtentSFCSReport(excelReportPath, "AOI Scan", testName, "PASS", timeStamp, "Scan AOI Pass", "PCB marked as passed by timeout: " + AOIPass);
        }

}
	}

}
