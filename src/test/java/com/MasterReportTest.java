package com;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.TimeoutException;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

import com.aventstack.extentreports.ExtentTest;
import com.google.protobuf.TextFormat.ParseException;

public class MasterReportTest extends TestClass{
	private static ExtentTest test;
	public static String getCellValueAsString(Cell cell)
	{
		if(cell == null)
		{
			return "";
		}
		switch(cell.getCellType())
		{
		case STRING:
			return cell.getStringCellValue();
		case NUMERIC:
			return String.valueOf(cell.getNumericCellValue());
		case BOOLEAN:
			return String.valueOf(cell.getBooleanCellValue());
		case BLANK:
			return "";
		default:
			return cell.toString();
		}
	}

    public static String getDateFromCell(Cell cell) {
        if (cell == null) return "";

        if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
            Date date = cell.getDateCellValue();
            SimpleDateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy");
            return dateFormat.format(date);
        } else if (cell.getCellType() == CellType.STRING) {
            try {
                String val = cell.getStringCellValue().trim();
                SimpleDateFormat inputFormat = new SimpleDateFormat("d/M/yyyy");
                Date parsed = inputFormat.parse(val);
                return new SimpleDateFormat("MM/dd/yyyy").format(parsed);
            } catch (Exception e) {
                return "";
            }
        }
        return "";
    }

@SuppressWarnings({ "unused", "resource" })
@Test
public void masterReport() throws IOException, InterruptedException, ParseException
{
	//String path = System.getProperty("user.dir")+"/Data.xlsx";
	String filePath = System.getProperty("user.dir")+"/"+ConfigClass.getProperty("INPUT_FILE");
	File file = new File(filePath);
	if(!file.exists())
	{
		System.out.println("Create the file:"+file);
	}
	FileInputStream stream = new FileInputStream(file);
	XSSFWorkbook workbook = new XSSFWorkbook(stream);
	int sheetIndex = Integer.parseInt(ConfigClass.getProperty("MASTER_SHEET"));
	XSSFSheet sheet = workbook.getSheetAt(sheetIndex);
	//XSSFSheet sheet = workbook.getSheetAt(9);
	for(int i =1;i<sheet.getPhysicalNumberOfRows();i++)
	{
		XSSFRow row = sheet.getRow(i);	
		String chooseStartDate = getDateFromCell(row.getCell(0));
		String chooseEndDate = getDateFromCell(row.getCell(1));
		String reportType = getCellValueAsString(row.getCell(2)).trim();
		String reportStatus = getCellValueAsString(row.getCell(3)).trim();
		String expectedType = getCellValueAsString(row.getCell(4));
		String productName = getCellValueAsString(row.getCell(5));
		String timestamp = new SimpleDateFormat("yyyyMMdd_HHmmss_SSS").format(new Date());
		String screenshotName = expectedType+"Master Report"+timestamp;
		//test = extent.createTest("Master Report -"+expectedType.toUpperCase());
		//String excelReportPath = System.getProperty("user.dir")+"/report/Wave_Input_Report.xlsx";
		String excelReportPath = System.getProperty("user.dir") + "/" + ConfigClass.getProperty("SFCS_REPORT");
		String timeStamp = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(new Date());
		String testName = "Master Report -"+expectedType.toUpperCase();
		driver.get(ConfigClass.getProperty("BASEURL")+"/master-report");
		driver.manage().window().maximize();
		test = extent.createTest("Master Report");
		test.info("<span style='color:black'><b>Navigate to the Master Report Screen</b>");
		WebDriverWait wait = new WebDriverWait(driver,Duration.ofSeconds(20));
	    WebElement icon = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("(//a[@class='nav-link text-secondary'])[1]")));
		icon.click();
		test.info("<span style='color:black'><b>Click Icon button</b>");
		JavascriptExecutor js = (JavascriptExecutor) driver;
		Thread.sleep(500);
		WebElement clickMasterReport = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//p[text()='Master Reports']")));
		js.executeScript("arguments[0].scrollIntoView(true)", clickMasterReport);
		clickMasterReport.click();
		test.info("<span style='color:black'><b> Click Master Report</b>");
		Thread.sleep(500);
        Cell startDateCell = row.getCell(0); 
        Cell endDateCell = row.getCell(1);  
        DataFormatter formatter = new DataFormatter();
        String startDate = formatter.formatCellValue(startDateCell);
        String endDate = formatter.formatCellValue(endDateCell);
        WebElement startDateInput = driver.findElement(By.name("start_date")); 
        startDateInput.clear();
        startDateInput.sendKeys(chooseStartDate);
        WebElement endDateInput = driver.findElement(By.name("end_date"));
        endDateInput.clear();
        endDateInput.sendKeys(chooseEndDate);
        Thread.sleep(800);
		WebElement selectReportType = wait.until(ExpectedConditions.visibilityOfElementLocated(By.name("report_type")));
		Select select = new Select(selectReportType);
		select.selectByVisibleText(reportType);
		Thread.sleep(800);
		WebElement selectReportStatus = wait.until(ExpectedConditions.visibilityOfElementLocated(By.name("report_status")));
		Select select_Report_Status = new Select(selectReportStatus);
		select_Report_Status.selectByVisibleText(reportStatus);
		Thread.sleep(800);
		WebElement clickSearchButton = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//button[text()=' Search']")));
		clickSearchButton.click();	
		WebElement exportBtn = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//button[contains(.,'Export to CSV') or //span[contains(text(),'Export to CSV')]]")));
        js.executeScript("arguments[0].scrollIntoView({block: 'center'});", exportBtn);
        js.executeScript("document.body.style.zoom='90%'");
        js.executeScript("document.querySelector('.ag-body-horizontal-scroll-viewport').scrollLeft = 1000;");
        WebElement scrollTheReport = driver.findElement(By.cssSelector(".ag-body-horizontal-scroll-viewport"));
        js.executeScript("arguments[0].scrollLeft = arguments[0].scrollWidth", scrollTheReport);
		Thread.sleep(500);
		try
		{
			WebDriverWait wait1 = new WebDriverWait(driver,Duration.ofSeconds(10));
		    WebElement msg = wait1.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[text()='No records found for the selected filters']")));
		    String messageText = msg.getText().trim().toLowerCase();
		if(messageText.contains("record found"))
		{
		 passScreenshot(test, "<b style='color:green;'>Report generated for the AOI, Wave and Rework</b>", screenshotName);
         test.pass("<b style='color:green;'>PCB entered successfully</b>");
         writeExtentSFCSReport(excelReportPath, "Master Report", testName, "PASS", timeStamp, "Report","Master report");
		}
		else if(messageText.contains("No records found"))
		{
			failScreenshot(test, "<b style='color:red;'>No Reports - No records found for the selected filters", screenshotName);
	         test.fail("<b style='color:red;'>No records founds");
	         writeExtentSFCSReport(excelReportPath, "Master Report", testName, "FAIL", timeStamp, "No records", "Records not found for selected filters");
		}
		 else 
	        {
	            failScreenshot(test, "<b style='color:red;'>Fail Report : " + messageText, screenshotName);
	            test.fail("Failed PCB message: " + messageText);
	            writeExtentSFCSReport(excelReportPath, "Master Report", testName, "FAIL", timeStamp,"Reports not displayed", "No records found for selected filters" );
	        
	        }
	    } 
	    catch (TimeoutException e)
	    {	       
	    	test.pass("<b style='color:green;'>Reports generated for the AOI,Wave and rework</b>");
	    	passScreenshot(test, "<b style='color:green;'>Report Displayed for AOI, Wave and Rework </b>", screenshotName);
	    	writeExtentSFCSReport(excelReportPath, "Master Report", testName, "PASS", timeStamp,"Report displayed", "Report shown for selected filters");
	    }
	}
         
		
	}	 

}
