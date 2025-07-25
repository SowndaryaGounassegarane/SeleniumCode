package com.example;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.Date;
import java.util.List;

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
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import com.ConfigClass;
import com.TestClass;
import com.aventstack.extentreports.ExtentTest;

public class AOIReportTest extends TestClass {
	private static ExtentTest test;
	public static String getCellValueAsString(Cell cell)
	{
		if(cell==null)
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
	case FORMULA:
			return cell.getCellFormula();
	case BLANK:
		return"";
		default:
		return cell.toString();	
	}
	}
	public static String getDateFromCell(Cell cell)
	{
		if(cell==null)
		{
			return"";
		}
		if(cell.getCellType()==CellType.NUMERIC &&DateUtil.isCellDateFormatted(cell))
		{
			Date date = cell.getDateCellValue();
			SimpleDateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy");
			return dateFormat.format(date);
		}
			else if(cell.getCellType()==CellType.STRING)
			{
				try
				{
					String val = cell.getStringCellValue().trim();
					SimpleDateFormat format = new SimpleDateFormat("d/M/yyyy");
					Date parsed = format.parse(val);
					 return new SimpleDateFormat("MM/dd/yyyy").format(parsed);
	            } catch (Exception e) 
				{
	                return "";
	            }
	        }
	        return "";
			}
			@SuppressWarnings({ "unused", "resource" })
            @Test
	public void AOIReportData() throws FileNotFoundException, IOException, InterruptedException
	{
		String filePath = System.getProperty("user.dir")+"/"+ConfigClass.getProperty("INPUT_FILE");
		File file = new File(filePath);
		if(!file.exists())
		{
			System.out.println("Create the file:"+ file);
		}
		FileInputStream stream = new FileInputStream(file);
		XSSFWorkbook workbook =new XSSFWorkbook(stream);
		int sheetIndex = Integer.parseInt(ConfigClass.getProperty("AOIREPORT_SHEET"));
		XSSFSheet sheet = workbook.getSheetAt(sheetIndex);
		for(int i=1;i<sheet.getPhysicalNumberOfRows();i++)
		{
			XSSFRow row =sheet.getRow(i);
			String selectStartDate = getDateFromCell(row.getCell(0));
			String selectEndDate = getDateFromCell(row.getCell(1));
			String aoiScanStatus =getDateFromCell(row.getCell(2));
			String expectedType = getCellValueAsString(row.getCell(3));
			String timestamp = new SimpleDateFormat("yyyyMMdd_HHmmss_SSS").format(new Date());
			String screenshotName = expectedType+"AOI Report"+timestamp;
			//String excelReportPath = System.getProperty("user.dir")+"/report/Wave_Input_Report.xlsx";
			String excelReportPath = System.getProperty("user.dir") + "/" + ConfigClass.getProperty("SFCS_REPORT");
			String timeStamp = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(new Date());
			String testName = "AOI Report -"+expectedType.toUpperCase();
			driver.get(ConfigClass.getProperty("BASEURL")+"/aoi-reports/");
			driver.manage().window().maximize();
			test = extent.createTest("AOI Report");
			WebDriverWait wait = new WebDriverWait(driver,Duration.ofSeconds(20));
		    WebElement icon = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("(//a[@class='nav-link text-secondary'])[1]")));
			icon.click();
			WebElement clickAOIReports = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//p[text()='AOI Reports']")));
			clickAOIReports.click();
			Cell startDate = row.getCell(0);
			Cell endDate = row.getCell(1);
		    DataFormatter df = new DataFormatter();
		    String start_Date = df.formatCellValue(startDate);
	        String end_Date = df.formatCellValue(endDate);
		    WebElement clickStartDate = driver.findElement(By.xpath("(//input[@type='date'])[1]"));
		    clickStartDate.clear();
		    clickStartDate.sendKeys(selectStartDate);
		    WebElement clickEndDate =driver.findElement(By.xpath("(//input[@type='date'])[2]"));
		    clickEndDate.clear();
		    clickEndDate.sendKeys(selectEndDate);
		    Thread.sleep(2000);
		    WebElement clickSearch = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//button[text()=' Search']")));
	        clickSearch.click();
	        Thread.sleep(5000); 
	        List<WebElement> reportRows = driver.findElements(By.xpath("//div[@class='ag-center-cols-viewport']//div[@role='row']"));
	        if (reportRows.size() > 0)
	        {
	            passScreenshot(test, "<b style='color:green;'>AOI report successfully generated and displayed on the screen</b>", screenshotName);
	            test.pass("<b style='color:green;'> AOI report generated for scan status </b>"+aoiScanStatus);
	            writeExtentSFCSReport(excelReportPath, "AOI Report", testName, "PASS", timeStamp, "AOI report generated", "AOI report displayed on screen for scan status:"+aoiScanStatus);
	            String rowXPath = "//div[@role='row' and .//div[contains(text(),'" + aoiScanStatus + "')]]";
                WebElement targetRow = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(rowXPath)));
                JavascriptExecutor js = (JavascriptExecutor) driver;
                js.executeScript("arguments[0].scrollIntoView({block:'center'});", targetRow);
                Thread.sleep(500);
                js.executeScript("document.querySelector('.ag-body-horizontal-scroll-viewport').scrollLeft = 3000;");
                Thread.sleep(100);    
                WebElement clickHistoryBtn = targetRow.findElement(By.xpath("//button[@class='btn btn-sm btn-primary']"));
                js.executeScript("arguments[0].scrollIntoView({block:'center'});", clickHistoryBtn);
                wait.until(ExpectedConditions.elementToBeClickable(clickHistoryBtn)).click();
                Thread.sleep(2000);
                passScreenshot(test, "<b style='color:green;'>Successsfully clicked History Button to view the updated records for AOI</b>", screenshotName);
	            test.pass("<b style='color:green;'>AOI report generated</b>");
                WebElement clickCloseBtn = driver.findElement(By.xpath("//button[@class='btn-close']"));
                clickCloseBtn.click();   
	        } else 
	        {
                failScreenshot(test, "<b style='color:red;'>AOI report not generated or data not found for the given date range</b>", screenshotName);
	            test.fail("<b style='color:red;'> No AOI report available for scan status </b>"+ aoiScanStatus);
	            writeExtentSFCSReport(excelReportPath, "AOI Report", testName, "FAIL", timeStamp, "AOI report generation failed", "No data found in report for selected date range"+aoiScanStatus);
	        }
	        
		  	}
	} 

}
