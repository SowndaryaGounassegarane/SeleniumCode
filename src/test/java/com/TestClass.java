package com;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.BeforeSuite;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.MediaEntityBuilder;
import com.aventstack.extentreports.reporter.ExtentSparkReporter;

import io.github.bonigarcia.wdm.WebDriverManager;

public class TestClass {
	protected static WebDriver driver;
	    public static ExtentReports extent;
	   public static ExtentTest test;

	    @BeforeSuite
	    public static void setUp() {
	    	 if (extent == null)
		        {
		        	 String path = System.getProperty("user.dir")+"\\report\\SFCSReports.html";
		             
		             ExtentSparkReporter e = new ExtentSparkReporter(path);
		             e.config().setReportName("SFCS Automation");
		             e.config().setDocumentTitle("SFCS Report");
		             extent = new ExtentReports();
		             extent.attachReporter(e);
		             extent.setSystemInfo("Sowndarya", "Automation Tester");
		             
	        }

	        if (driver == null) {
	            WebDriverManager.chromedriver().setup();
	            driver = new ChromeDriver();
	        }
	    }

	    @AfterSuite
	    public static void tearDown() {
	        try {
//	            System.out.println("Press Enter to close the browser...");
//	            new java.util.Scanner(System.in).nextLine(); 
	        } catch (Exception ignored) {}

//	        if (driver != null) {
//	            driver.quit();
//	            driver = null;
//	        }

	        if (extent != null) {
	            extent.flush();
	        }
	    }

	    private static final Set<String> clearedSheets = new HashSet<>();

	    public static void writeExtentSFCSReport(String path, String sheetName, String testName, String status,String timeStamp, String step, String details) throws IOException {
	        File file = new File(path);
	        XSSFWorkbook workbook;
	        XSSFSheet sheet;

	        if (file.exists() && file.length() > 0) {
	            try (FileInputStream stream = new FileInputStream(file)) {
	                workbook = new XSSFWorkbook(stream);
	            } catch (Exception e) {
	                System.out.println("Warning: Excel file is corrupted or unreadable. Creating a new workbook.");
	                workbook = new XSSFWorkbook();
	            }
	        } else {
	            workbook = new XSSFWorkbook();
	        }

	        if (!clearedSheets.contains(sheetName))
	        {
	            int index = workbook.getSheetIndex(sheetName);
	            if (index != -1) 
	            {
	                workbook.removeSheetAt(index);
	            }
	            sheet = workbook.createSheet(sheetName);
	            createHeaderAndSummary(sheet);
	            clearedSheets.add(sheetName);
	        } else {
	            sheet = workbook.getSheet(sheetName);
	            if (sheet == null) 
	            {
	                sheet = workbook.createSheet(sheetName);
	                createHeaderAndSummary(sheet);
	                clearedSheets.add(sheetName);
	            }
	        }

	        XSSFFont boldFont = workbook.createFont();
	        boldFont.setBold(true);
	        XSSFCellStyle passStyle = workbook.createCellStyle();
	        passStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
	        passStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	        XSSFCellStyle failStyle = workbook.createCellStyle();
	        failStyle.setFillForegroundColor(IndexedColors.ROSE.getIndex());
	        failStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	        XSSFCellStyle boldStyle = workbook.createCellStyle();
	        boldStyle.setFont(boldFont);
	        int insertRow = sheet.getLastRowNum() + 1;
	        if (insertRow < 2) insertRow = 2; // Skip headers
	        XSSFRow row = sheet.createRow(insertRow);
	        row.createCell(0).setCellValue(testName);
	        XSSFCell statusCell = row.createCell(1);
	        statusCell.setCellValue(status);
	        statusCell.setCellStyle(status.equalsIgnoreCase("PASS") ? passStyle : failStyle);
	        row.createCell(2).setCellValue(timeStamp);
	        row.createCell(3).setCellValue(step);
	        XSSFCell detailCell = row.createCell(4);
	        detailCell.setCellValue(details);
	        detailCell.setCellStyle(boldStyle);
	        int totalTests = 0, totalPass = 0, totalFail = 0;
	        for (int i = 2; i <= sheet.getLastRowNum(); i++) {
	            XSSFRow r = sheet.getRow(i);
	            if (r != null) 
	            {
	                XSSFCell sCell = r.getCell(1);
	                if (sCell != null && sCell.getCellType() == CellType.STRING) {
	                    String val = sCell.getStringCellValue().trim();
	                    if (val.equalsIgnoreCase("PASS")) totalPass++;
	                    else if (val.equalsIgnoreCase("FAIL")) totalFail++;
	                    totalTests++;
	                }
	            }
	        }

	        double passPct = totalTests > 0 ? (totalPass * 100.0 / totalTests) : 0.0;
	        double failPct = totalTests > 0 ? (totalFail * 100.0 / totalTests) : 0.0;
	        XSSFCellStyle totalStyle = workbook.createCellStyle();
	        totalStyle.setFillForegroundColor(IndexedColors.LIGHT_TURQUOISE.getIndex());
	        totalStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	        totalStyle.setFont(boldFont);
	        XSSFCellStyle passStyleSum = workbook.createCellStyle();
	        passStyleSum.setFillForegroundColor(IndexedColors.BRIGHT_GREEN.getIndex());
	        passStyleSum.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	        passStyleSum.setFont(boldFont);
	        XSSFCellStyle failStyleSum = workbook.createCellStyle();
	        failStyleSum.setFillForegroundColor(IndexedColors.CORAL.getIndex());
	        failStyleSum.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	        failStyleSum.setFont(boldFont);
	        XSSFCellStyle pctStyle = workbook.createCellStyle();
	        pctStyle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
	        pctStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	        pctStyle.setFont(boldFont);
	        XSSFRow summaryRow = sheet.getRow(1);
	        if (summaryRow == null) summaryRow = sheet.createRow(1);
	        XSSFCell c0 = summaryRow.createCell(0);
	        c0.setCellValue("Total Tests: " + totalTests);
	        c0.setCellStyle(totalStyle);
	        XSSFCell c1 = summaryRow.createCell(1);
	        c1.setCellValue("Total Pass: " + totalPass);
	        c1.setCellStyle(passStyleSum);
	        XSSFCell c2 = summaryRow.createCell(2);
	        c2.setCellValue("Total Fail: " + totalFail);
	        c2.setCellStyle(failStyleSum);
	        XSSFCell c3 = summaryRow.createCell(3);
	        c3.setCellValue("Pass %: " + String.format("%.2f", passPct));
	        c3.setCellStyle(pctStyle);
	        XSSFCell c4 = summaryRow.createCell(4);
	        c4.setCellValue("Fail %: " + String.format("%.2f", failPct));
	        c4.setCellStyle(pctStyle);
	        for (int i = 0; i <= 4; i++) {
	            sheet.autoSizeColumn(i);
	        }
	        try (FileOutputStream fos = new FileOutputStream(file)) {
	            workbook.write(fos);
	        }
	        workbook.close();
	    }
	    private static void createHeaderAndSummary(XSSFSheet sheet) {
	        XSSFRow header = sheet.createRow(0);
	        header.createCell(0).setCellValue("Test Name");
	        header.createCell(1).setCellValue("Status");
	        header.createCell(2).setCellValue("TimeStamp");
	        header.createCell(3).setCellValue("Step");
	        header.createCell(4).setCellValue("Details");
	        XSSFRow summary = sheet.createRow(1);
	        summary.createCell(0).setCellValue("Total Tests: 0");
	        summary.createCell(1).setCellValue("Total Pass: 0");
	        summary.createCell(2).setCellValue("Total Fail: 0");
	        summary.createCell(3).setCellValue("Pass %: 0.00");
	        summary.createCell(4).setCellValue("Fail %: 0.00");
	    }
	   
//protected static String takeScreenshot(String folder,String screenshotName) {
// try
// {
//     TakesScreenshot ts = (TakesScreenshot) driver;
//     File src = ts.getScreenshotAs(OutputType.FILE);
//     String path = System.getProperty("user.dir") + "/report/screenshots/" + screenshotName + ".png";
//     File dest = new File(path);  
//     dest.getParentFile().mkdirs();
//     FileUtils.copyFile(src, dest);
//     return path;
// }
// catch (IOException e)
// {
//     e.printStackTrace();
//     return null;
// }
//}
//public static String passScreenshot(ExtentTest test, String message, String screenshotName) throws IOException {
// String path = takeScreenshot("pass",screenshotName);
// if (path != null) 
// {
//     test.pass(message, MediaEntityBuilder.createScreenCaptureFromPath(path).build());
// } 
// else
// {
//     test.warning("Screenshot not available: " + message);
// }
//	return path;
//}
//
//
//public static String failScreenshot(ExtentTest test, String message, String screenshotName) throws IOException 
//{
// String path = takeScreenshot("fail",screenshotName);
// if (path != null) 
// {
// 	 
//     test.fail(message, MediaEntityBuilder.createScreenCaptureFromPath(path).build());
// } 
// else 
// {
//     test.warning("Screenshot not available for failure: " + message);
// }
//	return path;
//}
	    public static void setDriver(WebDriver webDriver) {
	        driver = webDriver;
	    }

	    public static String getBase64Screenshot() {
	        return ((TakesScreenshot) driver).getScreenshotAs(OutputType.BASE64);
	    }

	    public static void passScreenshot(ExtentTest test, String message, String screenshotName) {
	        String base64 = getBase64Screenshot();
	        test.pass(message, MediaEntityBuilder
	                .createScreenCaptureFromBase64String(base64, screenshotName).build());
	    }

	    public static void failScreenshot(ExtentTest test, String message, String screenshotName) {
	        String base64 = getBase64Screenshot();
	        test.fail(message, MediaEntityBuilder
	                .createScreenCaptureFromBase64String(base64, screenshotName).build());
	    }
	
public void selectWorkOrderFromAlias(String workOrderNumber) {
 String area = "";

 switch (workOrderNumber.toUpperCase()) {
     case "WO1":
         area = "SMT";
         break;
     case "WO2":
         area = "WAVE";
         break;
     default:
         test.fail("Unknown Work Order : " + workOrderNumber);
         return;
 }

 try
 {
 	WebDriverWait wait = new WebDriverWait(driver,Duration.ofSeconds(10));
     WebElement dropdown = wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector(".css-13cymwt-control")));
     dropdown.click();
     List<WebElement> options = wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.xpath("//div[contains(@class,'-option')]")));
     for (WebElement option : options) 
     {
         String optionText = option.getText().trim();
         if (optionText.toUpperCase().contains(area))
         {
             option.click();
             test.info("Selected Work Order for " + workOrderNumber + ": <b style='color:blue;'>" + optionText + "</b>");
             return;
         }
     }

     test.fail("No Work Order found for Area: " + area);
 } catch (Exception e) {
     test.fail("Exception while selecting Work Order for " + workOrderNumber + ": " + e.getMessage());
 }
}


}




