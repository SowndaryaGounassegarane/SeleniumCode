package com;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//import org.junit.Test;
import org.openqa.selenium.By;
import org.openqa.selenium.TimeoutException;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.Test;

import com.aventstack.extentreports.ExtentTest;

@SuppressWarnings("unused")
public class LoginTest extends TestClass{
	//public static  ExtentTest test;
	 public static String getCellValueAsString(Cell cell) {
	        if (cell == null) {
	            return "";
	        }

	        switch (cell.getCellType()) {
	            case STRING:
	                return cell.getStringCellValue().trim();
	            case NUMERIC:
	                double d = cell.getNumericCellValue();
	                if (d == Math.floor(d)) {
	                    return String.valueOf((long) d);
	                } 
	                else 
	                {
	                    return String.valueOf(d);
	                }
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
	  public void loginScreen() throws IOException, InterruptedException
		{

	       // String filePath = System.getProperty("user.dir") + "/Data.xlsx";
		   String filePath = System.getProperty("user.dir") + "/" + ConfigClass.getProperty("INPUT_FILE");
	        File file = new File(filePath);
	        if (!file.exists())
	        {
	            System.out.println("file not found create the file:" + file);
	            return;
	        }
	        FileInputStream stream = new FileInputStream(file);
	        XSSFWorkbook workbook = new XSSFWorkbook(stream);
	        int sheetIndex = Integer.parseInt(ConfigClass.getProperty("LOGIN_SHEET"));
			XSSFSheet sheet = workbook.getSheetAt(sheetIndex);
	       // XSSFSheet sheet = workbook.getSheetAt(0);
			for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++)
			{
			    XSSFRow row = sheet.getRow(i);
			    String userName = getCellValueAsString(row.getCell(0)).trim();
		        String password = getCellValueAsString(row.getCell(1)).trim();
		        String expectedType = getCellValueAsString(row.getCell(2));
		        String timestamp = new SimpleDateFormat("yyyyMMdd_HHmmss_SSS").format(new Date());
		        String screenshotName = expectedType + "Login" + timestamp;
			    test = extent.createTest("Login Test - " + expectedType.toUpperCase());
				
			   // String excelReportPath = System.getProperty("user.dir") + "/report/Login_Report.xlsx";
                String excelReportPath = System.getProperty("user.dir") + "/" + ConfigClass.getProperty("SFCS_REPORT");
			    String timeStamp = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(new Date());
			    String testName = "Login Test - " + expectedType.toUpperCase();
			    test.assignCategory("Login");
			    test.assignAuthor("Automation Testing");
			    test.assignDevice("Chrome");
                driver.get(ConfigClass.getProperty("BASEURL") + "/login");
                driver.manage().window().maximize();
			    test.info("<span style='color:black'><b>Navigate to Login screen</b>");           
			    WebElement user = driver.findElement(By.id("email"));
			    user.sendKeys(userName);
			    test.info("<span style='color:black'><b>Email: <span style='color:blue'><b> " + userName+"</b></span>");
			    WebElement pass = driver.findElement(By.id("password"));
			    pass.sendKeys(password);
			    test.info("<span style='color:black'><b>Password: <span style='color:blue'><b> " + password.replaceAll(".", "*"));
	            WebElement loginBtn = driver.findElement(By.xpath("//button[@type='submit']"));
			    loginBtn.click();
			    test.info("<span style='color:black'><b>Click Sign In button</b>");
			    try 
	            {
			        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
			        WebElement successMsg = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[text()='Signed in successfully!']")));
	                if (successMsg.isDisplayed()) 
	                {
			        	passScreenshot(test,"<b style='color:green;'> Login Success</b>",screenshotName);
			            test.pass("<b style='color:green;'>Signed in successfully!</b>" );
			            writeExtentSFCSReport(excelReportPath, "Login Report", testName, "PASS", timeStamp, "Valid Login", "Login success for Valid User Credentials:" +userName +" And "+password.replaceAll(".","*")+" - Signed in successfully");
			           
			        }
			    } 
	            catch (TimeoutException e) 
	            {
	            	test.fail("<b style='color:red;'>Login failed for user: </b> <span style='color:blue'><b>" + userName + " and "  + password.replaceAll(".", "*"));
	            	failScreenshot(test, "<b style='color:red;'>Login failed </b>", screenshotName);
	            	writeExtentSFCSReport(excelReportPath, "Login Report", testName, "FAIL", timeStamp, "Invalid Login", "Login failed for Invalid User Credentials: " + userName + " And "  + password.replaceAll(".", "*")+" - Error in Sign In for Invalid Credentials");
			    }
			    
			}
		         
}
		}

