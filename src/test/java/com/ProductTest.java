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
import org.openqa.selenium.TimeoutException;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

import com.aventstack.extentreports.ExtentTest;

public class ProductTest extends TestClass {
	private static ExtentTest test;
	@SuppressWarnings({ })

	public static String getCellValueAsString(Cell cell) {
	    if (cell == null) 
	    {
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

	
	@Test
	public void prodScreen() throws IOException, InterruptedException, TimeoutException
	{
		
		    //String path = System.getProperty("user.dir") + "/Data.xlsx";
		String filePath = System.getProperty("user.dir") + "/" + ConfigClass.getProperty("INPUT_FILE");
		    File file = new File(filePath);
		    if (!file.exists()) {
		        System.out.println("File not found: " + file);
		        return;
		    }

		    FileInputStream stream = new FileInputStream(file);
		    XSSFWorkbook workbook = new XSSFWorkbook(stream);
		    int sheetIndex = Integer.parseInt(ConfigClass.getProperty("PRODUCT_SHEET"));
			XSSFSheet sheet = workbook.getSheetAt(sheetIndex);
		    //XSSFSheet sheet = workbook.getSheetAt(1);

		    for (int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
		        try {
		            XSSFRow row = sheet.getRow(i);
		            String prodName = getCellValueAsString(row.getCell(0));
		            String prodCode = getCellValueAsString(row.getCell(1));
		            double targHr = Double.parseDouble(getCellValueAsString(row.getCell(2)));
		            double shiftHr = Double.parseDouble(getCellValueAsString(row.getCell(3)));
		            double Shift = Double.parseDouble(getCellValueAsString(row.getCell(4)));
		            String processStage = getCellValueAsString(row.getCell(5));
		            String smtSide = getCellValueAsString(row.getCell(6)); 
		            String prodDescp = getCellValueAsString(row.getCell(7));
		            String status = getCellValueAsString(row.getCell(8));
		            String expectedType = getCellValueAsString(row.getCell(9));
		            String timeStamp = new SimpleDateFormat("yyyyMMdd_HHmmss_SSS").format(new Date());
		            String screenshotName = expectedType+"Product"+timeStamp;
		            test = extent.createTest("Product Test -"+expectedType.toUpperCase());
		            //String excelReportPath = System.getProperty("user.dir")+"/report/ProductDB_Report.xlsx";
		            String excelReportPath = System.getProperty("user.dir") + "/" + ConfigClass.getProperty("SFCS_REPORT");
		            String timestamp = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(new Date());
		            String testName = "Product Test - "+expectedType.toUpperCase();
		            try {
		    			//Connection con = DriverManager.getConnection("jdbc:mysql://10.1.71.113:3306/qa_sfcs?useSSL=false&allowPublicKeyRetrieval=true","root", "Vepl@123");
		            	Connection con = DriverManager.getConnection( ConfigClass.getProperty("DB_URL"), ConfigClass.getProperty("DB_USERNAME"), ConfigClass.getProperty("DB_PASSWORD"));
		    			if(con.isClosed())
		    			{
		    				System.out.println("Database is not connected");
		    				
		    			}
		    			else
		    			{
		    				System.out.println("Database is connected");
		    			}
		    			 String deleteExistProductCode = prodCode; 
		                 String deleteQuery = "DELETE FROM products WHERE product_code = ?";
		                 PreparedStatement prepStat = con.prepareStatement(deleteQuery);
		                 prepStat.setString(1, deleteExistProductCode);
		                 int deleteDuplicateRows = prepStat.executeUpdate();

		                 if (deleteDuplicateRows > 0) 
		                 {
		                     System.out.println("Existing Product code deleted Successfully -"+ deleteExistProductCode);
		                 } else
		                 {
		                     System.out.println("Product code not found in DB -"+ deleteExistProductCode);
		                 }

		                 prepStat.close();
		                 con.close();
		             
                         Thread.sleep(1000);
		         } 
		            catch (SQLException e) 
		            {
		             e.printStackTrace();
		         }
		            test.assignCategory("Product");
		            test.assignAuthor("Automation Testing");
		            test.assignDevice("Chrome");
		            driver.get(ConfigClass.getProperty("BASEURL") + "/products");
		            driver.manage().window().maximize();
		          //  test.info("<span style='color:black'><b>Navigated to the Product screen</b>");
		            WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(15));
                    WebElement icon = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("(//a[@class='nav-link text-secondary'])[1]")));
		            icon.click();
		            //test.info(" <span style='color:black'><b> Click Icon button</b>");
                    WebElement clickProduct = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//p[text()='Products']")));
		            clickProduct.click();
		            test.info("<span style='color:black'><b>Click Product</b>");
		            Thread.sleep(500);
                    WebElement clickAddProduct = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//button[text()=' Add Product']")));
		            JavascriptExecutor js =(JavascriptExecutor) driver;
		            js.executeScript("arguments[0].click()", clickAddProduct);
		            test.info(" <span style='color:black'><b>Click Add Product</b>");
                    Thread.sleep(500);
		            WebElement enterProductName = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[@placeholder='Enter product name']")));
		            enterProductName.sendKeys(prodName);
		            test.info("<span style='color:black'><b> ProductName :  <span style='color:blue'><b> " + prodName+"</b></span>");
		            Thread.sleep(500);
		            WebElement enterProductCode = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[@placeholder='Enter product code']")));
		            enterProductCode.sendKeys(prodCode);
		            test.info("<span style='color:black'><b>ProductCode : <span style='color:blue'><b> " + prodCode+"</b></span>");
		            Thread.sleep(500);
		            WebElement enterTargetHourly = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("(//input[@class='form-control '])[3]")));
		            enterTargetHourly.sendKeys(String.valueOf(targHr));
		            test.info("<span style='color:black'><b>TargetHourly : <span style='color:blue'><b>  " + targHr+"</b></span>");
		            Thread.sleep(500);
		            WebElement enterShiftHours = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("(//input[@class='form-control '])[4]")));
		            enterShiftHours.sendKeys(String.valueOf(shiftHr));
		            test.info("<span style='color:black'><b>ShiftHour : <span style='color:blue'><b> " + shiftHr+"</b></span>");
		            Thread.sleep(500);
		            WebElement enterNoOfShifts = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("(//input[@class='form-control '])[5]")));
		            enterNoOfShifts.sendKeys(String.valueOf(Shift));
		            test.info("<span style='color:black'><b>NoOfShifts : <span style='color:blue'><b> " + Shift+"</b></span>");
		            WebElement dailyTarget=wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[@class='form-control bg-light']")));
		            String dalyTarget =dailyTarget.getAttribute("value");
		            test.info("<span style='color:black'><b>DailyTarget : <span style='color:blue'><b> " +dalyTarget+"</b></span>");
		            Thread.sleep(500);
		            String[] selectProcessStage = processStage.split(",");
		            for (String stage : selectProcessStage)
		            {
		                stage = stage.trim(); 
		                String checkboxXpath = String.format("//label[contains(normalize-space(),'%s')]", stage);
		                WebElement label = driver.findElement(By.xpath(checkboxXpath));
		                WebElement checkbox = label.findElement(By.xpath("//input[@type='checkbox']"));
		                if (!checkbox.isSelected()) 
		                {
		                   label.click(); 
		                   test.info("<span style='color:black'><b>Process Stage :<span style='color:blue'><b> " +stage+"</b></span>");
		                }
		            }
		         boolean isSMTForProcessStage = processStage!=null && processStage.contains("SMT");
		         if(isSMTForProcessStage && smtSide != null && !smtSide.trim().isEmpty())
		         {
		        	 try
		        	 {
		        	 WebElement smtRadio = driver.findElement(By.xpath("//input[@type='radio' and @value='" + smtSide + "']"));
			         smtRadio.click();
			         test.info("<span style='color:black'><b>SMT Side : <span style='color:blue'><b> " + smtSide+"</b></span>");
				    } 
				    catch (Exception e)
				    {
				        test.warning("<span style='color:black'><b>SMT side radio button not found for this process stage: <span style='color:blue'><b> " + smtSide);
				    }
				} 
		         else 
				{
				    test.info("<span style='color:red'><b> SMT Side not available for this process stage");
				}
		         		            
		            Thread.sleep(500);
		            WebElement enterProductDescription = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//textarea[@class='form-control']")));
		            enterProductDescription.sendKeys(prodDescp);
		            test.info("<span style='color:black'><b>Product Description : <span style='color:blue'><b> " + prodDescp+"</b></span>");
		            Thread.sleep(500);
		            WebElement selectStatus = driver.findElement(By.xpath("//select[@class='form-select']"));
		            Select selectTheStatus = new Select(selectStatus);
		            selectTheStatus.selectByVisibleText(status);
		            test.info("<span style='color:black'><b>Status :<span style='color:blue'><b>  " + status+"</b></span>");
		            Thread.sleep(500);
		            WebElement clickCreateButton = driver.findElement(By.xpath("//button[text()='Create']"));
		            clickCreateButton.click();
		            test.info("<span style='color:black'><b>Click Create</b>");
		            try 
	                {
	                    WebElement productSuccessMsg = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//div[contains(text(), 'Product created successfully') or contains(text(), 'already exists')]")));
	                    String successMsg = productSuccessMsg.getText();
	                    System.out.println(successMsg);
	                  if (successMsg.contains("Product created successfully")) 
	                    {
	                        test.pass("<b style='color:green;'> Product created successfully</b>");
	                        passScreenshot(test, "<b style='color:green;'> Product created </b>", screenshotName);
	                       writeExtentSFCSReport(excelReportPath, "Product Report", testName, "PASS", timestamp, "Valid Product Name and Product Code", "Product created successfully:" +prodName +" And "+prodCode+" - Product Created");
	                        System.out.println("Product created : " + prodName + ", " + prodCode);
	                    } 
	                  else if (successMsg.contains("product code already exists")) 
	                    {
	                        test.fail("<b style='color:red;'> Product code already exist </b> <span style='color:blue'><b>- Failed in creating products");
	                        failScreenshot(test, "Product code already exists", screenshotName);
	                        writeExtentSFCSReport(excelReportPath, "Product Report", testName, "FAIL", timestamp, "Duplicate Product Code", "Product code exist:" +prodName +" And "+prodCode+" - Product Code already exists");
	                        System.err.println("Product already exists : " + prodName + ", " + prodCode);
	                        System.out.println("Product code already exist for :"+prodName+","+prodCode);                  
	                        }
	                  else
	                        {
	                        test.fail(" <b style='color:red;'>Fail Message :<b style='color:red;'> " + successMsg);
	                        failScreenshot(test, "Product creation failed", screenshotName);
	                    }

	                } 
		            catch (Exception e)
		            {
	                    test.fail("<b style='color:red;'> Exception occured during product creation : <b style='color:red;'>" + e.getMessage());
	                    failScreenshot(test, "Product code already exist failed", screenshotName);
	                    writeExtentSFCSReport(excelReportPath, "Product Report", testName, "FAIL", timeStamp, "Duplicate Product Code", "Product code exist:" +prodName +" And "+prodCode+" - Product Code already exists");
	                    e.printStackTrace();
	                }

	            } 
		        catch (Exception e)
		        {
	                test.fail("Exception processing row : " + e.getMessage());
	                e.printStackTrace();
	            }
	        }

	        workbook.close();
	        stream.close();

} 

}
