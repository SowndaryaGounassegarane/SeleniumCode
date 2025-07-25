package com;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
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

import com.aventstack.extentreports.ExtentTest;

public class WorkOrderTest extends TestClass{
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
	public void testOrder() throws IOException, InterruptedException, TimeoutException
	{
		String path = System.getProperty("user.dir")+"/Data.xlsx";
		//String filePath = System.getProperty("user.dir") + "/" + ConfigClass.getProperty("INPUT_FILE");
		File file = new File(path);
		if(!file.exists())
		{
			System.out.println("Create the file:"+file);
			
		}
		FileInputStream stream = new FileInputStream(file);
		XSSFWorkbook workbook = new XSSFWorkbook(stream);
	 //   int sheetIndex = Integer.parseInt(ConfigClass.getProperty("WORK_ORDER_SHEET"));
		//XSSFSheet sheet = workbook.getSheetAt(sheetIndex);
	    XSSFSheet sheet =workbook.getSheetAt(2);
		for(int i=0;i<sheet.getPhysicalNumberOfRows();i++)
		{
		try 
		{
			XSSFRow row =sheet.getRow(i);
			if(row==null||row.getCell(0)==null||getCellValueAsString(row.getCell(0)).trim().isEmpty()) {
				continue;
			}
			
			String area = getCellValueAsString(row.getCell(0));
			String line =getCellValueAsString(row.getCell(1)).trim();
			String prodLevel =getCellValueAsString(row.getCell(2));
			DataFormatter formatter = new DataFormatter();
			String poNumber = formatter.formatCellValue(row.getCell(3)); 
			String prodCode=getCellValueAsString(row.getCell(4));
			String quantity = formatter.formatCellValue(row.getCell(5)); 
			String priority =getCellValueAsString(row.getCell(6)).trim();
			String expectedType = getCellValueAsString(row.getCell(7));
			String timestamp = new SimpleDateFormat("yyyyMMdd_HHmmss_SSS").format(new Date());
			String screenshotName = expectedType+"Work Order"+timestamp;
			test = extent.createTest("Work Order Creation -"+expectedType.toUpperCase());
			String excelReportPath = System.getProperty("user.dir")+"/report/Work_Order_Report.xlsx";
			//String excelReportPath = System.getProperty("user.dir") + "/" + ConfigClass.getProperty("SFCS_REPORT");
			String timeStamp = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(new Date());
			String testName = "Work Order -"+expectedType.toUpperCase();
		    test.assignCategory("Work Order");
	        test.assignAuthor("Automation Testing");
	        test.assignDevice("Chrome");
			driver.get(ConfigClass.getProperty("BASEURL")+"/work-orders");
			driver.manage().window().maximize();
			//test.info("<span style='color:black'><b>Navigated to the Work Order Screen</b>");
			WebDriverWait wait = new WebDriverWait(driver,Duration.ofSeconds(10));
			Thread.sleep(200);
            WebElement icon = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("(//a[@class='nav-link text-secondary'])[1]")));
	        icon.click();
	        //test.info("<span style='color:black'><b>Click icon button</b>");
	        Thread.sleep(200);
            WebElement clickWorkOrder = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//p[text()='Work Orders']")));
	        clickWorkOrder.click();
	        test.info("<span style='color:black'><b>Click Work Order</b>");
	        WebElement clickAddWorkOrder = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//button[contains(text(), 'Add Work Order')]")));
	        JavascriptExecutor js = (JavascriptExecutor)driver;
	        js.executeScript("arguments[0].click();", clickAddWorkOrder);
	        test.info("<span style='color:black'><b>Click Add Work Order</b>");
            Thread.sleep(200);
	        WebElement workOrderNumber = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[@class='form-control bg-light']")));
	        String number = workOrderNumber.getAttribute("value");
	        System.out.println(number);
	        test.info("<span style='color:black'><b>Generated Work Order Number :<span style='color:blue'><b>" +number+"</b></span>");
	        WebElement clickAreaDropdown = driver.findElement(By.cssSelector(".css-1yp4zep-control"));
	        clickAreaDropdown.click();
	        test.info("<span style='color:black'><b>Click Area Dropdown </b>");
	        Thread.sleep(500);
	        By areaOption = By.xpath("//div[contains(@class, '-menu')]//div[contains(@class, '-option') and normalize-space()='" + area + "']");
	        WebElement selectArea = wait.until(ExpectedConditions.visibilityOfElementLocated(areaOption));
	        Actions actions = new Actions(driver);
	        actions.moveToElement(selectArea).perform();
	        selectArea.click();
	        test.info("<span style='color:black'><b> Select Area :<span style='color:blue'><b>" +area+"</b></span>");
	        System.out.println(area);
	        WebElement clickLineDropdown =driver.findElement(By.cssSelector(".css-1yp4zep-control"));
	        clickLineDropdown.click();
	        Thread.sleep(500);
	        By lineOption = By.xpath("//div[contains(@class,'-menu')]//div[contains(@class,'-option') and normalize-space()='"+line+"']");
	        WebElement selectLine = wait.until(ExpectedConditions.presenceOfElementLocated(lineOption));
	        actions.moveToElement(selectLine).perform();
	        selectLine.click();
	        test.info("<span style='color:black'><b>Select Line : <span style='color:blue'><b>" +line+"</b></span>");
	        System.out.println(line);
	        WebElement extractLineCode = driver.findElement(By.xpath("(//input[@class='form-control'])[3]"));
	        String getLineCode =extractLineCode.getAttribute("value");
	        test.info("<span style='color:black'><b>Line Code : <span style='color:blue'><b>" +getLineCode+"</b></span>");
	        System.out.println(getLineCode);
	        Thread.sleep(100);
	        WebElement clickProductLevel = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//label[contains(text(), 'Product Level')]/following::div[contains(@class,'control')][1]")));
	        js.executeScript("arguments[0].scrollIntoView(true);", clickProductLevel);
	        Thread.sleep(200);
            clickProductLevel.click();
            actions.moveToElement(clickProductLevel).click().sendKeys(Keys.ARROW_DOWN).perform();
	        Thread.sleep(100); 
	        prodLevel = prodLevel.trim();
            By productLevelOption = By.xpath("//div[contains(@class,'-menu')]//div[contains(@class,'-option') and normalize-space(text())='" + prodLevel + "']");
            WebElement selectProdLevel = wait.until(ExpectedConditions.elementToBeClickable(productLevelOption));
	        selectProdLevel.click(); 
	        test.info("<span style='color:black'><b>Product Level : <span style='color:blue'><b>" +prodLevel+"</b></span>");
	        System.out.println( prodLevel);
            WebElement enterPoNumber = driver.findElement(By.xpath("(//input[@class='form-control '])[1]"));
            enterPoNumber.sendKeys(poNumber);
            test.info("<span style='color:black'><b>Po Number : <span style='color:blue'><b>" +poNumber+"</b></span>");
            System.out.println(poNumber);
            WebElement clickProductCode =wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//label[contains(text(), 'Product Code')]/following::div[contains(@class,'control')][1]")));
            js.executeScript("arguments[0].scrollIntoView(true);", clickProductCode);
            Thread.sleep(200);
            clickProductCode.click();
            actions.moveToElement(clickProductCode).click().sendKeys(Keys.ARROW_DOWN).perform();
            Thread.sleep(100);
            By productCodeOpt = By.xpath("//div[contains(@class,'-menu')]//div[contains(@class,'-option')and normalize-space(text())='"+prodCode+"']");
            WebElement selectProductCode = wait.until(ExpectedConditions.elementToBeClickable(productCodeOpt));
            selectProductCode.click();
            test.info("<span style='color:black'><b>Product Code : <span style='color:blue'><b>" +prodCode+"</b></span>");
            System.out.println(prodCode);
            WebElement extractProductDescp =driver.findElement(By.xpath("(//input[@class='form-control'])[4]"));
            String getProductDescp=extractProductDescp.getAttribute("value");
            test.info("<span style='color:black'><b>Product Description : <span style='color:blue'><b>" +getProductDescp+"</b></span>");
            System.out.println(getProductDescp);
            WebElement enterQuantity = driver.findElement(By.xpath("(//input[@class='form-control '])[2]"));
            enterQuantity.sendKeys(quantity);
            test.info("<span style='color:black'><b>Quantity : <span style='color:blue'><b>" +quantity+"</b></span>");
            System.out.println(quantity);
            WebElement clickPriority = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//label[contains(text(), 'Priority')]/following::div[contains(@class,'control')][1]")));
            js.executeScript("arguments[0].scrollIntoView(true);", clickPriority);
            Thread.sleep(500);
            clickPriority.click();
            By dropdownPriority = By.xpath("//div[contains(@class,'-menu')]");
            wait.until(ExpectedConditions.presenceOfElementLocated(dropdownPriority));
            priority = priority.trim();  
            Thread.sleep(100);
            actions.sendKeys(Keys.ARROW_DOWN).perform();
            Thread.sleep(300);  
            priority = priority.trim();
            By optionPriority = By.xpath("//div[contains(@class,'-menu')]//div[contains(@class,'-option') and normalize-space(text())='" + priority + "']");
            WebElement optPriority = wait.until(ExpectedConditions.elementToBeClickable(optionPriority));
            Thread.sleep(100);
            optPriority.click();  
            test.info("<span style='color:black'><b>Priority :<span style='color:blue'><b> " + priority+"</b></span>");
            System.out.println(priority);
            WebElement extractWorkOrderDate = driver.findElement(By.xpath("//input[@type='date']"));
            String workOrderDate =extractWorkOrderDate.getAttribute("value");
            test.info("<span style='color:black'><b>Work Order Creation Date : <span style='color:blue'><b>" +workOrderDate+"</b></span>");
            System.out.println(workOrderDate);
            WebElement clickCreateBtn = driver.findElement(By.xpath("(//button[@type='submit'])[2]"));
            clickCreateBtn.click();
            test.info("<span style='color:black'><b>Click Create button</b>");
            Thread.sleep(4000);
            try
            {
                By successMsg = By.xpath("//div[contains(text(),'Work Order created')]");
                WebElement msg = wait.until(ExpectedConditions.visibilityOfElementLocated(successMsg));
            	String workOrderMsg = msg.getText();
                System.out.println(workOrderMsg);
                test.info("<span style='color:black'><b>Success Message :<span style='color:green'><b>" + workOrderMsg+"</b></span>");
                passScreenshot(test,"<b style='color:green;'> Work Order created success</b>",screenshotName);
                test.pass("<b style='color:green;'> Work Order created successfully</b>");
                writeExtentSFCSReport(excelReportPath,"Work Order Report",testName,"PASS",timeStamp,"Creation of Work Order","Unique Work Order number is generated -" + number +" - Work Order Created Successfully");
                String rowXPath = "//div[@role='row' and .//div[contains(text(),'" + number + "')]]";
                WebElement targetRow = wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath(rowXPath)));
                js.executeScript("arguments[0].scrollIntoView({block:'center'});", targetRow);
                Thread.sleep(500);
                js.executeScript("document.querySelector('.ag-body-horizontal-scroll-viewport').scrollLeft = 3000;");
                Thread.sleep(100); 
                WebElement activateBtn = targetRow.findElement(By.xpath(".//button[contains(text(),'Activate')]"));
                js.executeScript("arguments[0].scrollIntoView({block:'center'});", activateBtn);
                wait.until(ExpectedConditions.elementToBeClickable(activateBtn)).click();
                System.out.println("Successfully clicked Activate for: " + number);
                driver.switchTo().alert().accept();

            }
			catch (Exception e)
      		{
      	        test.fail("Failed at row " + i + ": " + e.getMessage());
      		    failScreenshot(test, "<b style='color:red;'>Work order creation failed</b>",screenshotName);
      		    writeExtentSFCSReport(excelReportPath,"Work Order Report",testName,"FAIL",timeStamp,"Creation of Work order failed","Unique Work Order number is not generated - " + number + " - Work Order Not created");
      			System.err.println("Error at row " + i + ": " + e.getMessage());
      			e.printStackTrace();
      		}
           
		}        
		catch (Exception e) 
		{
	        e.getMessage();
	    } 
	   }
	} 

}
