package tests;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URL;
import java.util.concurrent.TimeUnit;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.testng.Assert;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Test;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.reporter.ExtentHtmlReporter;

import io.appium.java_client.windows.WindowsDriver;

public class CalculatorTest {

	private static WindowsDriver<WebElement> driver = null;
	private static WebElement calculatedResult = null;

	@BeforeClass
	private void setUp() throws Exception {

		//Setting up calculator
		DesiredCapabilities capabilities = new DesiredCapabilities();
		capabilities.setCapability("app", "Microsoft.WindowsCalculator_8wekyb3d8bbwe!App");
		driver = new WindowsDriver<WebElement>(new URL("http://127.0.0.1:4723"), capabilities);
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(2, TimeUnit.SECONDS);

		calculatedResult = driver.findElementByAccessibilityId("CalculatorResults");
		Assert.assertNotNull(calculatedResult);
	}

	@BeforeMethod
	public void Clear()	{
		//click 'çlear' button
		driver.findElementByName("Clear").click();
		Assert.assertEquals("0", getCalculatedResult());
	}

	protected String getCalculatedResult()	{
		return calculatedResult.getText().replace("Display is", "").trim();
	}

	protected void negate(String sign)	{
		if (sign.equalsIgnoreCase("-")) {
			driver.findElementByName("Positive Negative").click();			
		}
	}

	@AfterClass
	public static void TearDown() {
		//close application 
		calculatedResult = null;
		if (driver != null) {
			driver.quit();
		}
		driver = null;
	}

	protected String concatArray(String[] arr) {
		String str = "";
		for (String string : arr) {
			str += string + " ";
		}
		return str;
	}
	
	@Test
	public void calculation() throws IOException {

		ExtentHtmlReporter reporter = new ExtentHtmlReporter("./src/main/java/reporting/Calculator_Test_Extent_Report.html");
		ExtentReports extent = new ExtentReports();
		extent.attachReporter(reporter);
		ExtentTest test = extent.createTest("Windows Calculator Test", "Test Description");

		String filePath = "./src/main/java/CalculatorData.xlsx";

		FileInputStream input = new FileInputStream(filePath);
		XSSFWorkbook workbook = new XSSFWorkbook(input);		
		XSSFSheet sheet = workbook.getSheet("Sheet1");

		int rows = sheet.getLastRowNum();

		for (int i = 1; i <= rows; i++) {

			//get current row
			XSSFRow row = sheet.getRow(i);

			//retrieve data from excel and execute test
			String num1[] = row.getCell(1).getStringCellValue().split(" ");
			for (String num : num1) {
				driver.findElementByName(num).click();
			}
			String num1Sign = row.getCell(0).getStringCellValue();
			negate(num1Sign);
			String operand = row.getCell(2).getStringCellValue();
			driver.findElementByName(operand).click();
			String num2[] = row.getCell(4).getStringCellValue().split(" ");
			for (String num : num2) {
				driver.findElementByName(num).click();
			}
			String num2Sign = row.getCell(3).getStringCellValue();
			negate(num2Sign);
			driver.findElementByName("Equals").click();

			String screenshotPath = "A:/Work/eclipse-workspace/WindowsCalculatorTest/src/main/java/reporting/Maths_Addition_test_" + i + "_screenshot.png";

			File file = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
			FileUtils.copyFile(file, new File(screenshotPath));
			test.addScreenCaptureFromPath(screenshotPath, "screenshot");
			test.pass("Adding (" + num1Sign + ") " + concatArray(num1) + " and (" + num2Sign + ") " + concatArray(num2) + " results = " + getCalculatedResult());

			//Update the value of cell 
			Cell cell = row.getCell(5);
			if(cell == null){
				cell = row.createCell(5);
			}
			cell.setCellValue(getCalculatedResult());

			//Save the data to the spreadsheet
			FileOutputStream output = new FileOutputStream(new File(filePath));
			workbook.write(output);
			output.close();

			System.out.println("(" + num1Sign + ") " + concatArray(num1) + " " + operand + " (" + num2Sign + ") " + concatArray(num2) + " = "+ getCalculatedResult());
		}
		input.close();
		workbook.close();
		extent.flush();
	}

	/*@Test
	public void calculate() throws FilloException {

		Fillo fillo = new Fillo();
		Connection con = fillo.getConnection("./src/main/java/CalculatorData.xlsx");
		Recordset recSet = con.executeQuery("SELECT * FROM Sheet1");

		while (recSet.next()) {

			String num1 = recSet.getField("integer_1");
			driver.findElementByName(num1).click();

			String num1Sign = recSet.getField("sign_1");
			negate(num1Sign);

			String operand = recSet.getField("operator");
			driver.findElementByName(operand).click();

			String num2 = recSet.getField("integer_2");
			driver.findElementByName(num2).click();

			String num2Sign = recSet.getField("sign_2");
			negate(num2Sign);

			driver.findElementByName("Equals").click();

			System.out.println("(" + num1Sign + ") " + concatArray(num1) + " " + operand + " (" + num2Sign + ") " + concatArray(num2) + " = "+ getCalculatedResult());
			System.out.println("inserting to table...");

			con.executeUpdate("UPDATE Sheet1 SET Calculated_Result = '" + getCalculatedResult() + "' WHERE integer_1 = '" + num1
					+ "' AND integer_1 = '" + num1Sign + "' AND integer_2 = '" + num2 + "' AND sing_2 = '" + num2Sign + "'");

		}
		recSet.close();
		con.close();
	}*/
}
