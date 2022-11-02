package org.com;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Properties;
import java.util.Set;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.format.CellFormatType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;

import io.github.bonigarcia.wdm.WebDriverManager;

/**
 * 
 * @author Deepen
 * @description used to store all reusable methods
 * @date 29/08/2022
 */
public class BaseClass1 {

	protected static WebDriver driver;

	/**
	 * @description Used to set and launch browser based on given parameter
	 * @param browserType
	 */
	public static void getDriver(String browserType) {
		switch (browserType) {
		case "chrome":
			WebDriverManager.chromedriver().setup();
			driver = new ChromeDriver();
			break;
		case "firefox":
			WebDriverManager.firefoxdriver().setup();
			driver = new FirefoxDriver();
			break;
		case "ie":
			WebDriverManager.iedriver().setup();
			driver = new InternetExplorerDriver();
			break;
		case "edge":
			WebDriverManager.edgedriver().setup();
			driver = new EdgeDriver();
			break;
		default:
			break;
		}
	}

	/**
	 * @description Used to read property file
	 * @param key
	 * @return String
	 * @throws FileNotFoundException
	 * @throws IOException
	 */
	public static String getPropertyFileValue(String key) throws FileNotFoundException, IOException {
		Properties properties = new Properties();
		properties.load(new FileInputStream(System.getProperty("user.dir") + "//Config//Config.properties"));

		String value = (String) properties.get(key);
		return value;
	}

	/**
	 * @description Used to launch given url
	 * 
	 * @param url
	 */
	public static void enterUrl(String url) {
		driver.get(url);
	}

	/**
	 * @description Used to maximize the window
	 */
	public static void maximize() {
		driver.manage().window().maximize();
	}

	/**
	 * @description Used to find Webelement by using id
	 * 
	 * @param id
	 * @return WebElement
	 */
	public static WebElement findElementById(String id) {
		WebElement element = driver.findElement(By.id(id));
		return element;

	}

	/**
	 * @description Used to find Webelement by using name
	 * @param name
	 * @return WebElement
	 * 
	 */
	public static WebElement findElementByName(String name) {
		WebElement element = driver.findElement(By.name(name));
		return element;

	}

	/**
	 * @description Used to find webelement by using classname
	 * @param className
	 * @return WebElement
	 */
	public static WebElement findElementByClassName(String className) {
		WebElement element = driver.findElement(By.className(className));
		return element;
	}

	/**
	 * @description Used to find webelement by using xpath
	 * @param xpath
	 * @return WebElement
	 */
	public static WebElement findElementByXpath(String xpath) {
		WebElement element = driver.findElement(By.xpath(xpath));
		return element;
	}

	/**
	 * @description Used to get title from of current webpage
	 * @return String
	 */
	public static String getTitle() {
		String title = driver.getTitle();
		return title;
	}

	/**
	 * @description Used to get current url
	 * @return
	 */
	public static String geturl() {
		String url = driver.getCurrentUrl();
		return url;

	}

	/**
	 * @description Used to enter value in inputfields
	 * @param element
	 * @param data
	 */
	public static void elementSendKeys(WebElement element, String data) {
		element.sendKeys(data);
	}

	/**
	 * @description Used to enter value in inputfield and continue with enter
	 * @param element
	 * @param data
	 */
	public static void elementSendKeysWithEnter(WebElement element, String data) {
		element.sendKeys(data, Keys.ENTER);
	}

	/**
	 * @description Used to clear the text of that field
	 * @param element
	 */
	public static void elementClear(WebElement element) {
		element.clear();

	}

	/**
	 * @description Used to Click the element
	 * @param element
	 */
	public static void elementClick(WebElement element) {
		element.click();
	}

	/**
	 * @description Used to close all windows
	 */

	public static void closeAllWindow() {
		driver.quit();
	}

	/**
	 * @description Used to close current window
	 */
	public static void closeCurrentWindow() {
		driver.close();
	}

	/**
	 * @description Used to get text from a specific web element
	 * @param element
	 * @return String
	 */
	public static String getText(WebElement element) {
		String text = element.getText();
		return text;
	}

	/**
	 * @description Used to get given  atrributevalue from webelement by value
	 * @param element
	 * @return String
	 */
	public static String elementGetAttribute(WebElement element) {
		String attribute = element.getAttribute("value");
		return attribute;
	}

	/**
	 * @description Used to get given atrributevalue from webelement by name
	 * @param element
	 * @return String
	 */
	public static String getAttribute(WebElement element, String attributename) {
		String attribute = element.getAttribute(attributename);
		return attribute;
	}

	/**
	 * @description Used to select option in dropdown by entering visible text
	 * @param element
	 * @param text
	 */
	public static void selectOptionByText(WebElement element, String text) {
		Select select = new Select(element);
		select.selectByVisibleText(text);
	}

	/**
	 * @description Used to select option in dropdown by entering value
	 * @param element
	 * @param value
	 */
	public static void selectOptionByAttribute(WebElement element, String value) {
		Select select = new Select(element);
		select.selectByValue(value);
	}

	/**
	 * @description Used to select option in dropdown by entering index
	 * @param element
	 * @param index
	 */
	public static void selectOptionByIndex(WebElement element, int index) {
		Select select = new Select(element);
		select.selectByIndex(index);
	}

	/**
	 * @description Used to enter value in inputfields by Js
	 * @param element
	 * @param text
	 */
	public static void typeJs(WebElement element, String text) {
		JavascriptExecutor executor = (JavascriptExecutor) driver;
		executor.executeScript("arguments[0].setAttribute('value','" + text + "')", element);
	}

	/**
	 * @description Used to point the curser in particular area
	 * @param driver
	 * @param element
	 */
	public static void moveToElement(WebDriver driver, WebElement element) {
		Actions actions = new Actions(driver);
		actions.moveToElement(element).perform();
	}

	/**
	 * @description Used to move source webelement into destination webelement
	 * @param source
	 * @param destination
	 * 
	 */
	public static void dragandDrop(String source, String destination) {
		WebElement element = driver.findElement(By.id(source));
		WebElement element2 = driver.findElement(By.id(destination));
		Actions actions = new Actions(driver);
		actions.dragAndDrop(element, element2).perform();

	}

	/**
	 * @description Navigate to given url
	 * @param url
	 */
	public static void navigateurl(String url) {
		driver.navigate().to(url);
	}

	/**
	 * @description Navigate back to previous page
	 */
	public static void navigateBack() {
		driver.navigate().back();
	}

	/**
	 * @description Navigate to after page
	 */

	public static void navigateForward() {
		driver.navigate().forward();
	}

	/**
	 * @description Used to refresh particular page
	 */
	public static void navigateFresh() {
		driver.navigate().refresh();
	}

	/**
	 * @description Used to rightclick of particular element
	 * @param element
	 */
	public static void rightClick(WebElement element) {
		Actions actions = new Actions(driver);
		actions.contextClick(element).perform();

	}

	/**
	 * @description Used to enter value in inputfield without using sendkeys
	 * @param element
	 * @param text
	 */
	public static void keyUpKeyDown(WebElement element, String text) {
		Actions actions = new Actions(driver);
		actions.keyDown(Keys.SHIFT).sendKeys(element, text).keyUp(Keys.SHIFT).perform();
	}

	/**
	 * @description Used to double click that element
	 * @param element
	 */
	public static void doubleClick(WebElement element) {
		Actions actions = new Actions(driver);
		actions.doubleClick(element).perform();
	}
/**
 * @description Used to accept alert
 */
	public void alertOK() {
		Alert alert = driver.switchTo().alert();
		alert.accept();
		
	}
/**
 * @description Used to cancel the alert
 */
	public void alertDismiss() {
		Alert alert = driver.switchTo().alert();
		alert.dismiss();
		
	}
/**
 * @description Used to incert value in alert
 * @param text
 */
	public void alertPrompt(String text) {
		Alert alert = driver.switchTo().alert();
		alert.sendKeys(text);
		
	}
/**
 * @description Used to get an text from an alret
 * @return text
 */
	public String alertGetText() {
		Alert alert = driver.switchTo().alert();
		String text = alert.getText();
		return text;

	}

	/**
	 * @description Used to take screenshot of particular page
	 * @param path
	 * @return File
	 * @throws IOException
	 */
	public File getScreenShot(String path) throws IOException {
		TakesScreenshot screenshot = (TakesScreenshot) driver;
		File source = screenshot.getScreenshotAs(OutputType.FILE);
		File destination = new File(path);
		FileUtils.copyFile(source, destination);
		return source;

	}

	/**
	 * @description Used to get an entered value in element by js
	 * @param element
	 * @return Object
	 */
	public Object getJs(WebElement element) {
		JavascriptExecutor executor = (JavascriptExecutor) driver;
		Object object = executor.executeScript("return arguments[0].getAttribute('value')", element);
		return object;
	}

	/**
	 * @description Used to click an particular element by js
	 * @param element
	 */
	public static void clickJs(WebElement element) {
		JavascriptExecutor executor = (JavascriptExecutor) driver;
		executor.executeScript("arguments[0].click()", element);
	}

	/**
	 * @description Used to and Scroll down
	 * @param element
	 */
	public static void scrollDown(WebElement element) {
		JavascriptExecutor executor = (JavascriptExecutor) driver;
		executor.executeScript("arguments[0].scrollIntoView(true)", element);
	}

	/**
	 * @description Used to and Scroll Up
	 * 
	 * @param element
	 */
	public static void scrollUp(WebElement element) {
		JavascriptExecutor executor = (JavascriptExecutor) driver;
		executor.executeScript("arguments[0].scrollIntoView(false)", element);
	}
/**
 * @description Used to switch into a frame by frame index
 * @param index
 */
	public static void frame(int index) {
		driver.switchTo().frame(index);
	}
	/**
	 * @description Used to switch into a frame by frame name
	 * @param name
	 */
		
	public static void frame1(String name) {
		driver.switchTo().frame(name);
	}
	/**
	 *  @description Used to switch into a frame by frame element
	 * @param frameElement
	 */
	public static void frame2(WebElement frameElement) {
		driver.switchTo().frame(frameElement);
	}
	/**
	 * @description Used to switch into immediate  parent frame
	 * @param index
	 */
	public static void parentFrame() {
		driver.switchTo().parentFrame();
	}
	/**
	 * @description Used to switch into first frame
	 * @param index
	 */
	public static void defaultFrame() {
		driver.switchTo().defaultContent();
	}
	/**
	 * @description Used to get address of current window
	 * @return String
	 */
	public String parentWindow() {
		String handle = driver.getWindowHandle();
		return handle;
	}
/**
 * @description Used to get an address of all opened windows
 * @return Set<String>
 */
	public Set<String> allWindow() {
		Set<String> windowHandles = driver.getWindowHandles();
		return windowHandles;
	}
/**
 * description Used to checks whether dropdown supports multiple selection or not 
 * @param element
 * @return boolean
 */
	public boolean selectIsMultiple(WebElement element) {
		Select select = new Select(element);
		boolean b = select.isMultiple();
		return b;
	}
/**
 * @description  Used to deselect all the selected options
 * @param element
 */
	public static void deselectAll(WebElement element) {
		Select select = new Select(element);
		select.deselectAll();
	}
/**
 * @description Used to deselect particular  optiuon by index
 * @param element
 * @param index
 */
	public static void deselectOptionIndex(WebElement element, int index) {
		Select select = new Select(element);
		select.deselectByIndex(index);
	}
/**
 *  @description Used to deselect a patricular option by given value
 * @param element
 * @param value
 */
	public static void deselectOptionvalue(WebElement element, String value) {
		Select select = new Select(element);
		select.deselectByValue(value);
	}
/**
 * @description Used to deselect a patricular option by given text
 * @param element
 * @param text
 */
	public static void deselectOptionVisibleText(WebElement element, String text) {
		Select select = new Select(element);
		select.deselectByVisibleText(text);
	}
/**
 * @description Used to get the first selected option in a selected dropdown
 * @param element
 * @return WebElement
 */
	public WebElement getFirstSelectedOption(WebElement element) {
		Select select = new Select(element);
		WebElement firstSelectedOption = select.getFirstSelectedOption();
		return firstSelectedOption;
	}
/**
 * @description Used to get all selected option in a selected dropdown
 * @param element
 * @return list
 */
	public List allSelectedOption(WebElement element) {
		Select select = new Select(element);
		List allSelectedOptions = (List) select.getAllSelectedOptions();
		return allSelectedOptions;
	}
/**
 * @description Used to get all option in selected dropdown
 * @param element
 * @return List<WebElement>
 */
	public List<WebElement> getOptions(WebElement element) {
		Select select = new Select(element);
		List<WebElement> options = select.getOptions();
		return options;
	}
/**
 * @description Used to read value in excel sheet
 * @param path
 * @param sheetName
 * @param rowNum
 * @param cellNum
 * @return string
 * @throws IOException
 */
	public static String readValueExcel(String path, String sheetName, int rowNum, int cellNum) throws IOException {

		String res = null;

		File file = new File(path);
		FileInputStream stream = new FileInputStream(file);
		Workbook workbook = new XSSFWorkbook(stream);
		Sheet sheet = workbook.getSheet(sheetName);
		Row row = sheet.getRow(rowNum);
		Cell cell = row.getCell(cellNum);
		  CellType cellType = cell.getCellType();
		switch (cellType) {
		case STRING:
			res = cell.getStringCellValue();

			break;
		case NUMERIC:
			if (DateUtil.isCellDateFormatted(cell)) {
				Date dateCellValue = cell.getDateCellValue();
				SimpleDateFormat dateFormat = new SimpleDateFormat("dd/mm/yy");
				res = dateFormat.format(dateCellValue);
			} else {
				double d = cell.getNumericCellValue();
				long check = Math.round(d);

				if (check == d) {
					res = String.valueOf(d);
				}
				res = String.valueOf(check);
			}
			break;
		default:
			break;
		}
		workbook.close();
		return res;
	}
/**
 * @description Used to upadate value in excel sheets
 * @param path
 * @param sheetName
 * @param rowNum
 * @param cellNum
 * @param olddata
 * @param newData
 * @throws IOException
 */
	public static void update(String path, String sheetName, int rowNum, int cellNum, String olddata, String newData)
			throws IOException {

		File file = new File(path);
		FileInputStream stream = new FileInputStream(file);
		Workbook workbook = new XSSFWorkbook(stream);
		Sheet sheet = workbook.getSheet(sheetName);

		Row row = sheet.getRow(rowNum);
		Cell cell = row.getCell(cellNum);

		String value = cell.getStringCellValue();

		if (value.equals(olddata)) {
			cell.setCellValue(newData);

			FileOutputStream out = new FileOutputStream(file);
			workbook.write(out);
		}
	}
/**
 * @description Used to write value in cell in excel sheets
 * @param path
 * @param sheetName
 * @param rowNum
 * @param cellNum
 * @param newData
 * @throws IOException
 * /* Assume Row Already created 
 * */
 
	
	public static void writeData(String path, String sheetName, int rowNum, int cellNum, String newData)
			throws IOException {

		File file = new File(path);
		FileInputStream stream = new FileInputStream(file);
		Workbook workbook = new XSSFWorkbook(stream);
		Sheet sheet = workbook.getSheet(sheetName);

		Row row = sheet.getRow(rowNum);
		Cell cell = row.createCell(cellNum);

		// set Value
		cell.setCellValue(newData);

		FileOutputStream out = new FileOutputStream(file);
		workbook.write(out);
	}
}
