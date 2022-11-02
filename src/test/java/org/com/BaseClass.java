package org.com;

import java.awt.Dimension;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.time.Duration;
import java.util.Date;
import java.util.List;
import java.util.Set;


import org.apache.commons.io.FileUtils;
import org.apache.commons.lang3.time.DateUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
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
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.FluentWait;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.Wait;
import org.openqa.selenium.support.ui.WebDriverWait;

import io.github.bonigarcia.wdm.WebDriverManager;

public class BaseClass {
	
	static WebDriver driver;
	public static void getDriver() {
		WebDriverManager.chromedriver().setup();
		driver=new ChromeDriver();
	}
	public static void loadurl(String url) {
		driver.get(url);

	}
	public void type(WebElement element, String datas) {
		element.sendKeys(datas);
	}
	public void click(WebElement element ) {
		element.click();
	}
	public static void maximize() {
		driver.manage().window().maximize();
	}
	public static void closeAllWindow() {
		driver.quit();
	}
	public void clear() {
		
	}

	public WebElement findElementById(String id) {
		WebElement element = driver.findElement(By.id(id));
		return element;
		
	}
	public WebElement findElementByName(String name) {
		WebElement element = driver.findElement(By.name(name));
		return element;
		
	}
	public WebElement findElementByClassName(String className) {
	    WebElement element = driver.findElement(By.className(className));
	    return  element;
	}
	public WebElement findElementByXpath(String xpath) {
		WebElement element = driver.findElement(By.xpath(xpath));
		return element;
	}
	public String getTitle() {
		String title = driver.getTitle();
		return title;
	}
	public String geturl() {
		String url = driver.getCurrentUrl();
		return url;
		
	}
	public String getText(WebElement element) {
		String text = element.getText();
		return text;
	}

	public String insertedValue(WebElement element) {
		String attribute = element.getAttribute("value");
		return attribute;
	}
	public String getAttribute(WebElement element,String attributename ) {
		String attribute = element.getAttribute(attributename);
		return attribute;
	}
	public void selectOptionByText(WebElement element,String text) {
		Select select = new Select(element);
		select.selectByVisibleText(text);
	}
	public void selectOptionByAttribute(WebElement element,String value) {	
	Select select = new Select(element);
	select.selectByValue(value);
	}
	public void selectOptionByIndex(WebElement element,int index) {
		Select select = new Select(element);
		select.selectByIndex(index);
	}
	public void typeJs(WebElement element,String text) {
		JavascriptExecutor executor =(JavascriptExecutor) driver;
		executor.executeScript("arguments[0].setAttribute('value',"+text+")",element);
	}
	public void moveToElement(WebDriver driver,WebElement element) {
		Actions actions = new Actions(driver);
		actions.moveToElement(element).perform();
	}
	public WebElement dragandDrop(String source,String destination) {
		WebElement element = driver.findElement(By.id(source));
		WebElement element2 = driver.findElement(By.id(destination));
		Actions actions = new Actions(driver);
		actions.dragAndDrop(element, element2).perform();
		return element;	
	}

	public void navigateurl(String url) {
		driver.navigate().to(url);
	}
	public void navigateBack() {
		driver.navigate().back();
	}
	public void navigateForward() {
		driver.navigate().forward();
	}
	public void navigateFresh() {
		driver.navigate().refresh();
	}

	public void rightClick(WebElement element) {
	       Actions actions = new Actions(driver);
	    actions.contextClick(element).perform();
	    
	}
	public void keyUpKeyDown(WebElement element,String text ) {
		Actions actions = new Actions(driver);
		actions.keyDown(Keys.SHIFT).sendKeys(element,text).keyUp(Keys.SHIFT).perform();
	}
	public void doubleClick(WebElement element) {
		Actions actions = new Actions(driver);
		actions.doubleClick(element).perform();
	}
	public Alert alertOK() {
		Alert alert = driver.switchTo().alert();
		alert.accept();
		return alert;
	}
	public Alert alertDismiss() {
		Alert alert = driver.switchTo().alert();
		alert.dismiss();
		return alert;
	}
	public Alert alertPrompt(String  text) {
		Alert alert = driver.switchTo().alert();
		alert.sendKeys(text);
		return alert;
	}
	public String alertGetText() {
		Alert alert = driver.switchTo().alert();
		String text = alert.getText();
		return text;
		
	}
	private File getScreenShot(String path) throws IOException {
		TakesScreenshot screenshot = (TakesScreenshot) driver;
		File source = screenshot.getScreenshotAs(OutputType.FILE);
		File destination= new File(path);
		FileUtils.copyFile(source, destination);
		return source;
		
	}
	public Object getJs(WebElement element) {
		JavascriptExecutor executor =(JavascriptExecutor) driver;
		Object object = executor.executeScript("return arguments[0].getAttribute('value')", element);
	return object;
	}
	public void clickJs(WebElement element) {
		JavascriptExecutor executor =(JavascriptExecutor) driver;
		executor.executeScript("arguments[0].click()", element);
	}
	public void scrollDown(WebElement element) {
		JavascriptExecutor executor =(JavascriptExecutor) driver;
		executor.executeScript("arguments[0].scrollIntoView(true)", element);
	}
	public void scrollUp(WebElement element) {
		JavascriptExecutor executor =(JavascriptExecutor) driver;
		executor.executeScript("arguments[0].scrollIntoView(false)", element);
	}
	public void frame(int index) {
		driver.switchTo().frame(index);
	}
	public void frame1(String name) {
		driver.switchTo().frame(name);
	}
	public void frame2(WebElement frameElement) {
		driver.switchTo().frame(frameElement);
	}
	public void parentFrame() {
		driver.switchTo().parentFrame();
	}
	public void defaultFrame() {
		driver.switchTo().defaultContent();
	}
	public void implicitlyWait(int seconds) {
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(seconds));
	}
	public void explicitWait(int seconds) {
		WebDriverWait wait = new WebDriverWait(driver,Duration.ofSeconds(seconds));
	}
	public void fluentWait(int seconds, int seconds1,WebElement element) {
		Wait wait = new FluentWait(driver).withTimeout(Duration.ofSeconds(seconds)).pollingEvery(Duration.ofSeconds(seconds1)).ignoring(Throwable.class);
		WebElement click = (WebElement)wait.until(ExpectedConditions.elementToBeClickable(element));
	}
	public void pageLoadTimeOut(int seconds) {
	    driver.manage().timeouts().pageLoadTimeout(Duration.ofSeconds(seconds));
	}

	public String parentWindow() {
		String handle = driver.getWindowHandle();
		return handle;
	}
	public Set<String> allWindow() {
		Set<String> windowHandles = driver.getWindowHandles();
		return windowHandles;
	}
	public boolean selectIsMultiple(WebElement element) {
		Select select = new Select(element);
		boolean b = select.isMultiple();
		return b;
	}
	public void deselectAll(WebElement element) {
		Select select = new Select(element);
		select.deselectAll();
	}
	public void deselectOptionIndex(WebElement element,int index) {
		Select select = new Select(element);
		select.deselectByIndex(index);
	}
	public void deselectOptionvalue(WebElement element,String value) {
		Select select = new Select(element);
	    select.deselectByValue(value);
	}
	public void deselectOptionVisibleText(WebElement element,String text) {
		Select select = new Select(element);
		select.deselectByVisibleText(text);
	}
	public WebElement getFirstSelectedOption(WebElement element) {
		Select select = new Select(element);
		WebElement firstSelectedOption = select.getFirstSelectedOption();
		return firstSelectedOption;	
	}
	public List<WebElement> allSelectedOption(WebElement element) {
		Select select = new Select(element);
		List<WebElement> allSelectedOptions = select.getAllSelectedOptions();
		return allSelectedOptions;		
	}
	public List<WebElement> getOptions(WebElement element) {
		Select select = new Select(element);
		List<WebElement> options = select.getOptions();
		return options;

	}

	public List<WebElement> frameCount(String xpath) {

	   List<WebElement> findElements = driver.findElements(By.xpath(xpath));
	   return findElements;
	}
	public void sleep(int seconds) throws InterruptedException {
		Thread.sleep(seconds);
	}
	public Dimension getSize(WebElement element) {
		Dimension dimension = element.getSize();
		return dimension;
	}
	public void getData(String sheetname, int rownum, int cellnum, String olddata, String newdata) throws IOException {
		File file = new File("C:\\Users\\HP\\eclipse-workspace\\SampleProject\\Excel\\sample excel.xlsx");
		FileInputStream stream =new FileInputStream(file);
		Workbook workbook =new XSSFWorkbook(stream);
		Sheet sheet = workbook.getSheet(sheetname);
		Row row = sheet.getRow(rownum);
		Cell cell = row.getCell(cellnum);
		String value = cell.getStringCellValue();
		if(value.equals(olddata)) {
			cell.setCellValue(newdata);
			FileOutputStream out = new FileOutputStream(file);
			workbook.write(out);
		}
	}
	public void createCell(String sheetname, int rownum, int cellnum, String data) throws IOException {
		File file = new File("C:\\Users\\HP\\eclipse-workspace\\SampleProject\\Excel\\sample excel.xlsx");
		FileInputStream stream =new FileInputStream(file);
		Workbook workbook =new XSSFWorkbook(stream);
		Sheet sheet = workbook.getSheet(sheetname);
		Row row = sheet.getRow(rownum);
		Cell createCell = row.createCell(cellnum );
		createCell.setCellValue(data);
		FileOutputStream out = new FileOutputStream(file);
		workbook.write(out);
		}
	public void createWorkbook(String sheetname, int rownum, int cellnum, String data) throws IOException {
		File file = new File("C:\\Users\\HP\\eclipse-workspace\\SampleProject\\Excel\\sample excel.xlsx");
		FileInputStream stream =new FileInputStream(file);
		Workbook workbook =new XSSFWorkbook(stream);
		Sheet sheet = workbook.getSheet(sheetname);
		Row row = sheet.getRow(rownum);
		Cell createCell = row.createCell(cellnum);
		createCell.setCellValue(data);
		FileOutputStream out = new FileOutputStream(file);
		workbook.write(out);
		}
	public void createRow (String sheetname, int index, int index1, String data) throws IOException {
		File file = new File("C:\\Users\\HP\\eclipse-workspace\\SampleProject\\Excel\\sample excel.xlsx");
		FileInputStream stream =new FileInputStream(file);
		Workbook workbook =new XSSFWorkbook(stream);
		Sheet sheet = workbook.getSheet(sheetname);
		Row row = sheet.createRow(index);
		Cell createCell = row.createCell(index1);
		createCell.setCellValue(data);
		FileOutputStream out = new FileOutputStream(file);
		workbook.write(out);
		}
	public String getData(String sheetname,int rownum, int cellnum) throws IOException {
		String res=null;
		File file = new File("C:\\Users\\HP\\eclipse-workspace\\SampleProject\\Excel\\sample excel.xlsx");
		FileInputStream stream =new FileInputStream(file);
		Workbook workbook =new XSSFWorkbook(stream);
		Sheet sheet = workbook.getSheet(sheetname);
		Row row = sheet.getRow(rownum);
		Cell cell = row.getCell(cellnum);
		CellType type = cell.getCellType();
		switch (type) {
		case STRING:
			res= cell.getStringCellValue();
			
			break;
		case NUMERIC:
			if (DateUtils.isCellDateFormatted(cell)) {
				Date dateCellValue = cell.getDateCellValue();
				SimpleDateFormat dateFormat= new SimpleDateFormat("dd/mm/yy");
				res=dateFormat.format(dateCellValue);
			}else {
				double numericCellValue = cell.getNumericCellValue();
				BigDecimal valueof= BigDecimal.valueOf(numericCellValue);
				res= valueof.toString();		
			}
			break;
		default:
			break;
		}
		return res;
	}
	
	
	
	
	
}

