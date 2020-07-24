package com.ExamTask;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;
import org.testng.Reporter;
import org.testng.annotations.AfterClass;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Test;

import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Image;
import com.itextpdf.text.PageSize;
import com.itextpdf.text.Paragraph;
import com.itextpdf.text.pdf.PdfWriter;

public class SeatBookingSSwithPdf {

	WebDriver driver;

	@Test
	public void f() throws EncryptedDocumentException, IOException, DocumentException {
		//Declare web site URL
		String SiteUrl = "https://www.shohoz.com/";
		//set mouse action 
		Actions act = new Actions(driver); 
		Reporter.log("Starting Test Case.");
		driver.get(SiteUrl);
		Reporter.log("Navigated to shohoz.com");
		//Set Excel file path
		FileInputStream fis = new FileInputStream(System.getProperty("user.dir")+"\\src\\test\\resources\\TestData\\SearchContent.xlsx");
		//Workbook
		Workbook wb=WorkbookFactory.create(fis);
		//Excel Sheet
		Sheet sh=wb.getSheet("Destination");
		//Excel Row
		Row rw=sh.getRow(1);
		//Excel Column 1
		Cell cel1=rw.getCell(0);
		//Store from cell value
		String FromData=cel1.getStringCellValue();
		//Excel Column 2
		Cell cel2=rw.getCell(1);
		//Store to cell value
		String ToData=cel2.getStringCellValue();
		//locate bus
		WebElement Bus = driver.findElement(By.xpath("/html/body/header/div[2]/div/nav/ul/li[4]/a"));
		Bus.click();
		Reporter.log("Navigated to Bus Search Page.");
		//locate from and enter value
		WebElement From = driver.findElement(By.xpath("//*[@id=\"dest_from\"]"));
		From.click();
		From.sendKeys(FromData);
		Reporter.log("Entered Starting From location.");

		//locate to and enter value
		WebElement To = driver.findElement(By.xpath("//*[@id=\"dest_to\"]"));
		To.click();
		To.sendKeys(ToData);
		Reporter.log("Entered destination location.");

		//Locate Journey date picker and click to show calendar
		WebElement JourneyDatePicker = driver.findElement(By.xpath("//*[@id=\"doj\"]"));
		JourneyDatePicker.click();
		driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
		//Selected month and date
		String Month="August 2020";
		while(true)
		{
			//Locate selected month
			String PresentMonth=driver.findElement(By.xpath("/html/body/div[6]/div/div")).getText();
			if(PresentMonth.equalsIgnoreCase(Month))
			{
				break;
			}
			else
			{
				//Locate next icon
				WebElement NextIcon = driver.findElement(By.xpath("/html/body/div[6]/div/a[2]"));
				NextIcon.click();
			}
		}
		//Locate Date widget and select date
		WebElement DateWidget = driver.findElement(By.xpath("//*[@id=\"ui-datepicker-div\"]"));
		List<WebElement> rows = DateWidget.findElements(By.tagName("tr"));
		List<WebElement> cols = DateWidget.findElements(By.tagName("td"));
		for(WebElement cell:cols)
		{
			//select 11th date
			if(cell.getText().equals("11"))
			{
				cell.findElement(By.linkText("11"));
				act.doubleClick(cell).perform();
				break;
			}
		}
		//Locate return date picker and clear field
		WebElement ReturnDatePicker = driver.findElement(By.xpath("//*[@id=\"dor\"]"));
		ReturnDatePicker.clear();
		//Locate Search button and click
		WebElement SearchButton = driver.findElement(By.xpath("/html/body/div[2]/section/div[1]/div[3]/div[1]/form/ul/div[5]/div[2]/button"));
		SearchButton.click();
		Reporter.log("Clicked on the search button");
		//wait for 30sec
		driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
		//Locate dep timetable
		WebElement DepTemp =driver.findElement(By.xpath("/html/body/div[8]/div[2]/div[5]/div/div[2]/div/div/div/table/tbody/tr[1]/td[2]"));
		DepTemp.click();
		//take screenshot and store it in byte format
		byte[] SearchResultSS = ((TakesScreenshot)driver).getScreenshotAs(OutputType.BYTES);
		//Locate view seat button
		WebElement ViewSeatButton = driver.findElement(By.xpath("//*[@id=\"5606046\"]"));
		//ViewSeatButton.click();
		act.doubleClick(ViewSeatButton).perform();
		Reporter.log("Viewed available Seats.");
		//Locate seat
		WebElement Seat = driver.findElement(By.xpath("/html/body/div[8]/div[2]/div[5]/div/div[2]/div/div/div/table/tbody/tr[2]/td/table/tbody/tr[2]/td/div/div/div[1]/div[2]/div[1]/div[1]/ul[5]/li[2]/a/div"));
		act.doubleClick(Seat).perform();
		Reporter.log("Seat Selected.");
		//act.click(Seat).perform();
		//Locate boarding Point
		WebElement BoardingPoint = driver.findElement(By.xpath("//*[@id=\"boardingpoint\"]"));
		Select BoardingPointDrop = new Select(BoardingPoint);
		BoardingPointDrop.selectByValue("65596189");
		//take screenshot and store it in byte format
		byte[] BookedSeatSS = ((TakesScreenshot)driver).getScreenshotAs(OutputType.BYTES);

		//create document
		Document document = new Document();
		String output = "D:\\Automation\\Scripts\\FinalExamTask\\ResultPDF\\Seat Booking.pdf";
		FileOutputStream fos = new FileOutputStream(output);
		//instantiate pdf writer
		PdfWriter writer = PdfWriter.getInstance(document, fos);
		//open the pdf for writing 
		writer.open();
		document.open();
		//process content into add image
		Image img1 = Image.getInstance(SearchResultSS);
		Image img2 = Image.getInstance(BookedSeatSS);
		//set the size of the image
		img1.scaleToFit(PageSize.A4.getWidth()/2, PageSize.A4.getHeight()/2);
		img2.scaleToFit(PageSize.A4.getWidth()/2, PageSize.A4.getHeight()/2);
		//add image to pdf
		document.add(new Paragraph("Search Page Result "));
		document.add(img1);
		document.add(new Paragraph(" "));
		document.add(new Paragraph("Seat Booked "));
		document.add(img2);
		document.add(new Paragraph(" "));
		//close the files and save to local drive
		document.close();
		writer.close();

	}
	@BeforeMethod
	public void beforeMethod() {
		//Launch Firefox
		driver = new FirefoxDriver();
		//set browser size
		driver.manage().window().maximize();
	}

	@AfterMethod
	public void afterMethod() {
		//close driver
				driver.close();
	}

	@BeforeClass
	public void beforeClass() {
		//Declare Driver path
		String GeckoDriverPath = "D:\\Softwares\\Selenium Automation\\geckodriver.exe";
		//set gecko driver path
		System.setProperty("webdriver.gecko.driver",GeckoDriverPath);
	}

	@AfterClass
	public void afterClass() {
		Reporter.log("Firefox Closed");
	}

}
