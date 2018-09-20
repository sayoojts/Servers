package wp.st.com.wp.st.com;

import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.concurrent.TimeUnit;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;

import org.openqa.selenium.chrome.ChromeDriver;

import org.sikuli.script.Pattern;
import org.sikuli.script.Screen;

public class UrlFetch2 {
	public static void main(String[] args) throws Exception {

		Calendar cal = Calendar.getInstance();
		SimpleDateFormat sdf = new SimpleDateFormat("HH:mm:ss");
		System.out.println("Program started at " + sdf.format(cal.getTime()));

		
		File src = new File("D:\\Users\\sanooj\\Desktop\\cm23AUG2018bhav - Copy.xls");
		FileInputStream fis = new FileInputStream(src);
		HSSFWorkbook wb = new HSSFWorkbook(fis);
		HSSFSheet sheet1 = wb.getSheetAt(0);

		String stock;

		System.setProperty("webdriver.chrome.driver",
				"D:\\Users\\sanooj\\workspace\\Cucumber\\src\\main\\resources\\drivers\\chromedriver.exe");
		WebDriver driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(1, TimeUnit.SECONDS);
		

		driver.get("http://www.moneycontrol.com");
		Thread.sleep(4000);
		System.out.println(driver.getTitle());

		Screen s = new Screen();
		Pattern stockInput = new Pattern("C:\\prep\\prep2\\Testing\\Selenium\\Sikuli\\stockInputMoneyControl.png");
		Pattern searchButton = new Pattern("C:\\prep\\prep2\\Testing\\Selenium\\Sikuli\\searchButtonMoneyControl.png");
		Pattern financials = new Pattern("C:\\prep\\prep2\\Testing\\Selenium\\Sikuli\\financialsMoneyControl.png");
		
		JavascriptExecutor js = (JavascriptExecutor) driver;

		for (int j = 1; j <= 1918; j++) {
			String originalData = null;
			String derivedData = null;
			
			try {
				
				originalData = sheet1.getRow(j).getCell(12).getStringCellValue();
				derivedData = sheet1.getRow(j).getCell(36).getStringCellValue();
				System.out.println("The value of j is "+j);
				System.out.println("The original data is "+originalData+" and the derived data is "+derivedData);
				if (!originalData.equals(derivedData)) {
					try {// for any internet error
						System.out.println("The value of j is " + j);

						stock = sheet1.getRow(j).getCell(12).getStringCellValue();
						System.out.println("The stock symbol is " + stock);

						String str;

						Thread.sleep(3000);
						js.executeScript("window.scrollBy(0,-750)");
						// s.click();
						Thread.sleep(1000);
						s.type(stockInput, stock);
						s.click(searchButton);
						Thread.sleep(1000);

						
						
						try {
							str = driver.findElement(By.xpath("//*[@id='nChrtPrc']/div[4]/div[1]")).getText();

							String[] rawArray = str.split(" ");
							String nseId = rawArray[4];
							String isin = rawArray[7];

							System.out.println("The isin from the website is "
									+ driver.findElement(By.xpath("//*[@id='nChrtPrc']/div[4]/div[1]")).getText());
							System.out.println("The refined text is " + isin);

							if (isin.equals(stock)) {// only if the symbol in excel matches that of website write in excel

								try {
									js.executeScript("window.scrollBy(0,750)");
								} catch (Exception e) {

								}

								Thread.sleep(500);
								try {
									s.click(financials);
								} catch (Exception e) {

								}
								Thread.sleep(2500);
								String rawUrl = driver.getCurrentUrl();

								System.out.println("the url is " + rawUrl);
								sheet1.getRow(j).createCell(36).setCellValue(stock);
								sheet1.getRow(j).createCell(37).setCellValue(rawUrl);
								sheet1.getRow(j).createCell(38).setCellValue(isin);
								sheet1.getRow(j).createCell(39).setCellValue(nseId);

								FileOutputStream fout = new FileOutputStream(
										"D:\\Users\\sanooj\\Desktop\\cm23AUG2018bhav - Copy.xls");

								
								js.executeScript("window.scrollBy(0,-750)");
								wb.write(fout);
							} else {
								System.out.println("No Mathch");
								System.out.println(stock);
								System.out.println(isin);
							}
						} // end of try

						catch (Exception e) {

							String rawUrl = driver.getCurrentUrl();

							System.out.println("the url is " + rawUrl);
							sheet1.getRow(j).createCell(36).setCellValue(stock);
							sheet1.getRow(j).createCell(37).setCellValue(rawUrl);

							FileOutputStream fout = new FileOutputStream(
									"D:\\Users\\sanooj\\Desktop\\cm23AUG2018bhav - Copy.xls");

							
							js.executeScript("window.scrollBy(0,-750)");
							wb.write(fout);
						}
					} 
					catch (Exception e) {
						driver.navigate().refresh();

					}

				}

			}
			catch(Exception e) {
		/*		
				System.out.println("The value of j is "+j);
				
				if (!originalData.equals(derivedData)) {
					try {// for any internet error
						System.out.println("The value of j is " + j);

						stock = sheet1.getRow(j).getCell(12).getStringCellValue();
						System.out.println("The stock symbol is " + stock);

						String str;

						Thread.sleep(3000);
						js.executeScript("window.scrollBy(0,-750)");
						// s.click();
						Thread.sleep(1000);
						s.type(stockInput, stock);
						s.click(searchButton);
						Thread.sleep(1000);

						
						
						try {
							str = driver.findElement(By.xpath("//*[@id='nChrtPrc']/div[4]/div[1]")).getText();

							String[] rawArray = str.split(" ");
							String nseId = rawArray[4];
							String isin = rawArray[7];

							System.out.println("The isin from the website is "
									+ driver.findElement(By.xpath("//*[@id='nChrtPrc']/div[4]/div[1]")).getText());
							System.out.println("The refined text is " + isin);

							if (isin.equals(stock)) {// only if the symbol in excel matches that of website write in excel

								try {
									js.executeScript("window.scrollBy(0,750)");
								} catch (Exception e3) {

								}

								Thread.sleep(500);
								try {
									s.click(financials);
								} catch (Exception e4) {

								}
								Thread.sleep(2500);
								String rawUrl = driver.getCurrentUrl();

								System.out.println("the url is " + rawUrl);
								sheet1.getRow(j).createCell(36).setCellValue(stock);
								sheet1.getRow(j).createCell(37).setCellValue(rawUrl);
								sheet1.getRow(j).createCell(38).setCellValue(isin);
								sheet1.getRow(j).createCell(39).setCellValue(nseId);

								FileOutputStream fout = new FileOutputStream(
										"D:\\Users\\sanooj\\Desktop\\cm23AUG2018bhav - Copy.xls");

								
								js.executeScript("window.scrollBy(0,-750)");
								wb.write(fout);
							} else {
								System.out.println("No Mathch");
								System.out.println(stock);
								System.out.println(isin);
							}
						} // end of try

						catch (Exception e6) {

							String rawUrl = driver.getCurrentUrl();

							System.out.println("the url is " + rawUrl);
							sheet1.getRow(j).createCell(36).setCellValue(stock);
							sheet1.getRow(j).createCell(37).setCellValue(rawUrl);

							FileOutputStream fout = new FileOutputStream(
									"D:\\Users\\sanooj\\Desktop\\cm23AUG2018bhav - Copy.xls");

							
							js.executeScript("window.scrollBy(0,-750)");
							wb.write(fout);
						}
					} 
					catch (Exception e2) {
						driver.navigate().refresh();

					}

				}*/
				
				
			}
		} // end of for
		Calendar cal2 = Calendar.getInstance();
		SimpleDateFormat sdf2 = new SimpleDateFormat("HH:mm:ss");
		System.out.println("Program started at " + sdf.format(cal.getTime()));
		System.out.println("Program ended at " + sdf2.format(cal2.getTime()));

		/*String shutdownCommand;
		shutdownCommand = "shutdown.exe -s -t 0";
		Runtime.getRuntime().exec(shutdownCommand);
		System.exit(0);

		Thread.sleep(4000);

		try {
			Robot robot = new Robot();
			robot.keyPress(KeyEvent.VK_ENTER);
			robot.keyRelease(KeyEvent.VK_ENTER);
			
			robot.delay(5000);
			
			robot.keyPress(KeyEvent.VK_ENTER);
			robot.keyRelease(KeyEvent.VK_ENTER);
		} catch (Exception e) {

		}
*/
	}//end of main

}//end of class
