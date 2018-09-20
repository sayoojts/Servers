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

public class FindMissing {
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

		Thread.sleep(4000);
		driver.get("http://www.moneycontrol.com");
		
		System.out.println(driver.getTitle());

		Screen s = new Screen();
		Pattern stockInput = new Pattern("C:\\prep\\prep2\\Testing\\Selenium\\Sikuli\\stockInputMoneyControl.png");
		Pattern searchButton = new Pattern("C:\\prep\\prep2\\Testing\\Selenium\\Sikuli\\searchButtonMoneyControl.png");
		Pattern financials = new Pattern("C:\\prep\\prep2\\Testing\\Selenium\\Sikuli\\financialsMoneyControl.png");

		JavascriptExecutor js = (JavascriptExecutor) driver;

		
			
		
		for (int j = 1508; j <= 1918; j++) {
			
			try
			{//if any error within this loop, so that pgm can execute continiously
				//unexpected events
			System.out.println("The value of j is " + j);

			try {// check  if the data is null

				sheet1.getRow(j).getCell(38).getStringCellValue();
			}

			catch (Exception nullDataInCell) {

				stock = sheet1.getRow(j).getCell(0).getStringCellValue();
				String isinValue = sheet1.getRow(j).getCell(12).getStringCellValue();
				System.out.println("The stock symbol is " + stock);

				String str;
				String rawUrl = null;

				Thread.sleep(3000);
				js.executeScript("window.scrollBy(0,-1000)");
				// s.click();
				Thread.sleep(1000);
				s.type(stockInput, stock);
				Thread.sleep(2000);
				/*driver.findElement(
						By.xpath("(.//*[normalize-space(text()) and normalize-space(.)='All'])[1]/following::span[1]"))
						.click();*/
				s.click(searchButton);
				

				Thread.sleep(1000);
				
				String nseidStockName;

				try {//if there is only one value then fine
					str = driver.findElement(By.xpath("//*[@id='nChrtPrc']/div[4]/div[1]")).getText();
				}
				catch(Exception whenThereAreMultipleCompanies) {
					int size = driver.findElements(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div/table/tbody/tr")).size();
					
					System.out.println("the size is " +size);
					for(int i=1;i<=size;i++) {
						System.out.print("The "+i+"th element is ");
						//System.out.println(driver.findElement(By.xpath("//*[@id=\'mc_mainWrapper\']/div[3]/div[2]/div[1]/div["+i+"]")).getText());
						nseidStockName = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div/table/tbody/tr["+i+"]/td[2]/p/span[1]")).getText();
						System.out.println(driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div/table/tbody/tr["+i+"]/td[2]/p/span[1]")).getText());
						 if(nseidStockName.equalsIgnoreCase("NSE Id :"+stock)){
							 driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div/table/tbody/tr["+i+"]/td[1]/p/a")).click();
							 break;
						 }//end of if
					}//end of size for
					try {//if the exact match is in the 1nd column
					str = driver.findElement(By.xpath("//*[@id='nChrtPrc']/div[4]/div[1]")).getText();
					}
					catch(Exception secondColumn) {
						for(int i=1;i<=size;i++) {
							System.out.print("The "+i+"th element is ");
							//System.out.println(driver.findElement(By.xpath("//*[@id=\'mc_mainWrapper\']/div[3]/div[2]/div[1]/div["+i+"]")).getText());
							try {//bcoz, not all the rows have more than 1 value
								nseidStockName = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div/table/tbody/tr["+i+"]/td[2]/p/span[2]")).getText();	
							}
							catch(Exception noSecondValue) {
								nseidStockName = "nothing to worry go to the next";
							}
							
							//System.out.println(driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div/table/tbody/tr["+i+"]/td[2]/p/span[2]")).getText());
							 if(nseidStockName.equalsIgnoreCase("NSE Id :"+stock)){
								 driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div/table/tbody/tr["+i+"]/td[1]/p/a")).click();
								 break;
							 }//end of if
						}//end of size for
						str = driver.findElement(By.xpath("//*[@id='nChrtPrc']/div[4]/div[1]")).getText();
						
					}
				}//end of catch

					String[] rawArray = str.split(" ");
					String nseId = rawArray[4];
					String isin = rawArray[7];

					System.out.println("The isin from the website is "
							+ driver.findElement(By.xpath("//*[@id='nChrtPrc']/div[4]/div[1]")).getText());
					System.out.println("The refined text is " + isin);
					
					js.executeScript("window.scrollBy(0,1000)");
					
					

						Thread.sleep(500);

						s.click(financials);
						Thread.sleep(2500);
						rawUrl = driver.getCurrentUrl();

						System.out.println("the url is " + rawUrl);
						sheet1.getRow(j).createCell(36).setCellValue(stock);
						sheet1.getRow(j).createCell(37).setCellValue(rawUrl);
						sheet1.getRow(j).createCell(38).setCellValue(isin);
						sheet1.getRow(j).createCell(39).setCellValue(nseId);
						sheet1.getRow(j).createCell(40).setCellValue("check it once");

						FileOutputStream fout = new FileOutputStream(
								"D:\\Users\\sanooj\\Desktop\\cm23AUG2018bhav - Copy.xls");

						js.executeScript("window.scrollBy(0,-750)");
						wb.write(fout);
					
				
				

			}//end of catch "nullDataInCell"
		}//end of try 
		catch(Exception anyUnexpectedException) {
			//do nothing
		}
		} // end of for

	}// end of main

}// end of class
