package wp.st.com.wp.st.com;

import java.io.File;
import java.io.FileInputStream;

import java.io.FileOutputStream;
import java.util.concurrent.TimeUnit;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class RefinedRatiosSpecific {

	static String firstPart = "//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[";
	// //*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[7]/td[1]
	static String secondPart = "]/td[1]";

	static String colFirstPart = "//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[3]/td[";
	static String colSecondPart = "]";

	//// *[@id="mc_mainWrapper"]/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[3]/td[3]
	// *[@id="mc_mainWrapper"]/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[3]/td[4]

	static String rowIterationText;
	static String colIterationText;
	/*
	 * static HSSFWorkbook wb; static HSSFSheet sheet1; static FileInputStream fis;
	 * static File src;
	 */

	public static void main(String[] args) throws Exception {

		System.setProperty("webdriver.chrome.driver",
				"D:\\Users\\sanooj\\workspace\\Cucumber\\src\\main\\resources\\drivers\\chromedriver.exe");
		WebDriver driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(1, TimeUnit.SECONDS);

		int rowSize;// =
					// driver.findElements(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr/td[1]")).size();
		int colSize;// =
					// driver.findElements(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[3]/td")).size();

		File src = new File("D:\\Users\\sanooj\\Desktop\\RefinedRatios.xls");
		FileInputStream fis = new FileInputStream(src);
		HSSFWorkbook wb = new HSSFWorkbook(fis);
		HSSFSheet sheet1 = wb.getSheetAt(0);
		FileOutputStream fout = null;

		for (int stockIterator = 1; stockIterator <= 3; stockIterator++) {

			// System.out.println("The value of stockIterator is "+stockIterator);
			String balanceSheetUrl = sheet1.getRow(stockIterator).getCell(1).getStringCellValue();
			String ratioUrl = balanceSheetUrl.replaceAll("balance-sheetVI", "ratiosVI");

			driver.get(ratioUrl);

			rowSize = driver.findElements(By.xpath(
					"//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr/td[1]"))
					.size();
			colSize = driver.findElements(By.xpath(
					"//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[3]/td"))
					.size();

			ratioValues(stockIterator, 6, 5, driver, src, fis, wb, sheet1, fout, "Basic EPS (Rs.)", "Mar 18", rowSize,
					colSize, firstPart, secondPart, colFirstPart, colSecondPart);

			ratioValues(stockIterator, 7, 10, driver, src, fis, wb, sheet1, fout, "Diluted EPS (Rs.)", "Mar 18",
					rowSize, colSize, firstPart, secondPart, colFirstPart, colSecondPart);
			/*
			 * ratioValues(stockIterator,driver,src,fis,wb,sheet1,fout,"Diluted EPS (Rs.)"
			 * ,"Mar 17",rowSize,colSize,firstPart,secondPart,colFirstPart,colSecondPart);
			 * ratioValues(stockIterator,driver,src,fis,wb,sheet1,fout,"Diluted EPS (Rs.)"
			 * ,"Mar 16",rowSize,colSize,firstPart,secondPart,colFirstPart,colSecondPart);
			 * ratioValues(stockIterator,driver,src,fis,wb,sheet1,fout,"Diluted EPS (Rs.)"
			 * ,"Mar 15",rowSize,colSize,firstPart,secondPart,colFirstPart,colSecondPart);
			 * ratioValues(stockIterator,driver,src,fis,wb,sheet1,fout,"Diluted EPS (Rs.)"
			 * ,"Mar 14",rowSize,colSize,firstPart,secondPart,colFirstPart,colSecondPart);
			 */

			ratioValues(stockIterator, 8, 15, driver, src, fis, wb, sheet1, fout, "Cash EPS (Rs.)", "Mar 18", rowSize,
					colSize, firstPart, secondPart, colFirstPart, colSecondPart);
			/*
			 * ratioValues(stockIterator,driver,src,fis,wb,sheet1,fout,"Cash EPS (Rs.)"
			 * ,"Mar 17",rowSize,colSize,firstPart,secondPart,colFirstPart,colSecondPart);
			 * ratioValues(stockIterator,driver,src,fis,wb,sheet1,fout,"Cash EPS (Rs.)"
			 * ,"Mar 16",rowSize,colSize,firstPart,secondPart,colFirstPart,colSecondPart);
			 * ratioValues(stockIterator,driver,src,fis,wb,sheet1,fout,"Cash EPS (Rs.)"
			 * ,"Mar 15",rowSize,colSize,firstPart,secondPart,colFirstPart,colSecondPart);
			 * ratioValues(stockIterator,driver,src,fis,wb,sheet1,fout,"Cash EPS (Rs.)"
			 * ,"Mar 14",rowSize,colSize,firstPart,secondPart,colFirstPart,colSecondPart);
			 */

			ratioValues(stockIterator, 10, 20, driver, src, fis, wb, sheet1, fout,
					"Book Value [InclRevalReserve]/Share (Rs.)", "Mar 18", rowSize, colSize, firstPart, secondPart,
					colFirstPart, colSecondPart);
			ratioValues(stockIterator, 11, 25, driver, src, fis, wb, sheet1, fout,
					"Revenue from Operations/Share (Rs.)", "Mar 18", rowSize, colSize, firstPart, secondPart,
					colFirstPart, colSecondPart);
			ratioValues(stockIterator, 12, 30, driver, src, fis, wb, sheet1, fout, "PBDIT/Share (Rs.)", "Mar 18",
					rowSize, colSize, firstPart, secondPart, colFirstPart, colSecondPart);
			ratioValues(stockIterator, 13, 35, driver, src, fis, wb, sheet1, fout, "PBIT/Share (Rs.)", "Mar 18",
					rowSize, colSize, firstPart, secondPart, colFirstPart, colSecondPart);
			ratioValues(stockIterator, 14, 40, driver, src, fis, wb, sheet1, fout, "PBT/Share (Rs.)", "Mar 18", rowSize,
					colSize, firstPart, secondPart, colFirstPart, colSecondPart);
			ratioValues(stockIterator, 15, 45, driver, src, fis, wb, sheet1, fout, "Net Profit/Share (Rs.)", "Mar 18",
					rowSize, colSize, firstPart, secondPart, colFirstPart, colSecondPart);

			// Profitability Ratios
			ratioValues(stockIterator, 17, 50, driver, src, fis, wb, sheet1, fout, "PBDIT Margin (%)", "Mar 18",
					rowSize, colSize, firstPart, secondPart, colFirstPart, colSecondPart);
			ratioValues(stockIterator, 18, 55, driver, src, fis, wb, sheet1, fout, "PBIT Margin (%)", "Mar 18", rowSize,
					colSize, firstPart, secondPart, colFirstPart, colSecondPart);
			ratioValues(stockIterator, 19, 60, driver, src, fis, wb, sheet1, fout, "PBT Margin (%)", "Mar 18", rowSize,
					colSize, firstPart, secondPart, colFirstPart, colSecondPart);
			ratioValues(stockIterator, 20, 65, driver, src, fis, wb, sheet1, fout, "Net Profit Margin (%)", "Mar 18",
					rowSize, colSize, firstPart, secondPart, colFirstPart, colSecondPart);
			ratioValues(stockIterator, 21, 70, driver, src, fis, wb, sheet1, fout, "Return on Networth / Equity (%)",
					"Mar 18", rowSize, colSize, firstPart, secondPart, colFirstPart, colSecondPart);
			ratioValues(stockIterator, 22, 75, driver, src, fis, wb, sheet1, fout, "Return on Capital Employed (%)",
					"Mar 18", rowSize, colSize, firstPart, secondPart, colFirstPart, colSecondPart);
			ratioValues(stockIterator, 23, 80, driver, src, fis, wb, sheet1, fout, "Return on Assets (%)", "Mar 18",
					rowSize, colSize, firstPart, secondPart, colFirstPart, colSecondPart);
			ratioValues(stockIterator, 24, 85, driver, src, fis, wb, sheet1, fout, "Total Debt/Equity (X)", "Mar 18",
					rowSize, colSize, firstPart, secondPart, colFirstPart, colSecondPart);
			ratioValues(stockIterator, 25, 90, driver, src, fis, wb, sheet1, fout, "Asset Turnover Ratio (%)", "Mar 18",
					rowSize, colSize, firstPart, secondPart, colFirstPart, colSecondPart);

			// Liquidity Ratios
			ratioValues(stockIterator, 27, 95, driver, src, fis, wb, sheet1, fout, "Current Ratio (X)", "Mar 18",
					rowSize, colSize, firstPart, secondPart, colFirstPart, colSecondPart);
			ratioValues(stockIterator, 28, 100, driver, src, fis, wb, sheet1, fout, "Quick Ratio (X)", "Mar 18",
					rowSize, colSize, firstPart, secondPart, colFirstPart, colSecondPart);
			ratioValues(stockIterator, 29, 105, driver, src, fis, wb, sheet1, fout, "Inventory Turnover Ratio (X)",
					"Mar 18", rowSize, colSize, firstPart, secondPart, colFirstPart, colSecondPart);

			// Valuation Ratios
			ratioValues(stockIterator, 31, 110, driver, src, fis, wb, sheet1, fout, "Enterprise Value (Cr.)", "Mar 18",
					rowSize, colSize, firstPart, secondPart, colFirstPart, colSecondPart);
			ratioValues(stockIterator, 32, 115, driver, src, fis, wb, sheet1, fout, "EV/Net Operating Revenue (X)",
					"Mar 18", rowSize, colSize, firstPart, secondPart, colFirstPart, colSecondPart);
			ratioValues(stockIterator, 33, 120, driver, src, fis, wb, sheet1, fout, "EV/EBITDA (X)", "Mar 18", rowSize,
					colSize, firstPart, secondPart, colFirstPart, colSecondPart);
			ratioValues(stockIterator, 34, 125, driver, src, fis, wb, sheet1, fout,
					"MarketCap/Net Operating Revenue (X)", "Mar 18", rowSize, colSize, firstPart, secondPart,
					colFirstPart, colSecondPart);
			ratioValues(stockIterator, 35, 130, driver, src, fis, wb, sheet1, fout, "Price/BV (X)", "Mar 18", rowSize,
					colSize, firstPart, secondPart, colFirstPart, colSecondPart);
			ratioValues(stockIterator, 36, 135, driver, src, fis, wb, sheet1, fout, "Price/Net Operating Revenue",
					"Mar 18", rowSize, colSize, firstPart, secondPart, colFirstPart, colSecondPart);
			ratioValues(stockIterator, 37, 140, driver, src, fis, wb, sheet1, fout, "Earnings Yield", "Mar 18", rowSize,
					colSize, firstPart, secondPart, colFirstPart, colSecondPart);

			// missed one
			ratioValues(stockIterator, 9, 145, driver, src, fis, wb, sheet1, fout,
					"Book Value [ExclRevalReserve]/Share (Rs.)", "Mar 18", rowSize, colSize, firstPart, secondPart,
					colFirstPart, colSecondPart);
		} // end of for

	}// end of main

	public static void ratioValues(int stockIterator, int rowLineItemNumber, int writeCell, WebDriver driver, File src,
			FileInputStream fis, HSSFWorkbook wb, HSSFSheet sheet1, FileOutputStream fout, String lineItem,
			String codeYear, int rowSize, int colSize, String firstPart, String secondPart, String colFirstPart,
			String colSecondPart) throws Exception {

		sheet1.getRow(stockIterator).createCell(writeCell)
				.setCellValue(driver.findElement(By.xpath(
						"//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr["
								+ rowLineItemNumber + "]/td[2]"))
						.getText());// 18
		sheet1.getRow(stockIterator).createCell(writeCell + 1)
				.setCellValue(driver.findElement(By.xpath(
						"//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr["
								+ rowLineItemNumber + "]/td[3]"))
						.getText());// 17
		sheet1.getRow(stockIterator).createCell(writeCell + 2)
				.setCellValue(driver.findElement(By.xpath(
						"//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr["
								+ rowLineItemNumber + "]/td[4]"))
						.getText());// 16
		sheet1.getRow(stockIterator).createCell(writeCell + 3)
				.setCellValue(driver.findElement(By.xpath(
						"//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr["
								+ rowLineItemNumber + "]/td[5]"))
						.getText());// 15

		try {
			sheet1.getRow(stockIterator).createCell(writeCell + 4).setCellValue(driver.findElement(By.xpath(
					"//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr["
							+ rowLineItemNumber + "]/td[6]"))
					.getText());// 14
		} catch (Exception mar14NotThereInWebsite) {
			sheet1.getRow(stockIterator).createCell(writeCell + 4).setCellValue("NA");// 14
		}

		fout = new FileOutputStream("D:\\Users\\sanooj\\Desktop\\RefinedRatios.xls");
		wb.write(fout);

	}// end of ratio function

}// end of class
