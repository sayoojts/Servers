package wp.st.com.wp.st.com;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.concurrent.TimeUnit;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.openqa.selenium.By;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;


public class MissingRatios {
	
	public static void main(String[] args) throws Exception {
		
		Calendar cal = Calendar.getInstance();
		SimpleDateFormat sdf = new SimpleDateFormat("HH:mm:ss");
		System.out.println("Program started at " + sdf.format(cal.getTime()));

		
		File src = new File("D:\\Users\\sanooj\\Desktop\\ratios.xls");
		FileInputStream fis = new FileInputStream(src);
		HSSFWorkbook wb = new HSSFWorkbook(fis);
		HSSFSheet sheet2 = wb.getSheetAt(1);
		FileOutputStream fout;
		

		String balanceSheetUrl;
		String ratioUrl;
		String bEps18;
		String bEps17;
		String bEps16;
		String bEps15;
		String bEps14;
		String bEps13;
		String bEps12;
		String bEps11;
		String bEps10;
		
		String bookValue18;
		String bookValue17;
		String bookValue16;
		String bookValue15;
		String bookValue14;
		String bookValue13;
		String bookValue12;
		String bookValue11;
		String bookValue10;
		
		String dividentPerShare18;
		String dividentPerShare17;
		String dividentPerShare16;
		String dividentPerShare15;
		String dividentPerShare14;
		String dividentPerShare13;
		String dividentPerShare12;
		String dividentPerShare11;
		String dividentPerShare10;
		
		String rvnOpsPerShare18;
		String rvnOpsPerShare17;
		String rvnOpsPerShare16;
		String rvnOpsPerShare15;
		String rvnOpsPerShare14;
		String rvnOpsPerShare13;
		String rvnOpsPerShare12;
		String rvnOpsPerShare11;
		String rvnOpsPerShare10;
		
		String PBDITPerShare18;
		String PBDITPerShare17;
		String PBDITPerShare16;
		String PBDITPerShare15;
		String PBDITPerShare14;
		String PBDITPerShare13;
		String PBDITPerShare12;
		String PBDITPerShare11;
		String PBDITPerShare10;
		
		String netProfitPerShare18;
		String netProfitPerShare17;
		String netProfitPerShare16;
		String netProfitPerShare15;
		String netProfitPerShare14;
		String netProfitPerShare13;
		String netProfitPerShare12;
		String netProfitPerShare11;
		String netProfitPerShare10;
		
		String roce18;
		String roce17;
		String roce16;
		String roce15;
		String roce14;
		String roce13;
		String roce12;
		String roce11;
		String roce10;
		
		String debtByEquity18;
		String debtByEquity17;
		String debtByEquity16;
		String debtByEquity15;
		String debtByEquity14;
		String debtByEquity13;
		String debtByEquity12;
		String debtByEquity11;
		String debtByEquity10;


		System.setProperty("webdriver.chrome.driver",
				"D:\\Users\\sanooj\\workspace\\Cucumber\\src\\main\\resources\\drivers\\chromedriver.exe");
		WebDriver driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(1, TimeUnit.SECONDS);
		
		
		

		for (int j = 1; j <= 1919; j++) {
			
			try {
				sheet2.getRow(j).getCell(2).getStringCellValue();
				continue;
			}catch(Exception blankRatio) {
			
			System.out.println("The value of j is "+j);
			
			balanceSheetUrl = sheet2.getRow(j).getCell(1).getStringCellValue();
			
			
			
			ratioUrl = balanceSheetUrl.replaceAll("balance-sheetVI", "ratiosVI");
			System.out.println("The ratio url is "+ratioUrl);
			
			Thread.sleep(1500);
			
			driver.get(ratioUrl);
			
			
			
			try {//if the pop up accors
				driver.findElement(By.xpath("//*[@id='mcalertDiv']/div/div[1]/div[1]/a/img")).click();
			}catch(Exception noSuchPopup) {
				//do nothing
			}
			
			
			//basic eps first page
			bEps18 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[6]/td[3]")).getText();
			bEps17 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[6]/td[4]")).getText();
			bEps16 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[6]/td[5]")).getText();
			bEps15 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[6]/td[6]")).getText();
			
			
			System.out.println(bEps18+" "+ bEps17+" "+ bEps16+" "+ bEps15);
			
			//book value first page
			bookValue18 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[9]/td[3]")).getText();
			bookValue17 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[9]/td[4]")).getText();
			bookValue16 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[9]/td[5]")).getText();
			bookValue15 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[9]/td[6]")).getText();
			
			System.out.println(bookValue18+" "+ bookValue17+" "+ bookValue16+" "+ bookValue15);

			//dividend per share page 1
			dividentPerShare18 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[11]/td[3]")).getText();
			dividentPerShare17 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[11]/td[4]")).getText();
			dividentPerShare16 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[11]/td[5]")).getText();
			dividentPerShare15 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[11]/td[6]")).getText();
			
			System.out.println(dividentPerShare18+" "+ dividentPerShare17+" "+ dividentPerShare16+" "+ dividentPerShare15);
			
			//revenue from operations per share
			rvnOpsPerShare18 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[12]/td[3]")).getText();
			rvnOpsPerShare17 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[12]/td[4]")).getText();
			rvnOpsPerShare16 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[12]/td[5]")).getText();
			rvnOpsPerShare15 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[12]/td[6]")).getText();
		
			System.out.println(rvnOpsPerShare18+" "+ rvnOpsPerShare17+" "+ rvnOpsPerShare16+" "+ rvnOpsPerShare15);
			
			//PBDIT per share
			PBDITPerShare18 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[13]/td[3]")).getText();
			PBDITPerShare17 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[13]/td[4]")).getText();
			PBDITPerShare16 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[13]/td[5]")).getText();
			PBDITPerShare15 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[13]/td[6]")).getText();
		
			System.out.println(PBDITPerShare18+" "+ PBDITPerShare17+" "+ PBDITPerShare16+" "+ PBDITPerShare15);
			
			//net Profit Per Share
			netProfitPerShare18 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[16]/td[3]")).getText();
			netProfitPerShare17 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[16]/td[4]")).getText();
			netProfitPerShare16 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[16]/td[5]")).getText();
			netProfitPerShare15 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[16]/td[6]")).getText();
		
			System.out.println(netProfitPerShare18+" "+ netProfitPerShare17+" "+ netProfitPerShare16+" "+ netProfitPerShare15);

			//return on capital employed
			roce18 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[23]/td[3]")).getText();
			roce17 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[23]/td[4]")).getText();
			roce16 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[23]/td[5]")).getText();
			roce15 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[23]/td[6]")).getText();
		
			System.out.println(roce18+" "+ roce17+" "+ roce16+" "+ roce15);
			
			//total debt by equity
			debtByEquity18 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[25]/td[3]")).getText();
			debtByEquity17 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[25]/td[4]")).getText();
			debtByEquity16 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[25]/td[5]")).getText();
			debtByEquity15 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[25]/td[6]")).getText();
		
			System.out.println(debtByEquity18+" "+ debtByEquity17+" "+ debtByEquity16+" "+ debtByEquity15);
			
			//move to page2
			driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/div/div[1]/div[1]/table/tbody/tr/td/a/b")).click();
			
			//basic eps page 2
			bEps14 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[6]/td[2]")).getText();
			bEps13 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[6]/td[3]")).getText();
			bEps12 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[6]/td[4]")).getText();
			bEps11 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[6]/td[5]")).getText();
			bEps10 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[6]/td[6]")).getText();
			
			//book value page2
			bookValue14 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[9]/td[2]")).getText();
			bookValue13 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[9]/td[3]")).getText();
			bookValue12 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[9]/td[4]")).getText();
			bookValue11 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[9]/td[5]")).getText();
			bookValue10 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[9]/td[6]")).getText();
			
			
			System.out.println(bEps14+" "+ bEps13+" "+ bEps12+" "+ bEps11+" "+ bEps10);
			
			//dividend per share page 2
			dividentPerShare14 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[11]/td[2]")).getText();
			dividentPerShare13 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[11]/td[3]")).getText();
			dividentPerShare12 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[11]/td[4]")).getText();
			dividentPerShare11 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[11]/td[5]")).getText();
			dividentPerShare10 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[11]/td[6]")).getText();
			
			System.out.println(dividentPerShare14+" "+ dividentPerShare13+" "+ dividentPerShare12+" "+ dividentPerShare11+" "+ dividentPerShare10);
			
			//revenue from operations per share page 2
			rvnOpsPerShare14 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[12]/td[2]")).getText();
			rvnOpsPerShare13 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[12]/td[3]")).getText();
			rvnOpsPerShare12 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[12]/td[4]")).getText();
			rvnOpsPerShare11 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[12]/td[5]")).getText();
			rvnOpsPerShare10 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[12]/td[6]")).getText();
			
			System.out.println(rvnOpsPerShare14+" "+ rvnOpsPerShare13+" "+ rvnOpsPerShare12+" "+ rvnOpsPerShare11+" "+ rvnOpsPerShare10);
			
			//PBDIT per share page 2
			PBDITPerShare14 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[13]/td[2]")).getText();
			PBDITPerShare13 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[13]/td[3]")).getText();
			PBDITPerShare12 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[13]/td[4]")).getText();
			PBDITPerShare11 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[13]/td[5]")).getText();
			PBDITPerShare10 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[13]/td[6]")).getText();
			
			System.out.println(PBDITPerShare14+" "+ PBDITPerShare13+" "+ PBDITPerShare12+" "+ PBDITPerShare11+" "+ PBDITPerShare10);
			
			//net Profit Per Share page 2
			netProfitPerShare14 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[16]/td[2]")).getText();
			netProfitPerShare13 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[16]/td[3]")).getText();
			netProfitPerShare12 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[16]/td[4]")).getText();
			netProfitPerShare11 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[16]/td[5]")).getText();
			netProfitPerShare10 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[16]/td[6]")).getText();
			
			System.out.println(netProfitPerShare14+" "+ netProfitPerShare13+" "+ netProfitPerShare12+" "+ netProfitPerShare11+" "+ netProfitPerShare10);

			//return on capital employed page 2
			roce14 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[23]/td[2]")).getText();
			roce13 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[23]/td[3]")).getText();
			roce12 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[23]/td[4]")).getText();
			roce11 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[23]/td[5]")).getText();
			roce10 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[23]/td[6]")).getText();
			
			System.out.println(roce14+" "+ roce13+" "+ roce12+" "+ roce11+" "+ roce10);
			
			//total debt by equity page 2
			debtByEquity14 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[25]/td[2]")).getText();
			debtByEquity13 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[25]/td[3]")).getText();
			debtByEquity12 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[25]/td[4]")).getText();
			debtByEquity11 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[25]/td[5]")).getText();
			debtByEquity10 = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div[3]/div[2]/div[2]/div[2]/div[1]/table[2]/tbody/tr[25]/td[6]")).getText();
			
			System.out.println(debtByEquity14+" "+ debtByEquity13+" "+ debtByEquity12+" "+ debtByEquity11+" "+ debtByEquity10);
			
			//Enter eps
			sheet2.getRow(j).createCell(2).setCellValue(bEps18);
			sheet2.getRow(j).createCell(3).setCellValue(bEps17);
			sheet2.getRow(j).createCell(4).setCellValue(bEps16);
			sheet2.getRow(j).createCell(5).setCellValue(bEps15);
			sheet2.getRow(j).createCell(6).setCellValue(bEps14);
			sheet2.getRow(j).createCell(7).setCellValue(bEps13);
			sheet2.getRow(j).createCell(8).setCellValue(bEps12);
			sheet2.getRow(j).createCell(9).setCellValue(bEps11);
			sheet2.getRow(j).createCell(10).setCellValue(bEps10);
			
			//Enter book value
			sheet2.getRow(j).createCell(11).setCellValue(bookValue18);
			sheet2.getRow(j).createCell(12).setCellValue(bookValue17);
			sheet2.getRow(j).createCell(13).setCellValue(bookValue16);
			sheet2.getRow(j).createCell(14).setCellValue(bookValue15);
			sheet2.getRow(j).createCell(15).setCellValue(bookValue14);
			sheet2.getRow(j).createCell(16).setCellValue(bookValue13);
			sheet2.getRow(j).createCell(17).setCellValue(bookValue12);
			sheet2.getRow(j).createCell(18).setCellValue(bookValue11);
			sheet2.getRow(j).createCell(19).setCellValue(bookValue10);
			
			//Enter divident per share
			sheet2.getRow(j).createCell(20).setCellValue(dividentPerShare18);
			sheet2.getRow(j).createCell(21).setCellValue(dividentPerShare17);
			sheet2.getRow(j).createCell(22).setCellValue(dividentPerShare16);
			sheet2.getRow(j).createCell(23).setCellValue(dividentPerShare15);
			sheet2.getRow(j).createCell(24).setCellValue(dividentPerShare14);
			sheet2.getRow(j).createCell(25).setCellValue(dividentPerShare13);
			sheet2.getRow(j).createCell(26).setCellValue(dividentPerShare12);
			sheet2.getRow(j).createCell(27).setCellValue(dividentPerShare11);
			sheet2.getRow(j).createCell(28).setCellValue(dividentPerShare10);


			//Enter revenue from operations per share
			sheet2.getRow(j).createCell(29).setCellValue(rvnOpsPerShare18);
			sheet2.getRow(j).createCell(30).setCellValue(rvnOpsPerShare17);
			sheet2.getRow(j).createCell(31).setCellValue(rvnOpsPerShare16);
			sheet2.getRow(j).createCell(32).setCellValue(rvnOpsPerShare15);
			sheet2.getRow(j).createCell(33).setCellValue(rvnOpsPerShare14);
			sheet2.getRow(j).createCell(34).setCellValue(rvnOpsPerShare13);
			sheet2.getRow(j).createCell(35).setCellValue(rvnOpsPerShare12);
			sheet2.getRow(j).createCell(36).setCellValue(rvnOpsPerShare11);
			sheet2.getRow(j).createCell(37).setCellValue(rvnOpsPerShare10);
			
			//Enter PBDIT per share
			sheet2.getRow(j).createCell(38).setCellValue(PBDITPerShare18);
			sheet2.getRow(j).createCell(39).setCellValue(PBDITPerShare17);
			sheet2.getRow(j).createCell(40).setCellValue(PBDITPerShare16);
			sheet2.getRow(j).createCell(41).setCellValue(PBDITPerShare15);
			sheet2.getRow(j).createCell(42).setCellValue(PBDITPerShare14);
			sheet2.getRow(j).createCell(43).setCellValue(PBDITPerShare13);
			sheet2.getRow(j).createCell(44).setCellValue(PBDITPerShare12);
			sheet2.getRow(j).createCell(45).setCellValue(PBDITPerShare11);
			sheet2.getRow(j).createCell(46).setCellValue(PBDITPerShare10);
			
		    //Enter net Profit Per Share
			sheet2.getRow(j).createCell(47).setCellValue(netProfitPerShare18);
			sheet2.getRow(j).createCell(48).setCellValue(netProfitPerShare17);
			sheet2.getRow(j).createCell(49).setCellValue(netProfitPerShare16);
			sheet2.getRow(j).createCell(50).setCellValue(netProfitPerShare15);
			sheet2.getRow(j).createCell(51).setCellValue(netProfitPerShare14);
			sheet2.getRow(j).createCell(52).setCellValue(netProfitPerShare13);
			sheet2.getRow(j).createCell(53).setCellValue(netProfitPerShare12);
			sheet2.getRow(j).createCell(54).setCellValue(netProfitPerShare11);
			sheet2.getRow(j).createCell(55).setCellValue(netProfitPerShare10);
			
			//Enter return on capital employed
			sheet2.getRow(j).createCell(56).setCellValue(roce18);
			sheet2.getRow(j).createCell(57).setCellValue(roce17);
			sheet2.getRow(j).createCell(58).setCellValue(roce16);
			sheet2.getRow(j).createCell(59).setCellValue(roce15);
			sheet2.getRow(j).createCell(60).setCellValue(roce14);
			sheet2.getRow(j).createCell(61).setCellValue(roce13);
			sheet2.getRow(j).createCell(62).setCellValue(roce12);
			sheet2.getRow(j).createCell(63).setCellValue(roce11);
			sheet2.getRow(j).createCell(64).setCellValue(roce10);
			
			//Enter total debt by equity
			sheet2.getRow(j).createCell(65).setCellValue(debtByEquity18);
			sheet2.getRow(j).createCell(66).setCellValue(debtByEquity17);
			sheet2.getRow(j).createCell(67).setCellValue(debtByEquity16);
			sheet2.getRow(j).createCell(68).setCellValue(debtByEquity15);
			sheet2.getRow(j).createCell(69).setCellValue(debtByEquity14);
			sheet2.getRow(j).createCell(70).setCellValue(debtByEquity13);
			sheet2.getRow(j).createCell(71).setCellValue(debtByEquity12);
			sheet2.getRow(j).createCell(72).setCellValue(debtByEquity11);
			sheet2.getRow(j).createCell(73).setCellValue(debtByEquity10);
			
			
			
			fout = new FileOutputStream(
					"D:\\Users\\sanooj\\Desktop\\ratios.xls");
			wb.write(fout);
			
			}//end of first catch	
			
			
			
			
		}
		
		
		/*fout.flush();
		fout.close();
		fout = null;
	     System.gc();*/
		
	}

}
