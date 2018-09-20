package wp.st.com.wp.st.com;

import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class Xpath {
	public static void main(String[] args) throws Exception {
		
		System.setProperty("webdriver.chrome.driver",
				"D:\\Users\\sanooj\\workspace\\Cucumber\\src\\main\\resources\\drivers\\chromedriver.exe");
		WebDriver driver = new ChromeDriver();
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(1, TimeUnit.SECONDS);
		
		driver.get("http://www.moneycontrol.com/stocks/cptmarket/compsearchnew.php?search_data=&cid=&mbsearch_str=&topsearch_type=1&search_str=AKASH");
		Thread.sleep(2000);
		
		String stock = "AKASH";
		int size = driver.findElements(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div/table/tbody/tr")).size();
		
		System.out.println("the size is " +size);
		for(int i=1;i<=size;i++) {
			System.out.print("The "+i+"th element is ");
			//System.out.println(driver.findElement(By.xpath("//*[@id=\'mc_mainWrapper\']/div[3]/div[2]/div[1]/div["+i+"]")).getText());
			String nseidStockName = driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div/table/tbody/tr["+i+"]/td[2]/p/span[1]")).getText();
			System.out.println(driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div/table/tbody/tr["+i+"]/td[2]/p/span[1]")).getText());
			 if(nseidStockName.equalsIgnoreCase("NSE Id :"+stock)){
				 driver.findElement(By.xpath("//*[@id='mc_mainWrapper']/div[3]/div[2]/div/table/tbody/tr["+i+"]/td[2]/p/span[1]")).click();
				 break;
			 }
		}
		
		
		
		
	}

}
