package ApachePOI;

import org.testng.annotations.Test;
import org.testng.annotations.BeforeMethod;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.DataProvider;

public class Apache_POI {
	
	WebDriver driver;
	
  @Test(dataProvider = "data")
  public void f(String user, String pass) {
	  
	 System.out.println(user+":::"+pass); 
	  driver.get("https://mail.rediff.com/cgi-bin/login.cgi");
	  
	  driver.findElement(By.name("login")).sendKeys(user);
	  
	  driver.findElement(By.name("passwd")).sendKeys(pass);
	  
	  driver.findElement(By.name("proceed")).click();
	  
  }
  @BeforeMethod
  public void beforeMethod() {
	  
	  System.setProperty("webdriver.chrome.driver", "C:\\Apps\\ChromeDriver\\chromedriver.exe");
	  
	  driver=new ChromeDriver();
  }

  @AfterMethod
  public void afterMethod() {
	  
	  driver.close();
  }


  @DataProvider
  public Object[][] data() throws IOException {
	  
	  FileInputStream fp= new FileInputStream("C:\\Users\\prajktaudayku.more\\OneDrive - HCL Technologies Ltd\\Desktop\\Excel\\ApachePOIHSSFDATA.xls");
	  
	  HSSFWorkbook wb=new HSSFWorkbook(fp);
	  
	  HSSFSheet sheet=wb.getSheetAt(0);
	  int rowNum=sheet.getLastRowNum();
	  int cellNum=sheet.getRow(0).getLastCellNum();
	  
	  Object[][] obj=new Object[rowNum][cellNum];
	  
	  for (int i=0;i<rowNum;i++)
	  {
		 for(int j=0; j<cellNum;j++)
		 {
			 obj[i][j]=sheet.getRow(i+1).getCell(j).getStringCellValue();
			 
		 }
	  }
	  
	  wb.close();
	  
    return obj;
  }
}
