package com.coupon;

import java.io.File;
import java.util.List;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.firefox.FirefoxProfile;
import org.openqa.selenium.support.ui.Select;

import readExcel.readExcel;

import com.module.Module;

public class signUpdate {


	public static void main(String[] args) throws Exception {
		int count =0;
		FirefoxProfile pro =new FirefoxProfile(new File("C:\\Users\\lixuerui\\AppData\\Roaming\\Mozilla\\Firefox\\Profiles\\81xof17x.default"));
		WebDriver driver =new FirefoxDriver(pro);
		driver.navigate().to("http://a.zygames.com/");
		Thread.sleep(1000);
		driver.findElement(By.id("LoginName")).clear();
		driver.findElement(By.id("LoginName")).sendKeys("huhongbin");
		driver.findElement(By.id("LoginPassword")).clear();
		driver.findElement(By.id("LoginPassword")).sendKeys("123.123.123a");
		driver.findElement(By.className("login_button")).click();
		Thread.sleep(1000);
		Module mod =new Module();
		mod.getModule(driver, "�����(DEV)", "(Q)ǩ��V2����");
		readExcel read =new readExcel();
		String path ="D:\\Data\\coupon.xls";
		String sheetname ="signcoupon";
		read.getExcel(path, sheetname);
		Thread.sleep(1000);
		driver.findElement(By.xpath("html/body/table/tbody/tr/td/a[3]")).click();
		for(int i=1;i<read.rows;i++){
			count++;
			Thread.sleep(1000);
			if(read.readColumn("�������", i).equals("")) break;
			driver.findElement(By.id("PacketName")).clear();
			driver.findElement(By.id("PacketName")).sendKeys(read.readColumn("�������", i));
			Thread.sleep(1500);
			Select EventID =new Select(driver.findElement(By.id("EventID")));
			EventID.selectByVisibleText(read.readColumn("�",1));
			driver.findElement(By.cssSelector("input[value=\"����\"]")).click();
			Thread.sleep(1500);
			List<WebElement> row =driver.findElements(By.xpath("//*[@id='data']/table/tbody/tr"));
			if(row.size()>1){
				for(int j = 2;j<=row.size();j++){
					if(driver.findElement(By.xpath("//*[@id='data']/table/tbody/tr["+j+"]/td[2]")).getText().trim().contains(read.readColumn("�������", i))){
						driver.findElement(By.xpath("//*[@id='data']/table/tbody/tr["+j+"]/td[8]/input[1]")).click();
						Thread.sleep(1500);
						break;
					}
				}
			}
			driver.findElement(By.id("SignCounts")).clear();
			driver.findElement(By.id("SignCounts")).sendKeys(read.readColumn("ǩ������", i));
			driver.findElement(By.id("button2")).click();
            Thread.sleep(1000);
            System.out.println(count +read.readColumn("�������", i)+".�Ƿ��޸ĳɹ���"+driver.findElement(By.className("red")).getText().equals("�ɹ���"));
            driver.findElement(By.className("link-clew")).click();
            Thread.sleep(1000);
	    }
	}

}
