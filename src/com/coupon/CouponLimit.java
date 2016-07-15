package com.coupon;

import java.io.File;
import java.util.List;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.firefox.FirefoxProfile;

import readExcel.readExcel;

import com.module.Module;

public class CouponLimit {

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
		mod.getModule(driver, "兑换券系统", "礼包管理");
		readExcel read =new readExcel();
		String path ="D:\\Data\\coupon.xls";
		String sheetname ="Limit";
		read.getExcel(path, sheetname);
		for(int i=1;i<read.rows;i++){
			count++;
			Thread.sleep(1000);
			if(read.readColumn("礼包名称", i).equals("")) break;
			driver.findElement(By.id("PacketName")).clear();
			driver.findElement(By.id("PacketName")).sendKeys(read.readColumn("礼包名称", i));
			driver.findElement(By.className("bt-samll")).click();
			Thread.sleep(1500);
			List<WebElement> row =driver.findElements(By.xpath("//*[@id='data']/table/tbody/tr"));
			if(row.size()>1){
				for(int j=2;j<=row.size();j++){
					if(driver.findElement(By.xpath("//*[@id='data']/table/tbody/tr["+j+"]/td[2]")).getText().trim().contains(read.readColumn("礼包名称", i))){
						if(!read.readColumn("日限制", i).equals("")){
							driver.findElement(By.xpath("//*[@id='data']/table/tbody/tr["+j+"]/td[9]/input[1]")).click();
						}
						if(!read.readColumn("总限制", i).equals("")){
							driver.findElement(By.xpath("//*[@id='data']/table/tbody/tr["+j+"]/td[9]/input[2]")).click();
						}
						break;
					}
				}
			}
			Thread.sleep(1500);
			if(!read.readColumn("日限制", i).equals("")){
				driver.findElement(By.id("DayCounts")).clear();
				driver.findElement(By.id("DayCounts")).sendKeys(read.readColumn("日限制", i));
			}
			if(!read.readColumn("总限制", i).equals("")){
				driver.findElement(By.id("TotalCounts")).clear();
				driver.findElement(By.id("TotalCounts")).sendKeys(read.readColumn("总限制", i));
			}
			driver.findElement(By.id("button2")).click();
            Thread.sleep(1000);
            System.out.println(count +read.readColumn("礼包名称", i)+".是否新增成功："+driver.findElement(By.className("red")).getText().equals("成功！"));
            driver.findElement(By.className("link-clew")).click();
            Thread.sleep(1000);
		}
	}
}
