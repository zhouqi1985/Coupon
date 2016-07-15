package com.coupon;

import java.io.File;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.firefox.FirefoxProfile;
import org.openqa.selenium.support.ui.Select;

import readExcel.readExcel;

import com.module.Module;

public class addDuihuanJuan {

	public static void main(String[] args) throws Exception {
		int count=0;
		FirefoxProfile pro =new FirefoxProfile(new File("C:\\Users\\lixuerui\\AppData\\Roaming\\Mozilla\\Firefox\\Profiles\\81xof17x.default"));
		WebDriver driver =new FirefoxDriver(pro);
		driver.get("http://a.zygames.com/");
		Thread.sleep(1000);
		driver.findElement(By.id("LoginName")).clear();
		driver.findElement(By.id("LoginName")).sendKeys("huhongbin");
		driver.findElement(By.id("LoginPassword")).clear();
		driver.findElement(By.id("LoginPassword")).sendKeys("123.123.123a");
		driver.findElement(By.className("login_button")).click();
		Thread.sleep(1000);
		Module mod =new Module();
		mod.getModule(driver, "兑换券系统", "生成兑换券");
		readExcel read =new readExcel();
		read.getExcel("D:\\Data\\scduihuanjuan.xls", "生成兑换劵");
		for(int i=1;i<read.rows;i++){
			count++;
			Thread.sleep(1000);
			Select GameID =new Select(driver.findElement(By.id("GameID")));
			GameID.selectByVisibleText(read.readColumn("游戏", i));
			Thread.sleep(1500);
			Select ClassID =new Select(driver.findElement(By.id("ClassID")));
			ClassID.selectByVisibleText(read.readColumn("分类", i));
			Thread.sleep(1500);
			Select PacketID =new Select(driver.findElement(By.id("PacketID")));
			PacketID.selectByVisibleText(read.readColumn("礼包", i));
			Thread.sleep(1500);
			Select Site =new Select(driver.findElement(By.id("Site")));
			Site.selectByVisibleText(read.readColumn("区服", i));
			Thread.sleep(1500);
			Select DistributerID =new Select(driver.findElement(By.id("DistributerID")));
			DistributerID.selectByVisibleText(read.readColumn("分销商", i));
			if(read.readColumn("是否启用时间", i).equals("是")){
				driver.findElement(By.id("IsUseDate")).click();
				Thread.sleep(500);
				driver.findElement(By.id("StartTime")).clear();
				driver.findElement(By.id("StartTime")).sendKeys(read.readColumn("开始时间", i));
				Thread.sleep(500);
				driver.findElement(By.id("EndTime")).clear();
				driver.findElement(By.id("EndTime")).sendKeys(read.readColumn("结束时间", i));
				if(read.readColumn("是否UID绑定", i).equals("是")){
					driver.findElement(By.id("IsBindUserID")).click();
					Thread.sleep(500);
					driver.findElement(By.id("CouponLength")).clear();
					driver.findElement(By.id("CouponLength")).sendKeys(read.readColumn("兑换券长度", i));
					driver.findElement(By.id("userIDFilePath")).sendKeys(read.readColumn("UserID列表", i));
					Thread.sleep(500);
					driver.findElement(By.cssSelector("input[value=\"提交 \"]")).click();
					Thread.sleep(1000);
					System.out.println(count+"."+read.readColumn("礼包", i)+" 是否生成兑换券成功："+driver.findElement(By.className("red")).getText().trim().equals("成功！"));
					driver.findElement(By.className("link-clew")).click();
					Thread.sleep(1000);
					mod.getModule(driver, "兑换券系统", "生成兑换券");
				}else {
					driver.findElement(By.id("Count")).sendKeys(read.readColumn("数量", i));
					driver.findElement(By.id("CouponLength")).sendKeys(read.readColumn("兑换券长度", i));
					if(read.readColumn("是否允许多个账号使用", i).equals("是")){
						driver.findElement(By.id("IsRepeatable")).click();
					}
					Thread.sleep(500);
					driver.findElement(By.cssSelector("input[value=\"提交 \"]")).click();
					Thread.sleep(1000);
					System.out.println(count+"."+read.readColumn("礼包", i)+" 是否生成兑换券成功："+driver.findElement(By.className("red")).getText().trim().equals("成功！"));
					driver.findElement(By.className("link-clew")).click();
					Thread.sleep(1000);
					mod.getModule(driver, "兑换券系统", "生成兑换券");
				}
			}
			
		}

	}

}
