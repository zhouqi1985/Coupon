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
		mod.getModule(driver, "�һ�ȯϵͳ", "���ɶһ�ȯ");
		readExcel read =new readExcel();
		read.getExcel("D:\\Data\\scduihuanjuan.xls", "���ɶһ���");
		for(int i=1;i<read.rows;i++){
			count++;
			Thread.sleep(1000);
			Select GameID =new Select(driver.findElement(By.id("GameID")));
			GameID.selectByVisibleText(read.readColumn("��Ϸ", i));
			Thread.sleep(1500);
			Select ClassID =new Select(driver.findElement(By.id("ClassID")));
			ClassID.selectByVisibleText(read.readColumn("����", i));
			Thread.sleep(1500);
			Select PacketID =new Select(driver.findElement(By.id("PacketID")));
			PacketID.selectByVisibleText(read.readColumn("���", i));
			Thread.sleep(1500);
			Select Site =new Select(driver.findElement(By.id("Site")));
			Site.selectByVisibleText(read.readColumn("����", i));
			Thread.sleep(1500);
			Select DistributerID =new Select(driver.findElement(By.id("DistributerID")));
			DistributerID.selectByVisibleText(read.readColumn("������", i));
			if(read.readColumn("�Ƿ�����ʱ��", i).equals("��")){
				driver.findElement(By.id("IsUseDate")).click();
				Thread.sleep(500);
				driver.findElement(By.id("StartTime")).clear();
				driver.findElement(By.id("StartTime")).sendKeys(read.readColumn("��ʼʱ��", i));
				Thread.sleep(500);
				driver.findElement(By.id("EndTime")).clear();
				driver.findElement(By.id("EndTime")).sendKeys(read.readColumn("����ʱ��", i));
				if(read.readColumn("�Ƿ�UID��", i).equals("��")){
					driver.findElement(By.id("IsBindUserID")).click();
					Thread.sleep(500);
					driver.findElement(By.id("CouponLength")).clear();
					driver.findElement(By.id("CouponLength")).sendKeys(read.readColumn("�һ�ȯ����", i));
					driver.findElement(By.id("userIDFilePath")).sendKeys(read.readColumn("UserID�б�", i));
					Thread.sleep(500);
					driver.findElement(By.cssSelector("input[value=\"�ύ \"]")).click();
					Thread.sleep(1000);
					System.out.println(count+"."+read.readColumn("���", i)+" �Ƿ����ɶһ�ȯ�ɹ���"+driver.findElement(By.className("red")).getText().trim().equals("�ɹ���"));
					driver.findElement(By.className("link-clew")).click();
					Thread.sleep(1000);
					mod.getModule(driver, "�һ�ȯϵͳ", "���ɶһ�ȯ");
				}else {
					driver.findElement(By.id("Count")).sendKeys(read.readColumn("����", i));
					driver.findElement(By.id("CouponLength")).sendKeys(read.readColumn("�һ�ȯ����", i));
					if(read.readColumn("�Ƿ��������˺�ʹ��", i).equals("��")){
						driver.findElement(By.id("IsRepeatable")).click();
					}
					Thread.sleep(500);
					driver.findElement(By.cssSelector("input[value=\"�ύ \"]")).click();
					Thread.sleep(1000);
					System.out.println(count+"."+read.readColumn("���", i)+" �Ƿ����ɶһ�ȯ�ɹ���"+driver.findElement(By.className("red")).getText().trim().equals("�ɹ���"));
					driver.findElement(By.className("link-clew")).click();
					Thread.sleep(1000);
					mod.getModule(driver, "�һ�ȯϵͳ", "���ɶһ�ȯ");
				}
			}
			
		}

	}

}
