package com.coupon;

import java.io.File;
import java.util.List;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.firefox.FirefoxProfile;

import readExcel.readExcel;

import com.judge.Judge;
import com.module.Module;

public class cheakTime {

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
		mod.getModule(driver, "�һ�ȯϵͳ", "�������");
		readExcel read =new readExcel();
		String path ="D:\\Data\\coupon.xls";
		String sheetname ="cheakTime";
		read.getExcel(path, sheetname);
		for(int i=1;i<read.rows;i++){
			count++;
			String PacketName=read.readColumn("PacketName", i);
			if(PacketName.equals("")) break;
			Thread.sleep(1000);
			driver.findElement(By.id("PacketName")).clear();
			driver.findElement(By.id("PacketName")).sendKeys(PacketName);
			driver.findElement(By.className("bt-samll")).click();
			Thread.sleep(1500);
			List<WebElement> rows =driver.findElements(By.xpath("//*[@id='data']/table/tbody/tr"));
			if(rows.size()>1){
				for(int j=2;j<=rows.size();j++){
					if(driver.findElement(By.xpath("//*[@id='data']/table/tbody/tr["+j+"]/td[2]")).getText().trim().equals(PacketName)){
						driver.findElement(By.xpath("//*[@id='data']/table/tbody/tr["+j+"]/td[9]/input[3]")).click();
						break;
					}
				}
			}
			Thread.sleep(1500);
			driver.findElement(By.id("PacketName")).clear();
			driver.findElement(By.id("PacketName")).sendKeys(read.readColumn("PacketName1", i)); //�޸��������
			driver.findElement(By.id("Title")).clear();
			driver.findElement(By.id("Title")).sendKeys(read.readColumn("Title", i));
			driver.findElement(By.id("Message")).clear();
			driver.findElement(By.id("Message")).sendKeys(read.readColumn("Message", i));  //�޸��������
//			driver.findElement(By.id("DeleDate")).clear();
//			driver.findElement(By.id("DeleDate")).sendKeys(read.readColumn("DeleDate", i));  //�޸����ɾ��ʱ��
			Thread.sleep(1000);
            driver.findElement(By.cssSelector("input[value=\"�ύ \"]")).click();
	          Judge jud =new Judge();
	          while(jud.dealPotentialAlert(driver, true)){
	        	  jud.dealPotentialAlert(driver, true);
	          }
            Thread.sleep(1000);
            System.out.println(count + "."+read.readColumn("PacketName", i)+"�Ƿ������ɹ���"+driver.findElement(By.className("red")).getText().equals("�ɹ���"));
            driver.findElement(By.className("link-clew")).click();
            Thread.sleep(1000);
		}

	}

}
