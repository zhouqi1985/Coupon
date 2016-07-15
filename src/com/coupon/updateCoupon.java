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

import com.annex.addAnnex;
import com.judge.Judge;
import com.module.Module;

public class updateCoupon {

	public static void main(String[] args) throws Exception {
		int count=0;
		FirefoxProfile pro =new FirefoxProfile(new File("C:\\Users\\zhouqi\\AppData\\Roaming\\Mozilla\\Firefox\\Profiles\\vtco59ig.default"));
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
		read.getExcel("E:\\Data\\coupon.xls", "updateCoupon");             //礼包信息
		for(int i=1;i<read.rows;i++){
			count++;
			if(read.readColumn("礼包名称", i).equals("")) break;
			Thread.sleep(1000);
			//String Packetname=read.readColumn("模糊搜索", i);
			String PacketName=read.readColumn("礼包名称", i);
			String Gameid=read.readColumn("分类", i);
			driver.findElement(By.id("PacketName")).clear();
			driver.findElement(By.id("PacketName")).sendKeys(PacketName);
			driver.findElement(By.cssSelector("input[value=\"搜索\"]")).click();
			Thread.sleep(1000);
			List<WebElement> row =driver.findElements(By.xpath("//*[@id='data']/table/tbody/tr"));
			if(row.size()>1){
				for(int j=2;j<=row.size();j++){
					if(driver.findElement(By.xpath("//*[@id='data']/table/tbody/tr["+j+"]/td[2]")).getText().trim().equals(PacketName)){
						driver.findElement(By.xpath("//*[@id='data']/table/tbody/tr["+j+"]/td[9]/input[3]")).click();
						break;
					}
				}
			}
			Thread.sleep(1500);
			List<WebElement> row1 =driver.findElements(By.xpath("//*[@id='selectedItemBody']/tr"));
			if(row1.size()>0){
				for(int h=row1.size();h>=1;h--){
					driver.findElement(By.xpath("//*[@id='selectedItemBody']/tr["+h+"]/td[5]/input")).click();
					Thread.sleep(100);
				}	
			}
			Thread.sleep(1000);
			readExcel read2 =new readExcel();
			read2.getExcel("E:\\Data\\coupon.xls", "updateCoupon附件");    
			//添加附件
			for(int m=1;m<read2.rows;m++){
				String title =read2.readColumn("标题", m);
				if(title.equals(PacketName)){
					for(int n=1;n<read2.columns;n++){
						if(read2.readColumn("类别"+n,m).equals("")||read2.readColumn("类别"+n, m).equals("N"))  break;
						String itemName =read2.readColumn("附件"+n, m).trim();
						String Num =read2.readColumn("数量"+n,m);
						String Time =read2.readColumn("时间"+n, m);						
						String type=read2.readColumn("类别"+n, m);
						addAnnex annex=new addAnnex();
						annex.getAnnex(driver, itemName, Num, Time, type,Gameid);  //添加附件
					}
					break;
				}
			}

			Thread.sleep(500);
			//driver.findElement(By.id("Message")).clear();
			//driver.findElement(By.id("Message")).sendKeys(read.readColumn("邮件正文", i));
	        driver.findElement(By.cssSelector("input[value=\"提交 \"]")).click();
	        Judge jud =new Judge();
	          while(jud.dealPotentialAlert(driver, true)){
	        	  jud.dealPotentialAlert(driver, true);
	          }
	          Thread.sleep(1000);
	          System.out.println(count +read.readColumn("礼包名称", i)+".是否更新成功："+driver.findElement(By.className("red")).getText().equals("成功！"));
	          driver.findElement(By.className("link-clew")).click();
	          Thread.sleep(1000);
			
		}

	}

}
