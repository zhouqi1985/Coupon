package com.coupon;

import java.io.File;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.firefox.FirefoxProfile;
import org.openqa.selenium.support.ui.Select;

import readExcel.readExcel;

import com.annex.addAnnex;
import com.module.Module;

public class signCoupon {

	public static void main(String[] args) throws Exception {
		int count=0;
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
		readExcel read1 =new readExcel();
		read1.getExcel("D:\\Data\\coupon\\Sign_ZX.xls", "签到(1-15)");    //礼包信息
		for(int i=1;i<read1.rows;i++){
			count++;
			String PacketName=read1.readColumn("礼包名称", i).trim();
			String Gold=read1.readColumn("金币", i);
			String BindCash=read1.readColumn("绑点", i);
			String Gameid=read1.readColumn("游戏ID", i);
			String Classid=read1.readColumn("分类ID", i);
			String Typeid=read1.readColumn("礼包用途", i);
			String PacketDescription=read1.readColumn("礼包说明", i);
			String Title=read1.readColumn("邮件标题", i);
			String Message=read1.readColumn("邮件正文", i);
			String GiverName=read1.readColumn("赠送人", i);
			String DeleDate=read1.readColumn("礼包到期日期", i);
			if(PacketName.equals("")) break;
			Thread.sleep(1000);
			driver.findElement(By.className("bt-big")).click();
			Thread.sleep(1000);
			driver.findElement(By.id("PacketName")).sendKeys(PacketName);
			driver.findElement(By.id("Gold")).clear();
			driver.findElement(By.id("Gold")).sendKeys(Gold);
			driver.findElement(By.id("BindCash")).clear();
			driver.findElement(By.id("BindCash")).sendKeys(BindCash);
			readExcel read2 =new readExcel();
			read2.getExcel("D:\\Data\\coupon\\Sign_ZX.xls", "礼包附件");          //附件信息
			for(int j=1;j<read2.rows;j++){
				String title =read2.readColumn("标题", j);
				if(title.equals(PacketName)){
					for(int n=1;n<read2.columns;n++){
						if(read2.readColumn("类别"+n,j).equals("")||read2.readColumn("类别"+n,j).equals("N")) break;
						String itemName =read2.readColumn("附件"+n, j).trim();
						String Num =read2.readColumn("数量"+n,j);
						String Time =read2.readColumn("时间"+n, j);
						String type=read2.readColumn("类别"+n, j);
						addAnnex annex =new addAnnex();
						Thread.sleep(1000);
						annex.getAnnex(driver, itemName, Num, Time, type,Gameid);   //添加附件
					}
				}
			}
			 Thread.sleep(1000);
	         Select GameID2= new Select(driver.findElement(By.id("GameID2")));
	          GameID2.selectByVisibleText(Gameid);
	          Thread.sleep(1000);
	          Select ClassID= new Select(driver.findElement(By.id("ClassID")));
	          ClassID.selectByVisibleText(Classid);
	          Thread.sleep(1000);
	          Select TypeID= new Select(driver.findElement(By.id("TypeID")));
	          TypeID.selectByVisibleText(Typeid);
	          driver.findElement(By.id("PacketDescription")).sendKeys(PacketDescription);
	          driver.findElement(By.id("Title")).sendKeys(Title);
	          driver.findElement(By.id("Message")).sendKeys(Message);
	          driver.findElement(By.id("GiverName")).sendKeys(GiverName);
	          driver.findElement(By.id("DeleDate")).clear();
	          driver.findElement(By.id("DeleDate")).sendKeys(DeleDate);
	          driver.findElement(By.cssSelector("input[value=\"提交 \"]")).click();
	          Thread.sleep(1000);
	          System.out.println(count + ".是否新增成功："+driver.findElement(By.className("red")).getText().equals("成功！"));
	          driver.findElement(By.className("link-clew")).click();
	          Thread.sleep(1000);
		}

	}

}
