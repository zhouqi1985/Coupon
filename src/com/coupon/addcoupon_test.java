package com.coupon;

import java.io.File;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.firefox.FirefoxProfile;
import org.openqa.selenium.support.ui.Select;

import readExcel.readExcel;





//import com.module.Module;

public class addcoupon_test {

	public static void main(String[] args) throws Exception {
		int count=0;
		FirefoxProfile pro =new FirefoxProfile(new File("C:\\Users\\lixuerui\\AppData\\Roaming\\Mozilla\\Firefox\\Profiles\\81xof17x.default"));
		WebDriver driver =new FirefoxDriver(pro);
		driver.get("http://a.zygames.com/");
		Thread.sleep(1000);
		driver.findElement(By.id("LoginName")).clear();
		driver.findElement(By.id("LoginName")).sendKeys("admin");
		driver.findElement(By.id("LoginPassword")).clear();
		driver.findElement(By.id("LoginPassword")).sendKeys("123.123a");
		driver.findElement(By.className("login_button")).click();
		Thread.sleep(1000);
		//Module mod =new Module();
		//mod.getModule(driver, "�һ�ȯϵͳ", "�������");
		driver.get("http://a.zygames.com/CouponV2/AddPacket");
		readExcel read1 =new readExcel();
		read1.getExcel("E:\\Data\\coupon.xls", "Ť�����");    //�����Ϣ
		for(int i=1;i<read1.rows;i++){
			count++;
			String PacketName=read1.readColumn("�������", i).trim();
			String Gold=read1.readColumn("���", i);
			String BindCash=read1.readColumn("���", i);
			String Gameid=read1.readColumn("��ϷID", i);
			String Classid=read1.readColumn("����ID", i);
			String Typeid=read1.readColumn("�����;", i);
			String PacketDescription=read1.readColumn("���˵��", i);
			String Title=read1.readColumn("�ʼ�����", i);
			String Message=read1.readColumn("�ʼ�����", i);
			String GiverName=read1.readColumn("������", i);
			String DeleDate=read1.readColumn("�����������", i);
			if(PacketName.equals("")) break;
			//Thread.sleep(2000);
			//driver.findElement(By.xpath(".//div[@id='page']/table/tbody/tr/td[1]/input[1]")).click();
			Thread.sleep(2000);
			driver.findElement(By.id("PacketName")).sendKeys(PacketName);
			driver.findElement(By.id("Gold")).clear();
			driver.findElement(By.id("Gold")).sendKeys(Gold);
			driver.findElement(By.id("BindCash")).clear();
			driver.findElement(By.id("BindCash")).sendKeys(BindCash);
		
//			readExcel read2 =new readExcel();
//			read2.getExcel("D:\\Data\\coupon.xls", "���Ը���");          //������Ϣ
//			for(int j=1;j<read2.rows;j++){
//				String title =read2.readColumn("����", j);
//				if(title.equals("")) break;
//				if(title.equals(PacketName)){
//					for(int n=1;n<read2.columns;n++){
//						if(read2.readColumn("���"+n,j).equals("")||read2.readColumn("���"+n, j).equals("N"))  break;
//						String itemName =read2.readColumn("����"+n, j).trim();
//						String Num =read2.readColumn("����"+n,j);
//						String Time =read2.readColumn("ʱ��"+n, j);
//						String type=read2.readColumn("���"+n, j);
//						addAnnex annex =new addAnnex();
//						Thread.sleep(1000);
//						annex.getAnnex(driver, itemName, Num, Time, type);   //��Ӹ���
//					}
//				}
//			}
			 Thread.sleep(1000);
	         Select GameID2= new Select(driver.findElement(By.id("GameID2")));
	          GameID2.selectByVisibleText(Gameid);
	          Thread.sleep(1000);
	          Select TypeID= new Select(driver.findElement(By.id("TypeID")));
	          TypeID.selectByVisibleText(Typeid);
	          driver.findElement(By.id("PacketDescription")).sendKeys(PacketDescription);
	          driver.findElement(By.id("Title")).sendKeys(Title);
	          driver.findElement(By.id("Message")).sendKeys(Message);
	          driver.findElement(By.id("GiverName")).sendKeys(GiverName);
	          Select ClassID= new Select(driver.findElement(By.id("ClassID")));
	          ClassID.selectByVisibleText(Classid);
	          Thread.sleep(1000);
	          driver.findElement(By.id("DeleDate")).clear();
	          driver.findElement(By.id("DeleDate")).sendKeys(DeleDate);
	          driver.findElement(By.cssSelector("input[value=\"�ύ \"]")).click();
	          Thread.sleep(1000);
	          System.out.println(count + ".�Ƿ������ɹ���"+driver.findElement(By.className("red")).getText().equals("�ɹ���"));
	  		driver.get("http://a.zygames.com/CouponV2/AddPacket");
	          Thread.sleep(2000);
		}

	}

}
