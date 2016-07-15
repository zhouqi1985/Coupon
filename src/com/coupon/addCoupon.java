package com.coupon;

import java.io.File;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.firefox.FirefoxProfile;
import org.openqa.selenium.support.ui.Select;

import readExcel.readExcel;





import com.annex.addAnnex;
import com.judge.Judge;
import com.module.Module;

public class addCoupon {

	public static void main(String[] args) throws Exception {
		int count=0;
		FirefoxProfile pro =new FirefoxProfile(new File("C:\\Users\\zhouqi\\AppData\\Roaming\\Mozilla\\Firefox\\Profiles\\vtco59ig.default"));
		//这是指的是你需要的使用的浏览器的保存设置，注意这个版本要与此eclipse兼容
		WebDriver driver =new FirefoxDriver(pro);
		driver.get("http://a.zygames.com/");//你要打开的地址
		Thread.sleep(1000);//间隔操作响应的时间
		driver.findElement(By.id("LoginName")).clear();
		driver.findElement(By.id("LoginName")).sendKeys("huhongbin");//登陆页面的用户名
		driver.findElement(By.id("LoginPassword")).clear();
		driver.findElement(By.id("LoginPassword")).sendKeys("123.123.123a"); //登陆此页面的密码
		driver.findElement(By.className("login_button")).click();
		Thread.sleep(1000);
		Module mod =new Module();
		mod.getModule(driver, "兑换券系统", "礼包管理");//目标要打开的位置
		readExcel read1 =new readExcel();
		read1.getExcel("E:\\Data\\coupon.xls", "积分兑换"); //礼包信息及礼包的源路径
		for(int i=1;i<read1.rows;i++){
			count++;
			String PacketName=read1.readColumn("礼包名称", i).trim();//一 一与你后台模块上元素对应的内容
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
			driver.findElement(By.cssSelector("input[value=\"增加礼包\"]")).click();//获取这个按钮元素
			Thread.sleep(2000);
			driver.findElement(By.id("PacketName")).sendKeys(PacketName);;//元素在后台中英文注解
			driver.findElement(By.id("Gold")).clear();
			driver.findElement(By.id("Gold")).sendKeys(Gold);
			driver.findElement(By.id("BindCash")).clear();
			driver.findElement(By.id("BindCash")).sendKeys(BindCash);
			readExcel read2 =new readExcel();
			read2.getExcel("E:\\Data\\coupon.xls", "积分附件");          //附件信息
			for(int j=1;j<=read2.rows;j++){
				String title =read2.readColumn("标题", j);
				if(title.equals(PacketName)){
					for(int n=1;n<read2.columns;n++){
						if(read2.readColumn("类别"+n,j).equals("")||read2.readColumn("类别"+n, j).equals("N"))  break;
						String itemName =read2.readColumn("附件"+n, j).trim();
						String Num =read2.readColumn("数量"+n,j);
						String Time =read2.readColumn("时间"+n, j);						
						String type=read2.readColumn("类别"+n, j);
						Thread.sleep(1000);
						addAnnex annex =new addAnnex();
						annex.getAnnex(driver, itemName, Num, Time, type,Gameid);  //添加附件
					}
					break;
				}
			}
			 Thread.sleep(1000);
	         Select GameID2= new Select(driver.findElement(By.id("GameID2")));
	          GameID2.selectByVisibleText(Gameid);
	          Thread.sleep(1500);
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
	          Judge jud =new Judge();
	          while(jud.dealPotentialAlert(driver, true)){
	        	  jud.dealPotentialAlert(driver, true);
	          }
	          Thread.sleep(1000);
	          boolean red =driver.findElement(By.className("red")).getText().equals("成功！");
	          System.out.println(count +"."+PacketName+".是否新增成功："+red);
	          driver.findElement(By.className("link-clew")).click();
	          Thread.sleep(1000);
	          if(!red){
		          mod.repeatModule(driver, "兑换券系统", "礼包管理");
	          }
		}

	}

}
