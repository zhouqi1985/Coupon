package com.coupon;


import java.io.File;
import java.io.IOException;
import java.util.List;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.firefox.FirefoxProfile;
import org.openqa.selenium.support.ui.Select;

import jxl.Cell;
import jxl.JXLException;
import jxl.Sheet;
import jxl.Workbook;


public class coupon_X {

	public static void main(String[] args) throws JXLException, IOException, Exception {
		int count =0;
		//礼包基本信息数据在excel表Coupon的sheet：礼包基本信息
		Workbook wb =Workbook.getWorkbook(new File("D:\\Data\\coupon.xls"));
		Sheet st1 =wb.getSheet("世界杯投票礼包");
		//礼包基本信息数据在excel表Coupon的sheet：礼包配置
		Sheet st2=wb.getSheet("WC投票附件");
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
		//切换frame到左边模块
		driver.switchTo().frame("admin_left");
		//点击后台“兑换券系统”模块
		driver.findElement(By.xpath("//*[@id='red']/li[6]/span")).click(); 
		//点击礼包管理
		driver.findElement(By.xpath("//*[@id='red']/li[6]/ul/li[3]/span/a")).click();
		//切换到默认模块
		driver.switchTo().defaultContent();
		//切换到信息输入模块
		driver.switchTo().frame("admin_main");
		Thread.sleep(1000);
		for(int i=1;i<st1.getRows();i++){
			Cell[] cell=st1.getRow(i);
			count++;
			if(cell[1].getContents().equals("")) break;
			driver.findElement(By.className("bt-big")).click();
			Thread.sleep(1000);
			driver.findElement(By.id("PacketName")).sendKeys(cell[1].getContents());
			driver.findElement(By.id("Gold")).clear();
			driver.findElement(By.id("Gold")).sendKeys(cell[2].getContents());
			driver.findElement(By.id("BindCash")).clear();
			driver.findElement(By.id("BindCash")).sendKeys(cell[3].getContents());
			//添加附件
			for(int j=0;j<st2.getRows();j++){
				Cell[] cell2=st2.getRow(j);
				String itemName =cell2[3].getContents();
				int date =Integer.valueOf(cell2[4].getContents());
				String num =cell2[5].getContents();
				if(cell[1].getContents().equals(cell2[1].getContents())){
					driver.findElement(By.id("btnItemSelect")).click();
					Thread.sleep(1000);
					driver.findElement(By.id("ItemName")).clear();
					driver.findElement(By.id("ItemName")).sendKeys(itemName);
					driver.findElement(By.id("btSearch")).click();
					Thread.sleep(1500);
					List<WebElement> options =driver.findElements(By.cssSelector("table#itemListTable tr"));
					if(options.size()>1){
						for(int k=2;k<=options.size();k++){
							if(cell2[6].getContents().equals("服装")){
								 if(driver.findElement(By.xpath("//*[@id='itemListTable']/tbody/tr["+k+"]/td[3]")).getText().equals(itemName)&&driver.findElement(By.xpath("//*[@id='itemListTable']/tbody/tr["+k+"]/td[2]")).getText().contains("_")){
										driver.findElement(By.xpath("//*[@id='itemListTable']/tbody/tr["+k+"]/td[1]/input")).click();
						         }
							}else {
								if(driver.findElement(By.xpath("//*[@id='itemListTable']/tbody/tr["+k+"]/td[3]")).getText().equals(itemName)){
									driver.findElement(By.xpath("//*[@id='itemListTable']/tbody/tr["+k+"]/td[1]/input")).click();
								}
							}				
				      }		
				   }else{
					   driver.findElement(By.xpath("//html/body/div[6]/div[11]/button[1]")).click();
					   continue;
				   }
					driver.findElement(By.xpath("//html/body/div[6]/div[11]/button[2]")).click();
					Thread.sleep(200);
					List<WebElement> row =driver.findElements(By.cssSelector("tbody#selectedItemBody tr"));
					for(int h=1;h<=row.size();h++){
						if(driver.findElement(By.xpath("//*[@id='selectedItemBody']/tr[" + h + "]/td[2]")).getText().equals(itemName)){
							if(date>0){
								String time =String.valueOf(Integer.valueOf(cell2[4].getContents())*24);
								driver.findElement(By.xpath("//*[@id='selectedItemBody']/tr[" + h + "]/td[3]/input")).clear();
								driver.findElement(By.xpath("//*[@id='selectedItemBody']/tr[" + h + "]/td[3]/input")).sendKeys(time);
							}else {
								String time="-1";
								driver.findElement(By.xpath("//*[@id='selectedItemBody']/tr[" + h+ "]/td[3]/input")).clear();
								driver.findElement(By.xpath("//*[@id='selectedItemBody']/tr[" +h + "]/td[3]/input")).sendKeys(time);
							}
								driver.findElement(By.xpath("//*[@id='selectedItemBody']/tr["+h+"]/td[4]/input[2]")).clear();
			       				driver.findElement(By.xpath("//*[@id='selectedItemBody']/tr["+h+"]/td[4]/input[2]")).sendKeys(num);	
						}
					}
				}else{
					continue;
				}
			 }
			List<WebElement> row =driver.findElements(By.cssSelector("tbody#selectedItemBody tr"));
			for(int h=2;h<row.size();h++){
				if(!"".equals(driver.findElement(By.xpath("//*[@id='selectedItemBody']/tr[" + (h-1) + "]/td[2]")).getText())){
					if(driver.findElement(By.xpath("//*[@id='selectedItemBody']/tr[" + (h-1) + "]/td[2]")).getText().equals(driver.findElement(By.xpath("//*[@id='selectedItemBody']/tr[" +h + "]/td[2]")).getText())){
						driver.findElement(By.xpath("//*[@id='selectedItemBody']/tr[" + (h-1) + "]/td[5]/input")).click();
					}
				}else{
					continue;
				}
			}
            Select GameID2= new Select(driver.findElement(By.id("GameID2")));
            GameID2.selectByVisibleText(cell[4].getContents());
            Thread.sleep(1000);
            Select ClassID= new Select(driver.findElement(By.id("ClassID")));
            ClassID.selectByVisibleText(cell[5].getContents());
            Thread.sleep(1000);
            Select TypeID= new Select(driver.findElement(By.id("TypeID")));
            TypeID.selectByVisibleText(cell[6].getContents());
            driver.findElement(By.id("PacketDescription")).sendKeys(cell[7].getContents());
            driver.findElement(By.id("Title")).sendKeys(cell[8].getContents());
            driver.findElement(By.id("Message")).sendKeys(cell[11].getContents());
            driver.findElement(By.id("GiverName")).sendKeys(cell[9].getContents());
            driver.findElement(By.id("DeleDate")).clear();
            driver.findElement(By.id("DeleDate")).sendKeys(cell[10].getContents());
            driver.findElement(By.cssSelector("input[value=\"提交 \"]")).click();
            Thread.sleep(1000);
            System.out.println(count + ".是否新增成功："+driver.findElement(By.className("red")).getText().equals("成功！"));
            driver.findElement(By.className("link-clew")).click();
            Thread.sleep(1000);
			}
	       wb.close();
		}
	}

