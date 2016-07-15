package com.coupon;

import java.io.File;
import java.util.List;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.firefox.FirefoxProfile;

import com.module.Module;

public class HFujian {

	public static void main(String[] args) throws Exception {
		int count =0;
		//礼包基本信息数据在excel表Coupon的sheet：礼包基本信息
		Workbook wb =Workbook.getWorkbook(new File("D:\\Data\\coupon.xls"));
		Sheet st1=wb.getSheet("添加附件1");
		Sheet st2=wb.getSheet("附件内容1");
		FirefoxProfile pro =new FirefoxProfile(new File("C:\\Users\\lixuerui\\AppData\\Roaming\\Mozilla\\Firefox\\Profiles\\81xof17x.default"));
		WebDriver driver =new FirefoxDriver(pro);
		driver.navigate().to("http://a.zygames.com/");
		Thread.sleep(1000);
		driver.findElement(By.id("LoginName")).clear();
		driver.findElement(By.id("LoginName")).sendKeys("admin");
		driver.findElement(By.id("LoginPassword")).clear();
		driver.findElement(By.id("LoginPassword")).sendKeys("123.123a");
		driver.findElement(By.className("login_button")).click();
		Thread.sleep(1000);
		Module mod =new Module();
		mod.getModule(driver, "兑换券系统V2.0", "礼包管理");
		Thread.sleep(1000);
		for(int i=1;i<st1.getRows();i++){
			Cell[] cell =st1.getRow(i);
			count++;
			String PacketName=cell[0].getContents();
			String PacketDes=cell[1].getContents();
			String Pname =cell[2].getContents();
			if(cell[0].getContents().equals("")) break;
			driver.findElement(By.id("PacketName")).clear();
			driver.findElement(By.id("PacketName")).sendKeys(Pname);
			driver.findElement(By.className("bt-samll")).click();
			Thread.sleep(1000);
			List<WebElement> rows =driver.findElements(By.xpath("//*[@id='data']/table/tbody/tr"));
			if(rows.size()>1){
				for(int j=2;j<=rows.size();j++){
					if(driver.findElement(By.xpath("//*[@id='data']/table/tbody/tr["+j+"]/td[3]")).getText().trim().equals(PacketDes)){
						driver.findElement(By.xpath("//*[@id='data']/table/tbody/tr["+j+"]/td[9]/input[3]")).click();
						break;
					}
				}
			}
			for(int n=0;n<st2.getRows();n++){
				Cell[] cell2 =st2.getRow(n);
				String name =cell2[1].getContents();
				int date =Integer.valueOf(cell2[4].getContents());
				String num =cell2[5].getContents();
				String itemName =cell2[3].getContents();
				if(name.equals("")) break;
				if(name.equals(PacketName)){
					Thread.sleep(1000);
					driver.findElement(By.id("btnItemSelect")).click();
					Thread.sleep(1000);
					driver.findElement(By.id("ItemName")).clear();
					driver.findElement(By.id("ItemName")).sendKeys(itemName);
					driver.findElement(By.id("btSearch")).click();
					Thread.sleep(3000);
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
            driver.findElement(By.cssSelector("input[value=\"提交 \"]")).click();
            Thread.sleep(1000);
            System.out.println(count + ".是否新增成功："+driver.findElement(By.className("red")).getText().equals("成功！"));
            driver.findElement(By.className("link-clew")).click();
            Thread.sleep(1000);
		}
		wb.close();

	}

}
