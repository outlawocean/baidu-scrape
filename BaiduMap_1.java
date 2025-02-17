import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.concurrent.TimeUnit;

public class BaiduMap_1 {

    public static void main(String[] args) throws InterruptedException, IOException {

        System.setProperty("webdriver.chrome.driver", "");
        ChromeOptions options=new ChromeOptions();
//        

        //无沙盒模式
        //options.addArguments("no-sandbox");
        //关闭 chrome正在受自动测试软件控制
        options.setExperimentalOption("excludeSwitches", new String[]{"enable-automation"});
        options.addArguments("--remote-allow-origins=*")
                .addArguments("--window-size=1920,1080");

        ChromeDriver driver = new ChromeDriver(options);
        WebDriverWait wait = new WebDriverWait(driver,10);

        // 读取Excel文件
        String excelFilePath = ""; // 替换为你的Excel文件路径
        FileInputStream fileInputStream = new FileInputStream(excelFilePath);
        Workbook workbook = new XSSFWorkbook(fileInputStream);
        Sheet sheet = workbook.getSheetAt(0);

        // 遍历Excel的第一列，从第二行开始，读取前10个地址
        for (int i = 101; i <= 945 && i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row == null) continue;

            Cell addressCell = row.getCell(0);
            if (addressCell == null) continue;

            String addressToSearch = addressCell.getStringCellValue();

            // 打开百度地图并进行搜索
            driver.get("https://map.baidu.com/");
            TimeUnit.SECONDS.sleep(1); // 等待页面加载
            wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector("#sole-input"))).sendKeys(addressToSearch);
//            System.out.println("已输入");
            wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector("#search-button"))).click();
//            System.out.println("已点击搜索");

            //点击第一个搜索结果 #card-1 > div > div.poi-wrapper > ul > li.search-item.base-item
//            wait.until(ExpectedConditions.elementToBeClickable(By
//                            .cssSelector("#card-1 > div > div.poi-wrapper > ul > li.search-item.base-item")))
//                    .click();
            TimeUnit.SECONDS.sleep(2); // 等待页面加载
            try {
            WebElement element = wait.until(ExpectedConditions.elementToBeClickable(By
                            .cssSelector("#card-1 > div > div.poi-wrapper > ul > li.search-item.base-item")));
            ((JavascriptExecutor) driver).executeScript("arguments[0].click();", element);
            } catch (Exception e) {
                // 如果找不到元素或超时，设置默认值
                System.out.println(i+":"+addressToSearch+"未找到");
                continue;
            }



            TimeUnit.SECONDS.sleep(1); // 等待页面加载

            // 获取搜索结果
            String companyName = "";
            String address = "";
            String type = "";
            String phoneNumber = "";
            String mapLink = driver.getCurrentUrl();

            try {
                // 公司名称 #card-1 > div > div.poi-wrapper > ul > li:nth-child(1)
                companyName = wait.until(ExpectedConditions.visibilityOfElementLocated(
                                By.cssSelector("#generalheader > div.generalHead-left-header.animation-common > div.generalHead-left-header-title > span")))
                        .getText();
            } catch (Exception e) {
                // 如果找不到元素或超时，设置默认值
                System.out.println(i+":"+addressToSearch+"未找到");
                continue;
            }

                // 地址
                address = wait.until(ExpectedConditions.visibilityOfElementLocated(
                                By.cssSelector("#generalinfo > div.generalInfo-address-telnum > div.generalInfo-address.item > span.generalInfo-address-text")))
                        .getText();

                // 地点类型
            try {
                type = wait.until(ExpectedConditions.visibilityOfElementLocated(
                            By.cssSelector("#generalheader > div.generalHead-left-header.animation-common > div.generalHead-left-header-aoitag.animation-common > span")))
                        .getText();
            } catch (Exception e) {
                // 如果找不到元素或超时，设置默认值
                type = "null";
            }


                // 手机号码
//                phoneNumber = wait.until(ExpectedConditions.visibilityOfElementLocated(
//                                By.cssSelector("#generalinfo > div.generalInfo-address-telnum > div.generalInfo-telnum.item > span.clampword.generalInfo-telnum-text")))
//                        .getText();
            try {
                // 尝试获取手机号码
                phoneNumber = wait.until(ExpectedConditions.visibilityOfElementLocated(
                                By.cssSelector("#generalinfo > div.generalInfo-address-telnum > div.generalInfo-telnum.item > span.clampword.generalInfo-telnum-text")))
                        .getText();
            } catch (Exception e) {
                // 如果找不到元素或超时，设置默认值
                phoneNumber = "null";
            }



//            } catch (Exception e) {
//                System.out.println("Failed to retrieve some information for: " + addressToSearch);
//            }

            // 将数据写入Excel的相应列
            row.createCell(4).setCellValue(companyName); // B列：公司名称
            row.createCell(5).setCellValue(address);      // C列：地址
            row.createCell(12).setCellValue(type);         // D列：地点类型
            row.createCell(13).setCellValue(phoneNumber);  // E列：手机号码
            row.createCell(14).setCellValue(mapLink);      // F列：地图链接

            System.out.println(i+":"+addressToSearch);

            // 每次获取数据后立即保存到Excel
            try (FileOutputStream fileOutputStream = new FileOutputStream(excelFilePath)) {
                workbook.write(fileOutputStream);
            } catch (IOException e) {
                System.out.println("保存数据到Excel文件失败：" + e.getMessage());
            }
        }

        // 保存Excel文件
        fileInputStream.close();
        FileOutputStream fileOutputStream = new FileOutputStream(excelFilePath);
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        workbook.close();
    }
}
