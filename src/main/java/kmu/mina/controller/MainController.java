package kmu.mina.controller;

import jakarta.servlet.http.HttpServletResponse;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;

import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

@Controller
public class MainController {

    @RequestMapping("/")
    public String main() {
        return "main";
    }

    @PostMapping("/test")
    public void getTestExcel(@RequestParam("year") String year) {
        System.out.println("==> test ");
        System.out.println(year);
    }

    @PostMapping("/mina2")
    @ResponseBody
    public String getExcelDownLocation(@RequestParam("year2") String year, HttpServletResponse res) {
        System.out.println("==> test 2");

        String fontPath = "C:\\A_ThinkTree-Project\\KMU-Project\\MINA\\chromedriver-win64\\chromedriver.exe";

        System.setProperty("webdriver.chrome.driver", fontPath);

        ChromeOptions options = new ChromeOptions();
        options.addArguments("--remote-allow-origins=*");

        WebDriver driver = new ChromeDriver(options);
        driver.get("https://www.nba.com/stats/players/boxscores-traditional?Season="+year);

//        WebElement allOption = driver.findElement(By.xpath("//option[text()='All']"));
//        WebElement allOption = driver.findElement(By.xpath("//option[text()='401']"));
//        allOption.click();

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("NBA DATA");

        final String[] header = {"PLAYER", "TEAM","MATCH UP", "GAME DATE", "W/L", "MIN", "PTS", "FGM", "FGA", "FG%", "3PM", "3PA", "3P%", "FTM", "FTA", "FT%", "OREB", "DREB", "REB", "AST", "STL", "BLK", "TOV","PF","+/-"};
        Row row = sheet.createRow(0);
        for(int i=0; i<header.length; i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue(header[i]);
        }

        WebElement tbody = driver.findElement(By.className("Crom_body__UYOcU"));
        List<WebElement> rows = tbody.findElements(By.tagName("tr"));

        try {
            int rowLevel = 1;
            for(WebElement webRow : rows) {
                List<WebElement> columns = webRow.findElements(By.tagName("td"));
                int colNum = 0;
                row = sheet.createRow(rowLevel);
                for (WebElement column : columns) {
//                System.out.print(column.getText() + "\t");
                    String data = column.getText();
                    Cell cell = row.createCell(colNum++);
                    cell.setCellValue(data);
                }
                System.out.println(rowLevel+ "번 작성중...");
                rowLevel++;
            }
//            WebElement allOption = driver.findElement(By.xpath("//option[text()='All']"));
//            allOption.click();

            WebElement nextButton = driver.findElement(By.cssSelector("button[data-pos='next']"));
            System.out.println(nextButton);
            System.out.println(nextButton.getAttribute("disabled"));
            nextButton.click();

            boolean nextButtonBool = true;
            while (nextButtonBool) {
                WebElement tbody2 = driver.findElement(By.className("Crom_body__UYOcU"));
                List<WebElement> rows2 = tbody2.findElements(By.tagName("tr"));

                for(WebElement webRow : rows2) {
                    List<WebElement> columns = webRow.findElements(By.tagName("td"));
                    int colNum = 0;
                    row = sheet.createRow(rowLevel);
                    for (WebElement column : columns) {
//                System.out.print(column.getText() + "\t");
                        String data = column.getText();
                        Cell cell = row.createCell(colNum++);
                        cell.setCellValue(data);
                    }
                    System.out.println(rowLevel+ "번 작성중...");
                    rowLevel++;
                    if(rowLevel == 5000) {
                        nextButtonBool = false;
                    }
                }
                WebElement nextButton2 = driver.findElement(By.cssSelector("button[data-pos='next']"));
//                System.out.println(nextButton);
//                System.out.println(nextButton2.getAttribute("disabled"));
                if("true".equals(nextButton2.getAttribute("disabled"))) {
                    nextButtonBool = false;
                } else {
                    nextButton2.click();
                }
//                WebElement allOption = driver.findElement(By.xpath("//option[text()='All']"));

            }
            System.out.println("Excel Down 시작 >>>>>>>>>>>>> ");
        } catch (Exception e) {
            System.err.println("ERROR ! : "+e.getMessage());
            e.printStackTrace();
        }

        res.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        res.setHeader("Content-disposition", "attachment; filename="+year+".xlsx");
        try {
            OutputStream out = res.getOutputStream();
            workbook.write(out);
            out.flush();
            out.close();
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return "main";
    }

    public Row rowReturnCode(Row row, List<WebElement> rows, Sheet sheet, int rowLevel) {
        for(WebElement webRow : rows) {
            List<WebElement> columns = webRow.findElements(By.tagName("td"));
            int colNum = 0;
            row = sheet.createRow(rowLevel);
            for (WebElement column : columns) {
//                System.out.print(column.getText() + "\t");
                String data = column.getText();
                Cell cell = row.createCell(colNum++);
                cell.setCellValue(data);
            }
            System.out.println(rowLevel+ "번 작성중...");
            rowLevel++;
        }
        return row;
    }

    @PostMapping("/mina")
    @ResponseBody
    public String getNBAExcelData(@RequestParam("year") String year, HttpServletResponse res) {
//        System.out.println("=====> main test : ");
//        String year = "1951-52";

        String fontPath = "C:\\A_ThinkTree-Project\\KMU-Project\\MINA\\chromedriver-win64\\chromedriver.exe";

        System.setProperty("webdriver.chrome.driver", fontPath);

        ChromeOptions options = new ChromeOptions();
        options.addArguments("--remote-allow-origins=*");

        WebDriver driver = new ChromeDriver(options);
        driver.get("https://www.nba.com/stats/leaders?Season="+year+"&PerMode=PerGame&StatCategory=MIN");
//        driver.get("https://nba.com/stats/players/boxscores-traditional?Season=2013-14");
//        driver.get("https://www.nba.com/stats/leaders?Season=1983-84&PerMode=PerGame&StatCategory=MIN");

//        WebElement dropdown = driver.findElement(By.className("DropDown_dropdown__TMlAR"));
//        dropdown.click();

        WebElement allOption = driver.findElement(By.xpath("//option[text()='All']"));
        allOption.click();

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("NBA DATA");

        final String[] header = {"#", "PLAYER","TEAM", "GP", "MIN", "PTS", "FGM", "FGA", "FG%", "3PM", "3PA", "3P%", "FTM", "FTA", "FT%", "OREB", "DREB", "REB", "AST", "STL", "BLK", "TOV", "FF"};
        Row row = sheet.createRow(0);
        for(int i=0; i<header.length; i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue(header[i]);
        }

        WebElement tbody = driver.findElement(By.className("Crom_body__UYOcU"));
        List<WebElement> rows = tbody.findElements(By.tagName("tr"));

        try {
            int rowLevel = 1;
            for(WebElement webRow : rows) {
                List<WebElement> columns = webRow.findElements(By.tagName("td"));
                int colNum = 0;
                row = sheet.createRow(rowLevel);
                for (WebElement column : columns) {
//                System.out.print(column.getText() + "\t");
                    String data = column.getText();
                    Cell cell = row.createCell(colNum++);
                    cell.setCellValue(data);
                }
                System.out.println(rowLevel+ "번 작성중...");
                rowLevel++;
            }
            System.out.println("Excel Down 시작 >>>>>>>>>>>>> ");
        } catch (Exception e) {
            System.err.println("ERROR ! : "+e.getMessage());
            e.printStackTrace();
        }

        res.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        res.setHeader("Content-disposition", "attachment; filename="+year+".xlsx");
//        res.setHeader("Content-disposition", "attachment; filename=test.xlsx");
        try {
            OutputStream out = res.getOutputStream();
            workbook.write(out);
            out.flush();
            out.close();
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return "main";
    }

    public static void main(String[] args) {
        System.out.println("=====> main test : ");
        String fontPath = "C:\\A_ThinkTree-Project\\KMU-Project\\MINA\\chromedriver-win64\\chromedriver.exe";

        System.setProperty("webdriver.chrome.driver", fontPath); // 윈도우

        ChromeOptions options = new ChromeOptions();
        options.addArguments("--remote-allow-origins=*");

        WebDriver driver = new ChromeDriver(options);
        driver.get("https://www.nba.com/stats/leaders?Season=1951-52&PerMode=PerGame&StatCategory=MIN");
//        driver.get("https://www.nba.com/stats/leaders?Season=1983-84&PerMode=PerGame&StatCategory=MIN");

//        WebElement dropdown = driver.findElement(By.className("DropDown_dropdown__TMlAR"));
//        dropdown.click();

        WebElement allOption = driver.findElement(By.xpath("//option[text()='All']"));
        allOption.click();


        WebElement tbody = driver.findElement(By.className("Crom_body__UYOcU"));
        List<WebElement> rows = tbody.findElements(By.tagName("tr"));

        for(WebElement row : rows) {
            List<WebElement> columns = row.findElements(By.tagName("td"));
            for (WebElement column : columns) {
                System.out.print(column.getText() + "\t");
            }
            System.out.println();
        }
    }

}


