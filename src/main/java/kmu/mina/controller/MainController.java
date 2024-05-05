package kmu.mina.controller;

import jakarta.servlet.http.HttpServletResponse;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;

import java.io.IOException;
import java.io.OutputStream;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.atomic.AtomicBoolean;

@Controller
public class MainController {

    @RequestMapping("/")
    public String main() {
        return "main";
    }

    @PostMapping("/mina3")
    @ResponseBody
    public String getExcelDownThird(@RequestParam("year3") String year, HttpServletResponse res) {
        System.out.println("====> test 3 year : " + year);

        String fontPath = "C:\\A_ThinkTree-Project\\KMU-Project\\MINA\\chromedriver-win64\\chromedriver.exe";

        System.setProperty("webdriver.chrome.driver", fontPath);
        ChromeOptions options = new ChromeOptions();
        options.addArguments("--remote-allow-origins=*");

        WebDriver driver = new ChromeDriver(options);
        driver.get("https://www.nba.com/games?date="+year); //1946-11-02

        System.out.println("===> 요소 찾기");
        // 클래스 값이 NoDataMessage_base__xUA61 인 요소 찾기
//        WebElement element = driver.findElement(By.className("NoDataMessage_base__xUA61"));
//        System.out.println(element);

        System.out.println("check --> ");
        final WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));

        Workbook workbook = new XSSFWorkbook();
        // 클래스 값이 NoDataMessage_base__xUA61 인 요소가 있는지 여부 확인
        boolean isPresent = isElementPresent(driver, By.className("NoDataMessage_base__xUA61"));
        System.out.println("NoDataMessage_base__xUA61 클래스를 가진 요소가 존재하는가? " + isPresent);
        if (isPresent) {
            System.out.println("===> empty data ");
        } else {
            System.out.println("===> not empty data ");
            WebElement container = wait.until(ExpectedConditions.visibilityOfElementLocated(By.className("GamesView_gameCardsContainer__c_9fB")));
            // 해당 요소 바로 아래에 있는 div 요소 가져오기
            List<WebElement> divElements = container.findElements(By.xpath("./div"));

            // div 요소의 개수 출력
            System.out.println("div 요소의 개수: " + divElements.size());

            JavascriptExecutor js = (JavascriptExecutor) driver;
            for (int i = 0; i < divElements.size(); i++) {
//            for (int i = 0; i < 1; i++) {
                try {
                    Thread.sleep(8000);
                } catch (InterruptedException e) {
                    e.printStackTrace();
                }
                WebElement gameCardsContainer = driver.findElement(By.className("GamesView_gameCardsContainer__c_9fB"));
                WebElement secondGameCardDiv = gameCardsContainer.findElements(By.className("GameCard_gc__UCI46")).get(i);
                WebElement boxScoreLink2 = secondGameCardDiv.findElement(By.xpath(".//a[contains(text(),'BOX SCORE')]"));
                js.executeScript("arguments[0].click();", boxScoreLink2);

                // tbody content #1
//            WebElement mainDiv = driver.findElement(By.className("MaxWidthContainer_mwc__ID5AG"));
//            System.out.println("1 : ");
//            List<WebElement> tbodyList = mainDiv.findElements(By.className("StatsTableBody_tbody__uvj_P"));
//            System.out.println("===> tbody list size : "+tbodyList.size());
//            WebElement firstTbody = mainDiv.findElements(By.className("StatsTableBody_tbody__uvj_P")).get(0);
                WebElement firstStatsTableBody = driver.findElement(By.xpath("//div[@class='MaxWidthContainer_mwc__ID5AG']//tbody[@class='StatsTableBody_tbody__uvj_P'][1]"));
                List<WebElement> rows1 = firstStatsTableBody.findElements(By.tagName("tr"));
                int rowLevel = 1;
                Sheet sheet = workbook.createSheet((i+1)+" NBA GAMES");
                Row excelRow = sheet.createRow(0);
                try {
                    System.out.println("===> first tbody content ");
                    String[] header = {"PLAYER", "MIN", "FGM", "FGA", "FG%", "3PM", "3PA", "3P%", "FTM", "FTA", "FT%", "OREB", "DREB", "REB", "AST", "STL", "BLK", "TO", "PF", "PTS", "+/-"};

                    for (int j = 0; j < header.length; j++) {
                        Cell cell = excelRow.createCell(j);
                        cell.setCellValue(header[j]);
                    }

                    for (WebElement webRow : rows1) {
                        List<WebElement> columns = webRow.findElements(By.tagName("td"));
                        int colNum = 0;
                        excelRow = sheet.createRow(rowLevel);
                        for (WebElement column : columns) {
//                            System.out.print(column.getText() + "\t");
                            String data = column.getText();
                            Cell cell = excelRow.createCell(colNum++);
                            cell.setCellValue(data);
                        }
                        rowLevel++;
                    }

                    rowLevel++;
                    System.out.println("2 : ");
                    WebElement secondStatsTableBody = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("(//div[@class='MaxWidthContainer_mwc__ID5AG']//tbody[@class='StatsTableBody_tbody__uvj_P'])[2]")));
                    List<WebElement> rows2 = secondStatsTableBody.findElements(By.tagName("tr"));
                    System.out.println("===> second tbody content ");

                    excelRow = sheet.createRow(rowLevel++);
                    for(int j=0; j<header.length; j++) {
                        Cell cell = excelRow.createCell(j);
                        cell.setCellValue(header[j]);
                    }

                    for (WebElement webRow : rows2) {
                        List<WebElement> columns = webRow.findElements(By.tagName("td"));
                        int colNum = 0;
                        excelRow = sheet.createRow(rowLevel);
                        for (WebElement column : columns) {
//                            System.out.print(column.getText() + "\t");
                            String data = column.getText();
                            Cell cell = excelRow.createCell(colNum++);
                            cell.setCellValue(data);
                        }
                        rowLevel++;
                    }

                } catch (Exception e) {
                    System.out.println("ERROR : " + e.getMessage());
                    e.printStackTrace();
                }

                // 5초 대기
                try {
                    Thread.sleep(7000);
                } catch (InterruptedException e) {
                    e.printStackTrace();
                }

                System.out.println("===> SUMMARY 클릭 ");
                WebElement summaryLink = driver.findElement(By.xpath("//a[@class='InnerNavTabLink_link__Qz2Bi' and text()='Summary']"));
                js.executeScript("arguments[0].click();", summaryLink);

                try {
                    Thread.sleep(5000);
                } catch (InterruptedException e) {
                    e.printStackTrace();
                }
                System.out.println("===> search div > ");
                List<WebElement> infoRows = driver.findElements(By.xpath("//div[@class='Block_blockContent__6iJ_n']//div[@class='InfoCard_row__FO1v_']"));
                System.out.println(infoRows.size());

                String leadChanges = "";
                String timeTied = "";

                excelRow = sheet.createRow(++rowLevel);
                Cell gameInfoCell = excelRow.createCell(0);
                gameInfoCell.setCellValue("GAME INFO");
                rowLevel++;
                // GAME INFO 각 InfoCard_row__FO1v_ 요소에서 데이터를 추출
                for (WebElement infoRow : infoRows) {
                    System.out.println("===> get data : ");
                    WebElement labelElement = infoRow.findElement(By.xpath(".//div[@class='InfoCard_column__et46d']"));
                    WebElement valueElement = infoRow.findElement(By.xpath(".//div[@class='InfoCard_column__et46d']/following-sibling::*"));
                    System.out.print(infoRow.getText() + "\t");
                    String label = labelElement.getText();
                    String value = valueElement.getText();
                    int colNum = 0;
                    excelRow = sheet.createRow(rowLevel++);
                    if ("Date".equals(label) || "Location".equals(label) || "Officials".equals(label) || "Attendance".equals(label)) {
                        System.out.println("title ===> GAME INFO");
                        System.out.println(label + ": " + value);
                        Cell cell = excelRow.createCell(colNum++);
                        cell.setCellValue(label);
                        cell = excelRow.createCell(colNum);
                        cell.setCellValue(value);
                    }
                    if ("Lead Changes".equals(label)) {
                        System.out.println("title ===> LINESCORES");
                        System.out.println(label + ": " + value);
                        leadChanges = value;
                    } else if ("Times Tied".equals(label)) {
                        System.out.println("title ===> LINESCORES");
                        System.out.println(label + ": " + value);
                        timeTied = value;
                    }
                }

                // LINESCORES 표 데이터 추출
                System.out.println("leadChanges : " + leadChanges);
                System.out.println("timeTied : " + timeTied);

                // 첫 번째 테이블의 tbody 추출
//            WebElement firstTableBody = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("(//table[@class='GameLinescore_table__a1awr'])[1]/tbody")));
                try {
                    Thread.sleep(5000);
                } catch (InterruptedException e) {
                    e.printStackTrace();
                }
                // 첫 번째 GameLinescore_table__a1awr 테이블의 tbody 요소를 찾음
                WebElement firstTableBody = driver.findElement(By.xpath("(//table[@class='GameLinescore_table__a1awr'])[1]/tbody"));

                // 첫 번째 테이블의 각 행(tr)의 각 열(td) 값을 추출
                List<WebElement> firstRows = firstTableBody.findElements(By.tagName("tr"));
                System.out.println("===> firstRow ");
                System.out.println(firstRows.size());
                excelRow = sheet.createRow(rowLevel++);
                Cell linescores = excelRow.createCell(0);
                linescores.setCellValue("LINESCORES");

                WebElement summaryFirstHeader = driver.findElement(By.xpath("(//table[@class='GameLinescore_table__a1awr'])[1]/thead"));
                List<WebElement> headRow = summaryFirstHeader.findElements(By.tagName("tr"));
                for(WebElement webRow : headRow) {
                    List<WebElement> columns = webRow.findElements(By.tagName("th"));
                    int colNum = 0;
                    excelRow = sheet.createRow(++rowLevel);
                    for(WebElement column : columns) {
                        String text = (String) js.executeScript("return arguments[0].innerText;", column);
                        System.out.println("text: "+text);
                        Cell cell = excelRow.createCell(colNum++);
                        cell.setCellValue(text);
                    }
                }
                rowLevel++;
                for (WebElement row : firstRows) {
                    System.out.println("==> first table row ");
                    System.out.println(row);
                    int colNum = 0;
                    List<WebElement> cells = row.findElements(By.tagName("td"));
                    excelRow = sheet.createRow(rowLevel);
                    for (WebElement cell : cells) {
//                    System.out.print(cell.getText() + "\t");
                        String text = (String) js.executeScript("return arguments[0].innerText;", cell);
                        System.out.print(text + "\t");
                        Cell summaryFirstCell = excelRow.createCell(colNum++);
                        summaryFirstCell.setCellValue(text);
                        // td 요소의 하위 요소에 포함된 텍스트 가져오기
//                    String text = cell.findElement(By.xpath("./*")).getText();
//                    System.out.print(text + "\t");
                    }
                    System.out.println();
                    rowLevel++;
                }

                System.out.println("===> summary second table > ");

                // 두 번째 GameLinescore_table__a1awr 테이블의 tbody 요소를 찾음
                WebElement secondTableBody = driver.findElement(By.xpath("(//table[@class='GameLinescore_table__a1awr'])[2]/tbody"));
                // 두 번째 테이블의 각 행(tr)의 각 열(td) 값을 추출
                List<WebElement> secondRows = secondTableBody.findElements(By.tagName("tr"));
                System.out.println("===> firstRow ");
                System.out.println(secondRows.size());
                String[] summarySecondHeader = {"TEAM", "PITP", "FB PTS", "BIG LD", "BPTS", "TREB", "TOV", "TTOV", "POT"};
                rowLevel++;
                excelRow = sheet.createRow(rowLevel++);
                for(int j=0; j<summarySecondHeader.length; j++) {
                    Cell cell = excelRow.createCell(j);
                    cell.setCellValue(summarySecondHeader[j]);
                }
                for (WebElement row : secondRows) {
                    List<WebElement> cells = row.findElements(By.tagName("td"));
                    int colNum = 0;
                    excelRow = sheet.createRow(rowLevel);
                    for (WebElement cell : cells) {
                        String text = (String) js.executeScript("return arguments[0].innerText;", cell);
                        System.out.print(text + "\t");
                        Cell summarySecondCell = excelRow.createCell(colNum++);
                        summarySecondCell.setCellValue(text);
//                    System.out.print(cell.getText() + "\t");
                    }
                    System.out.println();
                    rowLevel++;
                }
                rowLevel++;
                excelRow = sheet.createRow(rowLevel++);
                Cell leadCD = excelRow.createCell(0);
                leadCD.setCellValue("Lead Changes");
                leadCD = excelRow.createCell(1);
                leadCD.setCellValue(leadChanges);
                excelRow = sheet.createRow(rowLevel);
                Cell times = excelRow.createCell(0);
                times.setCellValue("Times Tied");
                times = excelRow.createCell(1);
                times.setCellValue(timeTied);

                System.out.println(i+ "===> 번 째 완료");

                // 5초 대기
                try {
                    Thread.sleep(5000);
                } catch (InterruptedException e) {
                    e.printStackTrace();
                }

                System.out.println("===> back 이동");
                driver.get("https://www.nba.com/games?date="+year);

            }
        }

        res.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        res.setHeader("Content-disposition", "attachment; filename=" + year + "_GAME.xlsx");
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

    // 클래스 값을 가진 요소의 존재 여부를 확인하는 메서드
    public static boolean isElementPresent(WebDriver driver, By by) {
        try {
            driver.findElement(by);
            return true;
        } catch (org.openqa.selenium.NoSuchElementException e) {
            return false;
        }
    }

    @PostMapping("/mina2")
    @ResponseBody
    public String getExcelDownLocation(@RequestParam("year2") String year, HttpServletResponse res) {
        System.out.println("==> test 2");

        String fontPath = "C:\\A_ThinkTree-Project\\KMU-Project\\MINA\\chromedriver-win64\\chromedriver.exe";

        System.setProperty("webdriver.chrome.driver", fontPath);

        ChromeOptions options = new ChromeOptions();
        options.addArguments("--remote-allow-origins=*");
//        options.addArguments("--disable-gpu");			//gpu 비활성화
//        options.addArguments("--blink-settings=imagesEnabled=false"); //이미지 다운 안받음

        WebDriver driver = new ChromeDriver(options);
        driver.get("https://www.nba.com/stats/players/boxscores-traditional?Season=" + year);

//        WebElement allOption = driver.findElement(By.xpath("//option[text()='All']"));
//        WebElement allOption = driver.findElement(By.xpath("//option[text()='301']"));
//        allOption.click();

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("NBA DATA");

        final String[] header = {"PLAYER", "TEAM", "MATCH UP", "GAME DATE", "W/L", "MIN", "PTS", "FGM", "FGA", "FG%", "3PM", "3PA", "3P%", "FTM", "FTA", "FT%", "OREB", "DREB", "REB", "AST", "STL", "BLK", "TOV", "PF", "+/-"};
        Row row = sheet.createRow(0);
        for (int i = 0; i < header.length; i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue(header[i]);
        }

        WebElement tbody = driver.findElement(By.className("Crom_body__UYOcU"));
        List<WebElement> rows = tbody.findElements(By.tagName("tr"));

        try {
            List<List<String>> dataList = new ArrayList<>();
            int rowLevel = 1;
            for (WebElement webRow : rows) {
                List<WebElement> columns = webRow.findElements(By.tagName("td"));
//                int colNum = 0;
//                row = sheet.createRow(rowLevel);
                List<String> rowData = new ArrayList<>();
                for (WebElement column : columns) {
                    rowData.add(column.getText());
//                System.out.print(column.getText() + "\t");
//                    String data = column.getText();
//                    Cell cell = row.createCell(colNum++);
//                    cell.setCellValue(data);
                }
                dataList.add(rowData);
//                System.out.println(rowLevel+ "번 작성중...");
//                rowLevel++;
            }
//            WebElement allOption = driver.findElement(By.xpath("//option[text()='All']"));
//            allOption.click();

            WebElement nextButton = driver.findElement(By.cssSelector("button[data-pos='next']"));
//            System.out.println(nextButton);
//            System.out.println(nextButton.getAttribute("disabled"));
            nextButton.click();
            System.out.println(rowLevel + "50 >");
            boolean nextButtonBool = true;
            WebElement tbody2 = driver.findElement(By.className("Crom_body__UYOcU"));
            while (nextButtonBool) {
                List<WebElement> rows2 = tbody2.findElements(By.tagName("tr"));

                for (WebElement webRow : rows2) {
                    List<WebElement> columns = webRow.findElements(By.tagName("td"));
                    List<String> rowData = new ArrayList<>();
//                    int colNum = 0;
//                    row = sheet.createRow(rowLevel);
                    for (WebElement column : columns) {
                        rowData.add(column.getText());
//                        String data = column.getText();
//                        Cell cell = row.createCell(colNum++);
//                        cell.setCellValue(data);
                    }
                    dataList.add(rowData);
//                    System.out.println(rowLevel+ "번 작성중...");
//                    rowLevel++;
//                    if(rowLevel == 5000) {
//                        nextButtonBool = false;
//                    }
                }
                WebElement nextButton2 = driver.findElement(By.cssSelector("button[data-pos='next']"));
//                System.out.println(nextButton);
//                System.out.println(nextButton2.getAttribute("disabled"));
                if ("true".equals(nextButton2.getAttribute("disabled"))) {
                    nextButtonBool = false;
                } else {
                    nextButton2.click();
                }
//                WebElement allOption = driver.findElement(By.xpath("//option[text()='All']"));
            }
            System.out.println("excel insert : ");
            int rowNum = 0;
            for (List<String> rowData : dataList) {
                Row row2 = sheet.createRow(rowNum++);
                int colNum = 0;
                for (String cellData : rowData) {
                    Cell cell = row2.createCell(colNum++);
                    cell.setCellValue(cellData);
                }
            }
            System.out.println("Excel Down 시작 >>>>>>>>>>>>> ");
        } catch (Exception e) {
            System.err.println("ERROR ! : " + e.getMessage());
            e.printStackTrace();
        }

        res.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        res.setHeader("Content-disposition", "attachment; filename=" + year + ".xlsx");
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


//    @PostMapping("/mina2")
//    @ResponseBody
//    public String getExcelDownLocation(@RequestParam("year2") String year, HttpServletResponse res) {
//        System.out.println("==> test 2");
//
//        String fontPath = "C:\\A_ThinkTree-Project\\KMU-Project\\MINA\\chromedriver-win64\\chromedriver.exe";
//
//        System.setProperty("webdriver.chrome.driver", fontPath);
//
//        ChromeOptions options = new ChromeOptions();
//        options.addArguments("--remote-allow-origins=*");
////        options.addArguments("--disable-gpu");			//gpu 비활성화
////        options.addArguments("--blink-settings=imagesEnabled=false"); //이미지 다운 안받음
//
//        WebDriver driver = new ChromeDriver(options);
//        driver.get("https://www.nba.com/stats/players/boxscores-traditional?Season="+year);
//
////        WebElement allOption = driver.findElement(By.xpath("//option[text()='All']"));
////        WebElement allOption = driver.findElement(By.xpath("//option[text()='301']"));
////        allOption.click();
//
//        Workbook workbook = new XSSFWorkbook();
//        Sheet sheet = workbook.createSheet("NBA DATA");
//
//        final String[] header = {"PLAYER", "TEAM","MATCH UP", "GAME DATE", "W/L", "MIN", "PTS", "FGM", "FGA", "FG%", "3PM", "3PA", "3P%", "FTM", "FTA", "FT%", "OREB", "DREB", "REB", "AST", "STL", "BLK", "TOV","PF","+/-"};
//        Row row = sheet.createRow(0);
//        for(int i=0; i<header.length; i++) {
//            Cell cell = row.createCell(i);
//            cell.setCellValue(header[i]);
//        }
//
////        WebElement tbody = driver.findElement(By.className("Crom_body__UYOcU"));
////        List<WebElement> rows = tbody.findElements(By.tagName("tr"));
//
//        try {
//            ExecutorService executor = Executors.newFixedThreadPool(10); // 적절한 스레드 풀 크기 선택
//
//            AtomicBoolean hasNextPage = new AtomicBoolean(true);
//
//            while (hasNextPage.get()) {
//                List<WebElement> rows = driver.findElements(By.className("Crom_body__UYOcU")).get(0).findElements(By.tagName("tr"));
//
//                List<CompletableFuture<Void>> futures = new ArrayList<>();
//
//                // 각 페이지의 첫 번째 10개의 행을 가져와 병렬로 처리
//                for (int i = 0; i < 10 && i < rows.size(); i++) {
//                    WebElement webRow = rows.get(i);
//                    CompletableFuture<Void> future = CompletableFuture.runAsync(() -> {
//                        processRow(webRow); // 행 처리 메소드 호출
//                    }, executor);
//                    futures.add(future);
//                }
//
//                // 모든 작업이 완료될 때까지 대기
//                CompletableFuture<Void> allRowsProcessed = CompletableFuture.allOf(futures.toArray(new CompletableFuture[0]));
//
//                // 모든 작업이 완료되면 다음 페이지로 이동
//                allRowsProcessed.thenRun(() -> {
//                    hasNextPage.set(navigateToNextPage(driver)); // 다음 페이지로 이동하는 메소드 호출
//                }).join();
//            }
//            executor.shutdown(); // 작업이 완료되면 ExecutorService를 종료
//
//            System.out.println("Excel Down 시작 >>>>>>>>>>>>> ");
//
//        } catch (Exception e) {
//            System.err.println("ERROR ! : "+e.getMessage());
//            e.printStackTrace();
//        }
//
//        res.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
//        res.setHeader("Content-disposition", "attachment; filename="+year+".xlsx");
//        try {
//            OutputStream out = res.getOutputStream();
//            workbook.write(out);
//            out.flush();
//            out.close();
//            workbook.close();
//        } catch (IOException e) {
//            e.printStackTrace();
//        }
//        return "main";
//    }


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
        driver.get("https://www.nba.com/stats/leaders?Season=" + year + "&PerMode=PerGame&StatCategory=MIN");
//        driver.get("https://nba.com/stats/players/boxscores-traditional?Season=2013-14");
//        driver.get("https://www.nba.com/stats/leaders?Season=1983-84&PerMode=PerGame&StatCategory=MIN");

//        WebElement dropdown = driver.findElement(By.className("DropDown_dropdown__TMlAR"));
//        dropdown.click();

        WebElement allOption = driver.findElement(By.xpath("//option[text()='All']"));
        allOption.click();

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("NBA DATA");

        final String[] header = {"#", "PLAYER", "TEAM", "GP", "MIN", "PTS", "FGM", "FGA", "FG%", "3PM", "3PA", "3P%", "FTM", "FTA", "FT%", "OREB", "DREB", "REB", "AST", "STL", "BLK", "TOV", "FF"};
        Row row = sheet.createRow(0);
        for (int i = 0; i < header.length; i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue(header[i]);
        }

        WebElement tbody = driver.findElement(By.className("Crom_body__UYOcU"));
        List<WebElement> rows = tbody.findElements(By.tagName("tr"));

        try {
            int rowLevel = 1;
            for (WebElement webRow : rows) {
                List<WebElement> columns = webRow.findElements(By.tagName("td"));
                int colNum = 0;
                row = sheet.createRow(rowLevel);
                for (WebElement column : columns) {
//                System.out.print(column.getText() + "\t");
                    String data = column.getText();
                    Cell cell = row.createCell(colNum++);
                    cell.setCellValue(data);
                }
                System.out.println(rowLevel + "번 작성중...");
                rowLevel++;
            }
            System.out.println("Excel Down 시작 >>>>>>>>>>>>> ");
        } catch (Exception e) {
            System.err.println("ERROR ! : " + e.getMessage());
            e.printStackTrace();
        }

        res.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        res.setHeader("Content-disposition", "attachment; filename=" + year + ".xlsx");
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

        for (WebElement row : rows) {
            List<WebElement> columns = row.findElements(By.tagName("td"));
            for (WebElement column : columns) {
                System.out.print(column.getText() + "\t");
            }
            System.out.println();
        }
    }

}


