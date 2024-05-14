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
import java.util.Arrays;
import java.util.List;

@Controller
public class MainController {

    @RequestMapping("/")
    public String main() {
        return "main";
    }


    @PostMapping("/gameData")
    @ResponseBody
    public String getDownloadGameData(@RequestParam("year4") String year, HttpServletResponse res) {
        System.out.println("year : " + year);
        String[] day = year.split("-");
        String fontPath = "C:\\A_ThinkTree-Project\\KMU-Project\\MINA\\chromedriver-win64\\chromedriver.exe";

        System.setProperty("webdriver.chrome.driver", fontPath);
        ChromeOptions options = new ChromeOptions();
        options.addArguments("--remote-allow-origins=*");
        options.addArguments("--disable-popup-blocking");       // 팝업안띄움
        options.addArguments("--disable-gpu");                  // gpu 비활성화
        options.addArguments("--disable-images");
        options.addArguments("headless");                       // 브라우저 안띄움
        options.addArguments("--no-sandbox");
        options.addArguments("--blink-settings=imagesEnabled=false"); // 이미지 다운 안받음

        WebDriver driver = new ChromeDriver(options);
        driver.get("https://www.nba.com/games?date=" + year); // ex)1946-11-02

        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
        Workbook workbook = new XSSFWorkbook();

        /** 클래스 값이 NoDataMessage_base__xUA61 인 요소가 있는지 여부 확인 */
        boolean isPresent = isElementPresent(driver, By.className("NoDataMessage_base__xUA61"));
        System.out.println("NoDataMessage_base__xUA61 클래스를 가진 요소가 존재 유무 " + isPresent);
        if (!isPresent) {
            WebElement container = wait.until(ExpectedConditions.visibilityOfElementLocated(By.className("GamesView_gameCardsContainer__c_9fB")));
            /** 해당 요소 바로 아래에 있는 div 요소 가져오기 */
            List<WebElement> divElements = container.findElements(By.xpath("./div"));

            /** div 요소의 개수 출력 */
            System.out.println("GAME 개수: " + divElements.size());

            /** excel header :: 쿼터수는 변동이 있을 수 있음 */
            String[] header = {"Year", "Month", "Day", "Attendance", "TEAM", "HOME (홈)", "AWAY (어웨이)", "L(0)/W(1)", "PLAYER", "MIN", "Q1", "Q2", "Q3", "Q4", "OT1", "FINAL", "TOT_FGM", "TOT_FGA", "TOT_FG%", "TOT_3PM", "TOT_3PA", "TOT_3P%", "TOT_FTM", "TOT_FTA", "TOT_FT%", "TOT_OREB", "TOT_DREB", "TOT_REB", "TOT_AST", "TOT_STL", "TOT_BLK", "TOT_TO", "TOT_PF", "TOT_PTS", "TOT_+/-", "FGM", "FGA", "FG%", "3PM", "3PA", "3P%", "FTM", "FTA", "FT%", "OREB", "DREB", "REB", "AST", "STL", "BLK", "TO", "PF", "PTS", "+/-"};
            JavascriptExecutor js = (JavascriptExecutor) driver;

            Sheet sheet = workbook.createSheet("NBA GAMES");
            Row excelRow = sheet.createRow(0);
            for (int i = 0; i < header.length; i++) {
                Cell cell = excelRow.createCell(i);
                cell.setCellValue(header[i]);
            }
            List<String[]> excelList = new ArrayList<>();
            System.out.println("<<<< GAME SEARCH >>>>");

            for (int i = 0; i < divElements.size(); i++) {
//            for (int i = 0; i < 1; i++) {
                List<String[]> gameDataList = new ArrayList<>();
                try {
                    Thread.sleep(4000);
                } catch (InterruptedException e) {
                    e.printStackTrace();
                }

                /** BOX SCORE CLICK */
//                WebElement gameCardsContainer = driver.findElement(By.className("GamesView_gameCardsContainer__c_9fB"));
                WebElement gameCardsContainer = wait.until(ExpectedConditions.visibilityOfElementLocated(By.className("GamesView_gameCardsContainer__c_9fB")));
                WebElement secondGameCardDiv = gameCardsContainer.findElements(By.className("GameCard_gc__UCI46")).get(i);
                WebElement boxScoreLink2 = secondGameCardDiv.findElement(By.xpath(".//a[contains(text(),'BOX SCORE')]"));
                js.executeScript("arguments[0].click();", boxScoreLink2);

                WebElement firstStatsTableBody = driver.findElement(By.xpath("//div[@class='MaxWidthContainer_mwc__ID5AG']//tbody[@class='StatsTableBody_tbody__uvj_P'][1]"));
                List<WebElement> rows1 = firstStatsTableBody.findElements(By.tagName("tr"));
//                List<WebElement> rows1 = driver.findElements(By.xpath("//tr"));

                boolean totalFound = false;
                int firstRowNumber = 0;

                ArrayList<String> firstTotalData = new ArrayList<>();

                System.out.println("첫번째 TABLE 크롤링 시작 ");
                for (WebElement webRow : rows1) {
                    String[] excelData = new String[header.length];
                    Arrays.fill(excelData, "-");
                    List<WebElement> columns = webRow.findElements(By.tagName("td"));
                    int totalNum = 0;
                    int num = 8;
                    for (WebElement column : columns) {
                        String text = column.getText().split("\n")[0];
                        if (num == 10) {
                            num = 35;
                        }
                        if ("TOTALS".equals(text)) {
                            totalFound = true;
                        }
                        if (totalFound) {
                            firstTotalData.add(totalNum, text);
                            totalNum++;
                        } else {
                            excelData[num] = text;
                        }
                        num++;
                    }
                    if (!totalFound) {
                        firstRowNumber++;
                        gameDataList.add(excelData);
                    }
                }

                System.out.println("1 table : ");
                for (String[] array : gameDataList) {
                    for (String value : array) {
                        System.out.print(value + " ");
                    }
                    System.out.println();
                }
                System.out.println("total data : " + firstTotalData);
                System.out.println("firstRowNumber : " + firstRowNumber);

                WebElement secondStatsTableBody = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("(//div[@class='MaxWidthContainer_mwc__ID5AG']//tbody[@class='StatsTableBody_tbody__uvj_P'])[2]")));
                List<WebElement> rows2 = secondStatsTableBody.findElements(By.tagName("tr"));
                ArrayList<String> secondTotalData = new ArrayList<>();
                totalFound = false;
                int secondRowNumber = 0;

                System.out.println("second table start > ");
                for (WebElement webRow : rows2) {
                    String[] excelData = new String[header.length];
                    Arrays.fill(excelData, "-");
                    List<WebElement> columns = webRow.findElements(By.tagName("td"));

                    int totalNum = 0;
                    int num = 8;
                    for (WebElement column : columns) {
                        String text = column.getText().split("\n")[0];
                        if (num == 10) {
                            num = 35;
                        }
                        if ("TOTALS".equals(text)) {
                            totalFound = true;
                        }
                        if (totalFound) {
                            secondTotalData.add(totalNum, text);
                            totalNum++;
                        } else {
                            excelData[num] = text;
                        }
                        num++;
                    }
                    if (!totalFound) {
                        secondRowNumber++;
                        gameDataList.add(excelData);
                    }
                }

                System.out.println("2 table :");
                for (String[] array : gameDataList) {
                    for (String value : array) {
                        System.out.print(value + " ");
                    }
                    System.out.println();
                }
                System.out.println("second total data : " + secondTotalData);
                System.out.println("==> second row number : " + secondRowNumber);

                System.out.println("===> SUMMARY click ");
                WebElement summaryLink = driver.findElement(By.xpath("//a[@class='InnerNavTabLink_link__Qz2Bi' and text()='Summary']"));
                js.executeScript("arguments[0].click();", summaryLink);

                try {
                    Thread.sleep(3000);
                } catch (InterruptedException e) {
                    e.printStackTrace();
                }
                List<WebElement> infoRows = driver.findElements(By.xpath("//div[@class='Block_blockContent__6iJ_n']//div[@class='InfoCard_row__FO1v_']"));
//                List<WebElement> infoRows = wait.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.xpath("//div[@class='Block_blockContent__6iJ_n']//div[@class='InfoCard_row__FO1v_']")));

                String attendance = "-";        // 4번째에 들어감
                /** GAME INFO 각 InfoCard_row__FO1v_ 요소에서 데이터를 추출 */
                System.out.println("===> GAME INFO data : ");
                for (WebElement infoRow : infoRows) {
                    WebElement labelElement = infoRow.findElement(By.xpath(".//div[@class='InfoCard_column__et46d']"));
                    WebElement valueElement = infoRow.findElement(By.xpath(".//div[@class='InfoCard_column__et46d']/following-sibling::*"));
                    System.out.print(infoRow.getText() + "\t");
                    String label = labelElement.getText();
                    String value = valueElement.getText();
//                    System.out.print("label: " + label + " value: " + value);
                    if ("Attendance".equals(label)) {
                        attendance = value;
                    }
                }
                System.out.println("attendance : " + attendance);     // (-) 01-02 값 확인 필요
//                try {
//                    Thread.sleep(5000);
//                } catch (InterruptedException e) {
//                    e.printStackTrace();
//                }

                /** 첫 번째 GameLinescore_table__a1awr 테이블의 tbody 요소를 찾음 */
                WebElement firstTableBody = driver.findElement(By.xpath("(//table[@class='GameLinescore_table__a1awr'])[1]/tbody"));
//                WebElement firstTableBody = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("(//table[@class='GameLinescore_table__a1awr'])[1]/tbody")));
                /** 첫 번째 테이블의 각 행(tr)의 각 열(td) 값을 추출 */
                List<WebElement> firstRows = firstTableBody.findElements(By.tagName("tr"));
//                WebElement summaryFirstHeader = driver.findElement(By.xpath("(//table[@class='GameLinescore_table__a1awr'])[1]/thead"));
//                List<WebElement> headRow = summaryFirstHeader.findElements(By.tagName("tr"));
                ArrayList<String> firstTable = new ArrayList<>();
                ArrayList<String> secondTable = new ArrayList<>();

                int tableRow = 0;
                for (WebElement row : firstRows) {
                    System.out.println("==> first table row ");
                    List<WebElement> cells = row.findElements(By.tagName("td"));
                    int tableNum = 0;
                    for (WebElement cell : cells) {
                        String text = (String) js.executeScript("return arguments[0].innerText;", cell);
                        System.out.print(text + "\t");
                        if (tableRow == 0) {
                            firstTable.add(tableNum, text);
                        } else {
                            secondTable.add(tableNum, text);
                        }
                        tableNum++;
                    }
                    tableRow++;
                    System.out.println();
                }

                System.out.println("firstTable : " + firstTable);
                System.out.println("secondTable : " + secondTable);

                for (String[] array : gameDataList) {
                    array[0] = day[0];
                    array[1] = day[1];
                    array[2] = day[2];
                    array[3] = attendance;
                    array[5] = secondTable.get(0);  // home team
                    array[6] = firstTable.get(0);   // away team
                }

                /** array 9번부터 진행 */
                int awayTeamScore = Integer.parseInt(firstTable.get(firstTable.size() - 1));
                int homeTeamScore = Integer.parseInt(secondTable.get(secondTable.size() - 1));
                if (awayTeamScore > homeTeamScore) {
                    awayTeamScore = 1;
                    homeTeamScore = 0;
                } else {
                    awayTeamScore = 0;
                    homeTeamScore = 1;
                }
                System.out.println("awayTeamScore: " + awayTeamScore);
                System.out.println("homeTeamScore: " + homeTeamScore);

                for (int j = 0; j < firstRowNumber; j++) {
                    gameDataList.get(j)[4] = firstTable.get(0);
                    gameDataList.get(j)[7] = String.valueOf(awayTeamScore);
                    if(isFirstCharacterDigit(gameDataList.get(j)[9])) {
//                    if (!gameDataList.get(j)[9].startsWith("D")) {
                        switch (firstTable.size()) {
                            case 7:
                                gameDataList.get(j)[14] = firstTable.get(5); // OT1
                            case 6:
                                gameDataList.get(j)[13] = firstTable.get(4); // Q4
                            case 5:
                                gameDataList.get(j)[12] = firstTable.get(3); // Q3
                            case 4:
                                gameDataList.get(j)[11] = firstTable.get(2); // Q2
                            case 3:
                                gameDataList.get(j)[10] = firstTable.get(1); // Q1
                                break;
                        }
                        gameDataList.get(j)[15] = firstTable.get(firstTable.size() - 1);    // FINAL
                        gameDataList.get(j)[16] = firstTotalData.get(2);        // TOTO_FGM
                        gameDataList.get(j)[17] = firstTotalData.get(3);        // TOT_EGA
                        gameDataList.get(j)[18] = firstTotalData.get(4);        // TOT_FG%
                        gameDataList.get(j)[19] = firstTotalData.get(5);        //TOT_3PM
                        gameDataList.get(j)[20] = firstTotalData.get(6);        //
                        gameDataList.get(j)[21] = firstTotalData.get(7);
                        gameDataList.get(j)[22] = firstTotalData.get(8);
                        gameDataList.get(j)[23] = firstTotalData.get(9);
                        gameDataList.get(j)[24] = firstTotalData.get(10);
                        gameDataList.get(j)[25] = firstTotalData.get(11);
                        gameDataList.get(j)[26] = firstTotalData.get(12);
                        gameDataList.get(j)[27] = firstTotalData.get(13);
                        gameDataList.get(j)[28] = firstTotalData.get(14);
                        gameDataList.get(j)[29] = firstTotalData.get(15);
                        gameDataList.get(j)[30] = firstTotalData.get(16);
                        gameDataList.get(j)[31] = firstTotalData.get(17);
                        gameDataList.get(j)[32] = firstTotalData.get(18);
                        gameDataList.get(j)[33] = firstTotalData.get(19);
                        gameDataList.get(j)[34] = firstTotalData.get(20);
                    } else {
                        gameDataList.get(j)[9] = "-";
                    }
                }

                for (int j = firstRowNumber; j < gameDataList.size(); j++) {
                    gameDataList.get(j)[4] = secondTable.get(0);
                    gameDataList.get(j)[7] = String.valueOf(homeTeamScore);
                    if(isFirstCharacterDigit(gameDataList.get(j)[9])) {
//                    if (!gameDataList.get(j)[9].startsWith("D")) {
                        switch (secondTable.size()) {
                            case 7:
                                gameDataList.get(j)[14] = secondTable.get(5); // OT1
                            case 6:
                                gameDataList.get(j)[13] = secondTable.get(4); // Q4
                            case 5:
                                gameDataList.get(j)[12] = secondTable.get(3); // Q3
                            case 4:
                                gameDataList.get(j)[11] = secondTable.get(2); // Q2
                            case 3:
                                gameDataList.get(j)[10] = secondTable.get(1); // Q1
                                break;
                        }
                        gameDataList.get(j)[15] = secondTable.get(secondTable.size() - 1);      // TOTAL

                        gameDataList.get(j)[16] = secondTotalData.get(2);
                        gameDataList.get(j)[17] = secondTotalData.get(3);
                        gameDataList.get(j)[18] = secondTotalData.get(4);
                        gameDataList.get(j)[19] = secondTotalData.get(5);
                        gameDataList.get(j)[20] = secondTotalData.get(6);
                        gameDataList.get(j)[21] = secondTotalData.get(7);
                        gameDataList.get(j)[22] = secondTotalData.get(8);
                        gameDataList.get(j)[23] = secondTotalData.get(9);
                        gameDataList.get(j)[24] = secondTotalData.get(10);
                        gameDataList.get(j)[25] = secondTotalData.get(11);
                        gameDataList.get(j)[26] = secondTotalData.get(12);
                        gameDataList.get(j)[27] = secondTotalData.get(13);
                        gameDataList.get(j)[28] = secondTotalData.get(14);
                        gameDataList.get(j)[29] = secondTotalData.get(15);
                        gameDataList.get(j)[30] = secondTotalData.get(16);
                        gameDataList.get(j)[31] = secondTotalData.get(17);
                        gameDataList.get(j)[32] = secondTotalData.get(18);
                        gameDataList.get(j)[33] = secondTotalData.get(19);
                        gameDataList.get(j)[34] = secondTotalData.get(20);
                    } else {
                        gameDataList.get(j)[9] = "-";
                    }
                }

                for (String[] array : gameDataList) {
                    for (String value : array) {
                        System.out.print(value + " ");
                    }
                    System.out.println();
                }
                driver.get("https://www.nba.com/games?date=" + year);
                excelList.addAll(gameDataList);
            }   /** div for 문 종료 */

            System.out.println("[excel insert]");
            int rowNum = 1;
            for (String[] rowData : excelList) {
                Row row = sheet.createRow(rowNum++);
                int colNum = 0;
                for (String cellData : rowData) {
                    Cell cell = row.createCell(colNum++);
                    cell.setCellValue(cellData);
                }
            }
        }
        System.out.println(" !!! excel insert end & download start !!! ");

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

    public boolean isFirstCharacterDigit(String value) {
        if (value != null && value.length() > 0) {
            char firstChar = value.charAt(0);
            return Character.isDigit(firstChar);
        }
        return false;
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
        driver.get("https://www.nba.com/games?date=" + year); //1946-11-02

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
                Sheet sheet = workbook.createSheet((i + 1) + " NBA GAMES");
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
                    for (int j = 0; j < header.length; j++) {
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
                for (WebElement webRow : headRow) {
                    List<WebElement> columns = webRow.findElements(By.tagName("th"));
                    int colNum = 0;
                    excelRow = sheet.createRow(++rowLevel);
                    for (WebElement column : columns) {
                        String text = (String) js.executeScript("return arguments[0].innerText;", column);
                        System.out.println("text: " + text);
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
                for (int j = 0; j < summarySecondHeader.length; j++) {
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

                System.out.println(i + "===> 번 째 완료");

                // 5초 대기
                try {
                    Thread.sleep(5000);
                } catch (InterruptedException e) {
                    e.printStackTrace();
                }

                System.out.println("===> back 이동");
                driver.get("https://www.nba.com/games?date=" + year);

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


