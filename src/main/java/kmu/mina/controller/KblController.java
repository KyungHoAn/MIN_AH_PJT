package kmu.mina.controller;

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
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.NoSuchElementException;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@Controller
public class KblController {

    private final String driverPath = new File("src/main/resources/driver/chromedriver.exe").getAbsolutePath();


    @PostMapping("/kblPlayerTotalYear")
    public ResponseEntity<InputStreamResource> getKblPlayerTotalYear(@RequestParam("kblPlayerYear") String year) throws Exception {
        System.out.println("[KBL 선수 년도별 총합 데이터 다운로드]");
        System.out.println("year: "+year);

        HttpHeaders headers = new HttpHeaders();

        Workbook workbook = null;
        System.setProperty("webdriver.chrome.driver", driverPath);
        ChromeOptions options = new ChromeOptions();

        options.addArguments("--remote-allow-origins=*");
        options.addArguments("--disable-popup-blocking");       // 팝업안띄움
        options.addArguments("--disable-gpu");                  // gpu 비활성화
        options.addArguments("--disable-images");
        options.addArguments("headless");                       // 브라우저 안띄움
        options.addArguments("--no-sandbox");
        options.addArguments("--blink-settings=imagesEnabled=false"); // 이미지 다운 안받음

//        String kblDate = year + "-" + String.format("%02d", day);

        WebDriver driver = new ChromeDriver(options);

        driver.get("https://kbl.or.kr/game/archive-player");

        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(7));
        workbook = new XSSFWorkbook();

        System.out.println("[KBL 공식 game download start");
        headers.add("Content-Disposition", "attachment; filename="+year+".xlsx");

        try {
            System.out.println("try - catch");
//            Rank 	PLAYER 	TEAM	PTS	FGM	FGA	FG%	FTA	FTA	FT%	PP	PPA	PP%	OFF	DEF	TOT	AST	TO 	STL	BS	PF
            String[] header = {"Rank","PLAYER","TEAM","PTS","FGM","FGA","FG%","FT","FTA","FT%","PP","PPA","PP%","OFF","DEF","TOT","AST","TO","STL","BS","PF"};

            Sheet sheet = workbook.createSheet("KBL PLAYER_YEAR");
            Row excelRow = sheet.createRow(0);

            for(int i=0; i<header.length; i++) {
                Cell cell = excelRow.createCell(i);
                cell.setCellValue(header[i]);
            }

            List<String[]> excelList = new ArrayList<>();

            // 두 번째 <li> 요소의 <select> 태그 찾기
            WebElement secondSelect = driver.findElement(By.xpath("//ul[@class='filter-wrap']/li[2]/select"));

            // Select 객체 생성 및 해당 옵션 선택
            Select select = new Select(secondSelect);
            select.selectByValue(year);

            // 구분 선택
            WebElement lastLi = driver.findElement(By.cssSelector("ul.filter-wrap > li:last-child"));
            WebElement selectElement = lastLi.findElement(By.tagName("select"));
            Select selectAccum = new Select(selectElement);

            selectAccum.selectByValue("accum");

            try {
                Thread.sleep(3000);
            } catch (InterruptedException e) {
                e.printStackTrace();
            }

            // 테이블의 tbody 요소 찾기
            WebElement tbody = driver.findElement(By.cssSelector("div.archive-player-table01 table tbody"));

            // tbody 안의 모든 tr 요소 찾기
            List<WebElement> rows = tbody.findElements(By.tagName("tr"));

            // 각 행(tr)을 순회하면서 열(td) 값 출력
            for (WebElement row : rows) {
                List<WebElement> columns = row.findElements(By.tagName("td"));
                String[] cellData = new String[header.length];
                int i=0;
                for (WebElement column : columns) {
//                    System.out.print(column.getText() + "\t");
                    cellData[i] = column.getText();
                    i++;
                }
//                System.out.println();
                excelList.add(cellData);
            }

            System.out.println("===> main part ");

            // 테이블의 tbody 요소 찾기
            WebElement tbodyMain = driver.findElement(By.cssSelector("div.top-scroll-table.archive-player-table01 table tbody"));

            // tbody 안의 모든 tr 요소 찾기
            List<WebElement> rowsMain = tbodyMain.findElements(By.tagName("tr"));

            // 각 행(tr)을 순회하면서 열(td) 값 출력
            int a = 0;
            for (WebElement row : rowsMain) {
                List<WebElement> columns = row.findElements(By.tagName("td"));
                int i=3;
                for (WebElement column : columns) {
//                    System.out.print(column.getText() + "\t");
                    excelList.get(a)[i] = column.getText();
                    i++;
                }
//                System.out.println();
                a++;
            }

            System.out.println("final >>>>> ");

            int rowNum = 1;
            for(String[] data : excelList) {
                Row row = sheet.createRow(rowNum++);
                int colNum = 0;
                for(String colum : data) {
                    Cell cell = row.createCell(colNum++);
                    cell.setCellValue(colum);
                }
            }

        } catch (Exception e) {
            System.out.println("[ERROR] : "+e.getMessage());
            e.printStackTrace();
        }

        System.out.println("[KBL 년도별 선수 데이터 다운로드");
        headers.add("Content-Disposition", "attachment; filename="+year+".xlsx");

        // 엑셀 파일을 바이트 배열로 변환
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        workbook.write(out);
        out.flush();
        out.close();
        workbook.close();

        ByteArrayInputStream in = new ByteArrayInputStream(out.toByteArray());
        return ResponseEntity.ok()
                .headers(headers)
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .body(new InputStreamResource(in));
    }

    @PostMapping("/kblMainDownload")
    public ResponseEntity<InputStreamResource> getKblMainDownload(@RequestParam("kblMainYear") String year) throws IOException {
        System.out.println("[KBL 공식 hoem page game download]");
        System.out.println("year: " + year);
        String[] splitDate = year.split("-");
        System.out.println("year: "+ splitDate[0]);
        System.out.println("year: "+ splitDate[1]);
        HttpHeaders headers = new HttpHeaders();

        Workbook workbook = null;
        String chromDriver = new File("src/main/resources/driver/chromedriver.exe").getAbsolutePath();
        System.setProperty("webdriver.chrome.driver", chromDriver);
        ChromeOptions options = new ChromeOptions();

        options.addArguments("--remote-allow-origins=*");
        options.addArguments("--disable-popup-blocking");       // 팝업안띄움
        options.addArguments("--disable-gpu");                  // gpu 비활성화
        options.addArguments("--disable-images");
//        options.addArguments("headless");                       // 브라우저 안띄움
        options.addArguments("--no-sandbox");
        options.addArguments("--blink-settings=imagesEnabled=false"); // 이미지 다운 안받음

//        String kblDate = year + "-" + String.format("%02d", day);

        WebDriver driver = new ChromeDriver(options);

        driver.get("https://www.kbl.or.kr/game/schedule-list");

        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(5));
        workbook = new XSSFWorkbook();

        System.out.println("[KBL 공식 game download start");
        headers.add("Content-Disposition", "attachment; filename="+year+".xlsx");

        try {
            String[] header = {"YEAR", "MONTH", "DAY", "TEAM", "HOME", "AWAY", "L(0)/W(1)", "PLAYER", "POSITION", "MIN", "1Q", "2Q", "3Q", "4Q", "OT1", "OT2", "TOTAL", "SCORE", "TOT_FGM" , "TOT_FGA", "TOT_FG%", "TOT_3PM", "TOT_3PA", "TOT_3P%", "TOT_FTM", "TOT_FTA", "TOT_FT%", "TOT_OREB", "TOT_DREB", "TOT_REB",	"TOT_AST", "TOT_STL", "TOT_BLK", "TOT_TO", "TOT_PF", "TOT_PTS", "TOT_+/-", "FGM", "FGA", "FG%", "3PM", "3PA", "3P%", "FTM", "FTA", "FT%", "OREB", "DREB", "REB", "AST", "STL", "BLK", "TO", "PF"};

            // No, Name,  Min, Pts, 2PT(M/a) 2pt%, 3pt(M/a), 3pt%, FG(M/A) M/A	%	M/A	%	M/A	%	M/A	%	OR	DR	TOT DK	AST	TO	Stl(v)	BS(BLK v)	PF(v)	FO	PP
            year = "2023";
            String month = "05";

            // 첫 번째 select 요소를 찾고 값 설정
            WebElement yearSelectElement = wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector("ul.controller-wrap select:nth-of-type(1)")));
            Select yearSelect = new Select(yearSelectElement);
            yearSelect.selectByValue(year);

            // JavaScriptExecutor를 사용하여 두 번째 select 요소를 찾고 값 설정
            JavascriptExecutor js = (JavascriptExecutor) driver;
            WebElement monthSelectElement = (WebElement) js.executeScript(
                    "return document.querySelectorAll('ul.controller-wrap select')[1]"
            );

            // 두 번째 select 요소에 값을 설정
            if (monthSelectElement != null) {
                Select monthSelect = new Select(monthSelectElement);
                monthSelect.selectByValue(month);
            } else {
                System.out.println("Month select element not found.");
            }

            // 페이지가 완전히 로드될 때까지 대기
            wait.until(ExpectedConditions.jsReturnsValue("return document.readyState == 'complete'"));

            // con-box 요소들이 나타날 때까지 대기
            List<WebElement> conBoxElements = wait.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.cssSelector("div.contents div.con-box")));
            int conBoxCount = conBoxElements.size();
            System.out.println("Number of con-box elements: " + conBoxCount);

            if (conBoxCount > 0) {
                for(int i=0; i<conBoxCount; i++) {
                    System.out.println("====> conBox count : " + i);

                    conBoxElements = wait.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.cssSelector("div.contents div.con-box")));
                    // 첫 번째 con-box 안의 button.pb 요소 찾기
                    WebElement firstConBox = conBoxElements.get(i);
                    System.out.println("====> conBox count : 1" + i);
                    WebElement buttonPb = firstConBox.findElement(By.cssSelector("button.pb"));
                    System.out.println("====> conBox count : 2" + i);
                    // 버튼 클릭
                    buttonPb.click();
                    System.out.println("Clicked the button.pb in the first con-box.");

                    // 특정 요소가 로드될 때까지 기다림
                    WebElement summaryTable = wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector("div.summary-table table tbody")));

                    // 팀별 점수 테이블의 데이터 가져오기
                    List<WebElement> rows = summaryTable.findElements(By.tagName("tr"));

                    for (WebElement row : rows) {
                        List<WebElement> cells = row.findElements(By.tagName("td"));
                        String team = cells.get(0).getText();
                        String q1 = cells.get(1).getText();
                        String q2 = cells.get(2).getText();
                        String q3 = cells.get(3).getText();
                        String q4 = cells.get(4).getText();
                        String eq = cells.get(5).getText();
                        String total = cells.get(6).getText();

                        System.out.println("Team: " + team);
                        System.out.println("1Q: " + q1);
                        System.out.println("2Q: " + q2);
                        System.out.println("3Q: " + q3);
                        System.out.println("4Q: " + q4);
                        System.out.println("EQ: " + eq);
                        System.out.println("TOTAL: " + total);
                        System.out.println("----------");
                    }

                    System.out.println("first table data ===========================");

                    // playerDetail 탭을 찾아서 클릭
                    WebElement playerDetailTab = wait.until(ExpectedConditions.elementToBeClickable(By.cssSelector("ul.tab-style01 li[data-key='playerDetail']")));
                    playerDetailTab.click();

                    // archive-team-table01-wrap 클래스가 있는 첫 번째 div 요소 내부의 테이블 데이터 가져오기
//            WebElement tableDiv = driver.findElement(By.xpath("(//div[@class='archive-team-table01-wrap'])[1]"));
                    WebElement tableDivFirst = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("(//div[@class='archive-team-table01-wrap'])[1]")));

                    // tableDiv에서 tbody 안의 모든 텍스트 데이터 가져오기
                    List<WebElement> rowsFirst = tableDivFirst.findElements(By.cssSelector("table tbody tr"));

                    // 각 행을 순회하며 데이터 출력
                    for (WebElement row : rowsFirst) {
                        List<WebElement> cells = row.findElements(By.tagName("td"));
                        for (WebElement cell : cells) {
                            System.out.print(cell.getText() + "\t");
                        }
                        System.out.println(); // 다음 줄로 넘어감
                    }

                    System.out.println("second table data ===========================");

                    WebElement tableDivSecond = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("(//div[@class='archive-team-table01-wrap'])[2]")));

                    // tableDiv에서 tbody 안의 모든 텍스트 데이터 가져오기
                    List<WebElement> rowsSecond = tableDivSecond.findElements(By.cssSelector("table tbody tr"));

                    // 각 행을 순회하며 데이터 출력
                    for (WebElement row : rowsSecond) {
                        List<WebElement> cells = row.findElements(By.tagName("td"));
                        for (WebElement cell : cells) {
                            System.out.print(cell.getText() + "\t");
                        }
                        System.out.println(); // 다음 줄로 넘어감
                    }

                    driver.navigate().back();

                    try {
                        Thread.sleep(5000);
                    } catch (InterruptedException e) {
                        e.printStackTrace();
                    }

                }
            } else {
                System.out.println("No con-box elements found.");
            }

        } catch (Exception e){
            e.printStackTrace();
            System.out.println("[ERROR] :" + e.getMessage());
        }

        // 엑셀 파일을 바이트 배열로 변환
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        workbook.write(out);
        out.flush();
        out.close();
        workbook.close();

        ByteArrayInputStream in = new ByteArrayInputStream(out.toByteArray());
        return ResponseEntity.ok()
                .headers(headers)
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .body(new InputStreamResource(in));
    }

    @PostMapping("/kblGameDownload")
    public ResponseEntity<InputStreamResource> getKblGameDownload(@RequestParam("kblYear") String year) throws IOException {
        System.out.println("[KBL GAME DATA DOWNLOAD START]");
        System.out.println("year : "+year);
        String[] splitDate = year.split("-");
//        System.out.println("day : "+day);
        HttpHeaders headers = new HttpHeaders();
        String excelName;

        Workbook workbook = null;
        String chromDriver = new File("src/main/resources/driver/chromedriver.exe").getAbsolutePath();
        System.setProperty("webdriver.chrome.driver", chromDriver);
        ChromeOptions options = new ChromeOptions();

        options.addArguments("--remote-allow-origins=*");
        options.addArguments("--disable-popup-blocking");       // 팝업안띄움
        options.addArguments("--disable-gpu");                  // gpu 비활성화
        options.addArguments("--disable-images");
        options.addArguments("headless");                       // 브라우저 안띄움
        options.addArguments("--no-sandbox");
        options.addArguments("--blink-settings=imagesEnabled=false"); // 이미지 다운 안받음

//        String kblDate = year + "-" + String.format("%02d", day);

        WebDriver driver = new ChromeDriver(options);
        boolean errorFlagAndStopFlag = false;

//        int errorCount = 0;
//        while(!errorFlagAndStopFlag) {
//        }
//            driver.get("https://m.sports.naver.com/basketball/schedule/index?category=kbl&date=2023-11-02");

        driver.get("https://m.sports.naver.com/basketball/schedule/index?category=kbl&date="+year);

        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(2));
        workbook = new XSSFWorkbook();

        boolean isPresent = isElementPresent(driver, By.className("ScheduleLeagueType_match_list_container__1v4b0"));
        System.out.println(year + " 경기 유무 : " + isPresent);

        try {
            if (isPresent) {
                WebElement container = wait.until(ExpectedConditions.visibilityOfElementLocated(By.className("ScheduleLeagueType_match_list_container__1v4b0")));
                /** 해당 요소 바로 아래에 있는 div 요소 가져오기 */
                List<WebElement> divElements = container.findElements(By.xpath("./div"));
                System.out.println("날짜 수: " + divElements.size());

//                String[] header = {"Year", "Month", "Day", "Attendance", "TEAM", "HOME (홈)", "AWAY (어웨이)", "L(0)/W(1)", "PLAYER", "MIN", "Q1", "Q2", "Q3", "Q4", "OT1", "OT2", "FINAL", "TOT_FGM", "TOT_FGA", "TOT_FG%", "TOT_3PM", "TOT_3PA", "TOT_3P%", "TOT_FTM", "TOT_FTA", "TOT_FT%", "TOT_OREB", "TOT_DREB", "TOT_REB", "TOT_AST", "TOT_STL", "TOT_BLK", "TOT_TO", "TOT_PF", "TOT_PTS", "TOT_+/-", "FGM", "FGA", "FG%", "3PM", "3PA", "3P%", "FTM", "FTA", "FT%", "OREB", "DREB", "REB", "AST", "STL", "BLK", "TO", "PF", "PTS", "+/-"};

                JavascriptExecutor js = (JavascriptExecutor) driver;
                String[] header = {"YEAR", "MONTH", "DAY", "TEAM", "HOME", "AWAY", "L(0)/W(1)", "PLAYER", "POSITION", "MIN", "1Q", "2Q", "3Q", "4Q", "OT1", "OT2", "TOTAL", "SCORE", "TOT_FGM" , "TOT_FGA", "TOT_FG%", "TOT_3PM", "TOT_3PA", "TOT_3P%", "TOT_FTM", "TOT_FTA", "TOT_FT%", "TOT_OREB", "TOT_DREB", "TOT_REB",	"TOT_AST", "TOT_STL", "TOT_BLK", "TOT_TO", "TOT_PF", "TOT_PTS", "TOT_+/-", "FGM", "FGA", "FG%", "3PM", "3PA", "3P%", "FTM", "FTA", "FT%", "OREB", "DREB", "REB", "AST", "STL", "BLK", "TO", "PF"};
                Sheet sheet = workbook.createSheet("KBL GAMES");
                Row excelRow = sheet.createRow(0);
                for (int i=0; i< header.length; i++) {
                    Cell cell = excelRow.createCell(i);
                    cell.setCellValue(header[i]);
                }
                List<String[]> excelList = new ArrayList<>();

                for (int i = 0; i < divElements.size(); i++) {
                    WebElement gameCardsContainer = wait.until(ExpectedConditions.visibilityOfElementLocated(By.className("ScheduleLeagueType_match_list_container__1v4b0")));
                    WebElement secondCardContainer = gameCardsContainer.findElements(By.className("ScheduleLeagueType_match_list_group__18ML9")).get(i);

                    WebElement emElement = secondCardContainer.findElement(By.cssSelector(".ScheduleLeagueType_group_title__S2Z_g .ScheduleLeagueType_title__2Kalm"));
                    String gameDate = emElement.getText();
                    System.out.println("game 시작 날짜: " + gameDate);
                    Pattern pattern = Pattern.compile("\\d+일");
                    Matcher matcher = pattern.matcher(gameDate);

                    String day = "";
                    if (matcher.find()) {
                        day = matcher.group().replace("일", "");
                    }
                    System.out.println("Extracted day: " + day);

                    try {
                        Thread.sleep(3000);
                    } catch (InterruptedException e) {
                        e.printStackTrace();
                    }
                    WebElement thirdCardContainer = secondCardContainer.findElement(By.className("ScheduleLeagueType_match_list__1-n6x"));
                    List<WebElement> liElements = thirdCardContainer.findElements(By.xpath("./li[contains(@class, 'MatchBox_match_item__3_D0Q') and contains(@class, 'type_end')]"));

                    for (int j = 0; j < liElements.size(); j++) {
                        WebElement liElement = liElements.get(j);
                        WebElement recordLink = liElement.findElement(By.cssSelector("a[href*='/record']"));
                        JavascriptExecutor executor = (JavascriptExecutor) driver;
                        executor.executeScript("arguments[0].click();", recordLink);

                        List<String[]> gameDataList = new ArrayList<>();

                        try {
                            Thread.sleep(5000);
                        } catch (InterruptedException e) {
                            e.printStackTrace();
                        }

                        WebElement teamElement = wait.until(ExpectedConditions.visibilityOfElementLocated(By.className("ScoreBox_title__2jVlJ")));

                        // 팀명 요소에서 홈팀과 원정팀 이름 가져오기
                        WebElement homeTeamElement = teamElement.findElement(By.className("ScoreBox_home__2uCuR"));
                        WebElement awayTeamElement = teamElement.findElement(By.className("ScoreBox_away__26sht"));

                        // 팀 이름을 String으로 저장
                        String homeTeamName = homeTeamElement.getText();
                        String awayTeamName = awayTeamElement.getText();

                        // 쿼터수
                        WebElement tableElement = wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector("div.ScoreBox_round__1sOLq table.ScoreBox_board_table__3V6uh")));

                        WebElement firstRow = tableElement.findElement(By.xpath(".//tbody/tr[1]"));
                        List<WebElement> firstRowCells = firstRow.findElements(By.tagName("td"));

                        WebElement secondRow = tableElement.findElement(By.xpath(".//tbody/tr[2]"));
                        List<WebElement> secondRowCells = secondRow.findElements(By.tagName("td"));


                        String[] homeQ = new String[6];
                        System.out.println("First Row Values:");
                        int hqNum = 0;
                        for (WebElement cell : firstRowCells) {
                            System.out.println(cell.getText());
                            homeQ[hqNum] = cell.getText();
                            hqNum++;
                        }

                        String[] arrayQ = new String[6];
                        System.out.println("Second Row Values:");
                        int aqNum = 0;
                        for (WebElement cell : secondRowCells) {
                            System.out.println(cell.getText());
                            arrayQ[aqNum] = cell.getText();
                            aqNum++;
                        }

                        // Total
                        tableElement = wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector("div.ScoreBox_result__3atD0 table.ScoreBox_board_table__3V6uh")));

                        firstRow = tableElement.findElement(By.xpath(".//tbody/tr[1]"));
                        firstRowCells = firstRow.findElements(By.tagName("td"));

                        secondRow = tableElement.findElement(By.xpath(".//tbody/tr[2]"));
                        secondRowCells = secondRow.findElements(By.tagName("td"));

                        int homeTeamScore = 0;
                        int awayTeamScore = 0;
                        for (WebElement cell : firstRowCells) {
                            homeTeamScore = Integer.parseInt(cell.getText());
                        }
                        for (WebElement cell : secondRowCells) {
                            awayTeamScore = Integer.parseInt(cell.getText());
                        }


                        String finalHomeScore = (homeTeamScore > awayTeamScore) ? "1" : "0";
                        String finalAwayScore = (awayTeamScore > homeTeamScore) ? "1" : "0";

                        // KBL 선수 상세 데이터 [2개 Table]
                        System.out.println("[KBL 선수 상세 데이터 시작]");
                        // 모든 PlayerRecord_player_record_area__1HO0u div 요소 찾기
                        List<WebElement> playerRecordDivs = wait.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.className("PlayerRecord_player_record_area__1HO0u")));

                        String[] excelData;


                        boolean homeCheck = true;
                        // 각 playerRecordDiv 안에 있는 player_list를 처리
                        for (WebElement playerRecordDiv : playerRecordDivs) {
                            // player_list 요소 찾기
                            WebElement playerList = playerRecordDiv.findElement(By.className("player_list"));
                            List<WebElement> playerItems = playerList.findElements(By.className("PlayerRecord_player_item__16fJR"));

                            // 각 player_item에서 필요한 정보 추출
                            for (WebElement playerItem : playerItems) {
                                excelData = new String[header.length];
                                String playerNumber = playerItem.findElement(By.className("PlayerRecord_number__koWes")).getText(); // 선수 번호
                                String playerName = playerItem.findElement(By.className("PlayerRecord_name__2aL7m")).getText(); // 선수 이름
                                String playerPosition = playerItem.findElement(By.className("PlayerRecord_position__1O96P")).getText(); // 선수 포지션
                                String playerUrl = playerItem.findElement(By.tagName("a")).getAttribute("href");    // 선수 URL

                                // 정보 출력
//                                System.out.println("Number: " + playerNumber);
//                                System.out.println("Name: " + playerName);
//                                System.out.println("Position: " + playerPosition);
//                                System.out.println("URL: " + playerUrl);
//                                String[] playerData = {playerName, playerPosition};
                                Arrays.fill(excelData, "-");
                                excelData[0] = splitDate[0];
                                excelData[1] = splitDate[1];
                                excelData[2] = day;
                                if (homeCheck) {
                                    excelData[3] = homeTeamName;
                                    excelData[6] = finalHomeScore;
                                    excelData[16] = String.valueOf(homeTeamScore);
                                } else {
                                    excelData[3] = awayTeamName;
                                    excelData[6] = finalAwayScore;
                                    excelData[16] = String.valueOf(awayTeamScore);
                                }
                                excelData[4] = homeTeamName;
                                excelData[5] = awayTeamName;
                                excelData[7] = playerName;
                                excelData[8] = playerPosition;
                                gameDataList.add(excelData);
                            }
                            System.out.println("============= TABLE SECTOR =============");
                            homeCheck = false;
                        }

                        homeCheck = true;
                        System.out.println("[KBL 선수 이름 데이터 종료 ]");

                        ArrayList<String[]> homeTeamData = new ArrayList<>();
                        System.out.println("===> home Team Data size : "+homeTeamData.size());
                        ArrayList<String[]> awayTeamData = new ArrayList<>();

                        System.out.println("[KBL 선수별 데이터]");
                        List<WebElement> tables = wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.className("PlayerRecord_record_table__52KqA")));

                        js = (JavascriptExecutor) driver;
                        for (int a = 0; a < tables.size(); a++) {
                            WebElement table = tables.get(a);
                            System.out.println("Table " + (a + 1) + ":");
                            List<WebElement> tbodys = table.findElements(By.tagName("tbody"));

                            for (WebElement tbody : tbodys) {
                                List<WebElement> rows = tbody.findElements(By.tagName("tr"));
                                int number = 0;

                                for (WebElement row : rows) {
                                    List<WebElement> cells = row.findElements(By.tagName("td"));
                                    String[] homeData = new String[16];
                                    String[] awayData = new String[16];
                                    int cellNum = 0;
                                    for (WebElement cell : cells) {
                                        String cellText = (String) js.executeScript("return arguments[0].textContent;", cell);
                                        System.out.print(cellText.trim() + "\t");
                                        if (homeCheck) {
                                            homeData[cellNum] = cellText.trim();
                                        } else {
                                            awayData[cellNum] = cellText.trim();
                                        }
                                        cellNum++;
                                    }
                                    if(homeCheck) {
                                        homeTeamData.add(homeData);
                                    } else {
                                        awayTeamData.add(awayData);
                                    }
                                    System.out.println();
                                }
                                System.out.println("========================================");
                                homeCheck = false;
                            }
                        }

                        int HomeArrayNum = 0;
                        for (String[] homeArray : homeTeamData) {
                            int num = 0;
                            String[] fgmData;
                            String[] tP;    // 3점
                            String[] ft;
                            for(String result : homeArray) {
                                if(num == 0) {
                                    gameDataList.get(HomeArrayNum)[9] = result;
                                }if(num == 1) {
                                    gameDataList.get(HomeArrayNum)[17] = result;
                                }if(num == 2) {
                                    gameDataList.get(HomeArrayNum)[48] = result;
                                }if(num == 3) {
                                    gameDataList.get(HomeArrayNum)[49] = result;
                                }if(num == 4) {
                                    gameDataList.get(HomeArrayNum)[50] = result;
                                }if(num == 5) {
                                    gameDataList.get(HomeArrayNum)[51] = result;
                                }if(num == 6) {
                                    fgmData = result.split("/");
                                    gameDataList.get(HomeArrayNum)[37] = fgmData[0];
                                    gameDataList.get(HomeArrayNum)[38] = fgmData[1];
                                }if(num == 7) {
                                    gameDataList.get(HomeArrayNum)[39] = result;
                                }if(num == 8) {
                                    tP = result.split("/");
                                    gameDataList.get(HomeArrayNum)[40] = tP[0];
                                    gameDataList.get(HomeArrayNum)[41] = tP[1];
                                }if(num == 9) {
                                    gameDataList.get(HomeArrayNum)[42] = result;
                                }if(num == 10) {
                                    ft = result.split("/");
                                    gameDataList.get(HomeArrayNum)[43] = ft[0];
                                    gameDataList.get(HomeArrayNum)[44] = ft[1];
                                }if(num == 11) {
                                    gameDataList.get(HomeArrayNum)[45] = result;
                                }if(num == 12) {
                                    gameDataList.get(HomeArrayNum)[46] = result;
                                }if(num == 13) {
                                    gameDataList.get(HomeArrayNum)[47] = result;
                                }if(num == 14) {
                                    gameDataList.get(HomeArrayNum)[52] = result;
                                }if(num == 15) {
                                    gameDataList.get(HomeArrayNum)[53] = result;
                                }
                                num ++;
                            }
                            HomeArrayNum++;
                        }

                        int AwayArrayNum = homeTeamData.size();
                        for (String[] awayArray : awayTeamData) {
                            int num = 0;
                            String[] fgmData;
                            String[] tP;    // 3점
                            String[] ft;
                            for(String result : awayArray) {
                                if(num == 0) {
                                    gameDataList.get(AwayArrayNum)[9] = result;
                                }if(num == 1) {
                                    gameDataList.get(AwayArrayNum)[17] = result;
                                }if(num == 2) {
                                    gameDataList.get(AwayArrayNum)[48] = result;
                                }if(num == 3) {
                                    gameDataList.get(AwayArrayNum)[49] = result;
                                }if(num == 4) {
                                    gameDataList.get(AwayArrayNum)[50] = result;
                                }if(num == 5) {
                                    gameDataList.get(AwayArrayNum)[51] = result;
                                }if(num == 6) {
                                    fgmData = result.split("/");
                                    gameDataList.get(AwayArrayNum)[37] = fgmData[0];
                                    gameDataList.get(AwayArrayNum)[38] = fgmData[1];
                                }if(num == 7) {
                                    gameDataList.get(AwayArrayNum)[39] = result;
                                }if(num == 8) {
                                    tP = result.split("/");
                                    gameDataList.get(AwayArrayNum)[40] = tP[0];
                                    gameDataList.get(AwayArrayNum)[41] = tP[1];
                                }if(num == 9) {
                                    gameDataList.get(AwayArrayNum)[42] = result;
                                }if(num == 10) {
                                    ft = result.split("/");
                                    gameDataList.get(AwayArrayNum)[43] = ft[0];
                                    gameDataList.get(AwayArrayNum)[44] = ft[1];
                                }if(num == 11) {
                                    gameDataList.get(AwayArrayNum)[45] = result;
                                }if(num == 12) {
                                    gameDataList.get(AwayArrayNum)[46] = result;
                                }if(num == 13) {
                                    gameDataList.get(AwayArrayNum)[47] = result;
                                }if(num == 14) {
                                    gameDataList.get(AwayArrayNum)[52] = result;
                                }if(num == 15) {
                                    gameDataList.get(AwayArrayNum)[53] = result;
                                }
                                num ++;
                            }
                            AwayArrayNum++;
                        }

                        if(homeTeamData.size() > 0) {
                            // Totals
                            int totalPoints = 0;
                            int totalORB = 0;
                            int totalDRB = 0;
                            int totalAST = 0;
                            int totalSTL = 0;
                            int totalBLK = 0;
                            int totalTO = 0;
                            int totalPF = 0;
                            int totalFGM = 0;
                            int totalFGA = 0;
                            int total3PM = 0;
                            int total3PA = 0;
                            int totalFTM = 0;
                            int totalFTA = 0;

                            for (String[] playerStats : homeTeamData) {
                                totalPoints += Integer.parseInt(playerStats[1]);
                                totalORB += Integer.parseInt(playerStats[12]);  // 공격 리바운드
                                totalDRB += Integer.parseInt(playerStats[13]);  // 수비 리바운드
                                totalAST += Integer.parseInt(playerStats[3]);
                                totalSTL += Integer.parseInt(playerStats[4]);
                                totalBLK += Integer.parseInt(playerStats[5]);
                                totalTO += Integer.parseInt(playerStats[14]);
                                totalPF += Integer.parseInt(playerStats[15]);

                                String[] fg = playerStats[6].split("/");
                                totalFGM += Integer.parseInt(fg[0]);
                                totalFGA += Integer.parseInt(fg[1]);

                                String[] threePt = playerStats[8].split("/");
                                total3PM += Integer.parseInt(threePt[0]);
                                total3PA += Integer.parseInt(threePt[1]);

                                String[] ft = playerStats[10].split("/");
                                totalFTM += Integer.parseInt(ft[0]);
                                totalFTA += Integer.parseInt(ft[1]);
                            }

                            int totalREB = totalORB + totalDRB;
                            double totalFGPct = (totalFGM / (double) totalFGA) * 100;
                            double total3PPct = (total3PM / (double) total3PA) * 100;
                            double totalFTPct = (totalFTM / (double) totalFTA) * 100;

//                            System.out.println("Total PTS: " + totalPoints); // PTS 총 득점 ok
//                            System.out.println("Total ORB: " + totalORB);       // 공격 리바운드의 개수 7
//                            System.out.println("Total DRB: " + totalDRB);       // 수비 리바운드 개수 24
//                            System.out.println("Total REB: " + totalREB);       // 총 리바운드의 개수 (공격 리바운드 + 수비 리바운드)
//                            System.out.println("Total AST: " + totalAST);       // 어시스트의 개수 ok
//                            System.out.println("Total STL: " + totalSTL);       // 스틸의 개수 ok
//                            System.out.println("Total BLK: " + totalBLK);       // 블록의 개수 ok
//                            System.out.println("Total TO: " + totalTO);         // 턴오버의 개수 ok
//                            System.out.println("Total PF: " + totalPF);         // 개인 파울의 개수 ok
//                            System.out.println("Total FGM: " + totalFGM);       // 성공한 야투의 개수 ok
//                            System.out.println("Total FGA: " + totalFGA);       // 총 시도 야투 개수 ok
//                            System.out.println("Total FG%: " + String.format("%.1f", totalFGPct));      // FG % ok
//                            System.out.println("Total 3PM: " + total3PM);       // 성공 3점 슛 ok
//                            System.out.println("Total 3PA: " + total3PA);       // 시도 3점 슛  ok
//                            System.out.println("Total 3P%: " + String.format("%.1f", total3PPct));  // 3P% ok
//                            System.out.println("Total FTM: " + totalFTM);       // 성공 자유투 개수  ok
//                            System.out.println("Total FTA: " + totalFTA);       // 시도 자유투 개수  ok
//                            System.out.println("Total FT%: " + String.format("%.1f", totalFTPct));  // ok

                            // TODO +/- 선수의 출전 시간 동안 팀의 득점 차이는 공식이 없는 것 같다? 나오지 않는 부분인것 같음
                            // TODO OT1 - OT2 는 나오지 ㅇ낳는 부분인듯
                            // 순서 : FGM, FGA, FG%, 3PM, 3PA, 3P%, FTM, FTA, FT%, OREB(x), DREB(x), REB, AST, STL, BLK, TO, PF, PTS, +/-(x)
                            for(int a=0; a<homeTeamData.size(); a++) {
                                if(homeQ[0] != null) {
                                    gameDataList.get(a)[10] = homeQ[0];
                                }if(homeQ[1] != null) {
                                    gameDataList.get(a)[11] = homeQ[1];
                                }if(homeQ[2] != null) {
                                    gameDataList.get(a)[12] = homeQ[2];
                                }if(homeQ[3] != null) {
                                    gameDataList.get(a)[13] = homeQ[3];
                                }
                                gameDataList.get(a)[18] = String.valueOf(totalFGM);
                                gameDataList.get(a)[19] = String.valueOf(totalFGA);
                                gameDataList.get(a)[20] = String.valueOf(totalFGPct);
                                gameDataList.get(a)[21] = String.valueOf(total3PM);
                                gameDataList.get(a)[22] = String.valueOf(total3PA);
                                gameDataList.get(a)[23] = String.valueOf(total3PPct);
                                gameDataList.get(a)[24] = String.valueOf(totalFTM);
                                gameDataList.get(a)[25] = String.valueOf(totalFTA);
                                gameDataList.get(a)[26] = String.valueOf(totalFTPct);
                                gameDataList.get(a)[27] = String.valueOf(totalORB);
                                gameDataList.get(a)[28] = String.valueOf(totalDRB);
                                gameDataList.get(a)[29] = String.valueOf(totalREB);
                                gameDataList.get(a)[30] = String.valueOf(totalAST);
                                gameDataList.get(a)[31] = String.valueOf(totalSTL);
                                gameDataList.get(a)[32] = String.valueOf(totalBLK);
                                gameDataList.get(a)[33] = String.valueOf(totalTO);
                                gameDataList.get(a)[34] = String.valueOf(totalPF);
                                gameDataList.get(a)[35] = String.valueOf(totalPoints);
                            }
                        }

                        if(awayTeamData.size() > 0) {
                            // Totals
                            int totalPoints = 0;
                            int totalORB = 0;
                            int totalDRB = 0;
                            int totalAST = 0;
                            int totalSTL = 0;
                            int totalBLK = 0;
                            int totalTO = 0;
                            int totalPF = 0;
                            int totalFGM = 0;
                            int totalFGA = 0;
                            int total3PM = 0;
                            int total3PA = 0;
                            int totalFTM = 0;
                            int totalFTA = 0;

                            for (String[] playerStats : awayTeamData) {
                                totalPoints += Integer.parseInt(playerStats[1]);
                                totalORB += Integer.parseInt(playerStats[12]);  // 공격 리바운드
                                totalDRB += Integer.parseInt(playerStats[13]);  // 수비 리바운드
                                totalAST += Integer.parseInt(playerStats[3]);
                                totalSTL += Integer.parseInt(playerStats[4]);
                                totalBLK += Integer.parseInt(playerStats[5]);
                                totalTO += Integer.parseInt(playerStats[14]);
                                totalPF += Integer.parseInt(playerStats[15]);

                                String[] fg = playerStats[6].split("/");
                                totalFGM += Integer.parseInt(fg[0]);
                                totalFGA += Integer.parseInt(fg[1]);

                                String[] threePt = playerStats[8].split("/");
                                total3PM += Integer.parseInt(threePt[0]);
                                total3PA += Integer.parseInt(threePt[1]);

                                String[] ft = playerStats[10].split("/");
                                totalFTM += Integer.parseInt(ft[0]);
                                totalFTA += Integer.parseInt(ft[1]);
                            }

                            int totalREB = totalORB + totalDRB;
                            double totalFGPct = (totalFGM / (double) totalFGA) * 100;
                            double total3PPct = (total3PM / (double) total3PA) * 100;
                            double totalFTPct = (totalFTM / (double) totalFTA) * 100;

                            // TODO +/- 선수의 출전 시간 동안 팀의 득점 차이는 공식이 없는 것 같다? 나오지 않는 부분인것 같음
                            // TODO OT1 - OT2 는 나오지 ㅇ낳는 부분인듯
                            // 순서 : FGM, FGA, FG%, 3PM, 3PA, 3P%, FTM, FTA, FT%, OREB(x), DREB(x), REB, AST, STL, BLK, TO, PF, PTS, +/-(x)
                            for(int a=homeTeamData.size(); a<gameDataList.size(); a++) {
                                if(arrayQ[0] != null) {
                                    gameDataList.get(a)[10] = arrayQ[0];
                                }if(arrayQ[1] != null) {
                                    gameDataList.get(a)[11] = arrayQ[1];
                                }if(arrayQ[2] != null) {
                                    gameDataList.get(a)[12] = arrayQ[2];
                                }if(arrayQ[3] != null) {
                                    gameDataList.get(a)[13] = arrayQ[3];
                                }
                                gameDataList.get(a)[18] = String.valueOf(totalFGM);
                                gameDataList.get(a)[19] = String.valueOf(totalFGA);
                                gameDataList.get(a)[20] = String.format("%.1f", totalFGPct);
                                gameDataList.get(a)[21] = String.valueOf(total3PM);
                                gameDataList.get(a)[22] = String.valueOf(total3PA);
                                gameDataList.get(a)[23] =  String.format("%.1f", total3PPct);
                                gameDataList.get(a)[24] = String.valueOf(totalFTM);
                                gameDataList.get(a)[25] = String.valueOf(totalFTA);
                                gameDataList.get(a)[26] = String.format("%.1f", totalFTPct);
                                gameDataList.get(a)[27] = String.valueOf(totalORB);
                                gameDataList.get(a)[28] = String.valueOf(totalDRB);
                                gameDataList.get(a)[29] = String.valueOf(totalREB);
                                gameDataList.get(a)[30] = String.valueOf(totalAST);
                                gameDataList.get(a)[31] = String.valueOf(totalSTL);
                                gameDataList.get(a)[32] = String.valueOf(totalBLK);
                                gameDataList.get(a)[33] = String.valueOf(totalTO);
                                gameDataList.get(a)[34] = String.valueOf(totalPF);
                                gameDataList.get(a)[35] = String.valueOf(totalPoints);
                            }
                        }

                        System.out.println("===> total data summary >> ");
                        for (String[] dataArray : gameDataList) {
                            System.out.print("Array: ");
                            for (String data : dataArray) {
                                System.out.print(data + " ");
                            }
                            System.out.println();  // New line for each array
                        }

                        excelList.addAll(gameDataList);

                        driver.navigate().back();

                        try {
                            Thread.sleep(8000);
                        } catch (InterruptedException e) {
                            e.printStackTrace();
                        }

                        // 페이지 네비게이션 후, thirdCardContainer와 liElements를 다시 찾아 업데이트
                        gameCardsContainer = wait.until(ExpectedConditions.visibilityOfElementLocated(By.className("ScheduleLeagueType_match_list_container__1v4b0")));
                        secondCardContainer = gameCardsContainer.findElements(By.className("ScheduleLeagueType_match_list_group__18ML9")).get(i);
                        thirdCardContainer = secondCardContainer.findElement(By.className("ScheduleLeagueType_match_list__1-n6x"));
                        liElements = thirdCardContainer.findElements(By.xpath("./li[contains(@class, 'MatchBox_match_item__3_D0Q') and contains(@class, 'type_end')]"));
                    }

                    int rowNum = 1;
                    for(String[] rowData : excelList) {
                        Row row = sheet.createRow(rowNum++);
                        int colNum = 0;
                        for(String cellData : rowData) {
                            Cell cell = row.createCell(colNum++);
                            cell.setCellValue(cellData);
                        }
                    }
                }
            }
        } catch (NoSuchElementException e) {
            System.out.println("ERROR [NoSuchElementException] : " + e.getMessage());
//            errorCount++;
//            if(errorCount == 3)
//                return null;
        }
        catch (Exception e) {
            System.out.println("ERROR : " + e.getMessage());
//            errorCount++;
//            if(errorCount == 3)
//                return null;
        }

        System.out.println("[KBL excel download started]");
        headers.add("Content-Disposition", "attachment; filename="+year+".xlsx");

        // 엑셀 파일을 바이트 배열로 변환
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        workbook.write(out);
        out.flush();
        out.close();
        workbook.close();

        ByteArrayInputStream in = new ByteArrayInputStream(out.toByteArray());
        return ResponseEntity.ok()
                .headers(headers)
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .body(new InputStreamResource(in));
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

}
