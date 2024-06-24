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
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.*;

import java.io.*;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.NoSuchElementException;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@Controller
public class MainController {

    @RequestMapping("/")
    public String main() {
        return "main";
    }

    private final String driverPath = new File("src/main/resources/driver/chromedriver.exe").getAbsolutePath();

//    @GetMapping("/kblGameDownload")
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
                String[] header = {"YEAR", "MONTH", "DAY", "TEAM", "HOME", "AWAY", "L(0)/W(1)", "PLAYER", "POSITION", "MIN", "1Q", "2Q", "3Q", "4Q", "OT1", "OT2", "TOTAL", "SCORE", "TOT_FGM" , "TOT_FGA", "TOT_FG%", "TOT_3PM", "TOT_3PA", "TOT_3P%", "TOT_FTM", "TOT_FTA", "TOT_FT%", "TOT_OREB", "TOT_DREB", "TOT_REB",	"TOT_AST", "TOT_STL", "TOT_BLK", "TOT_TO", "TOT_PF", "TOT_PTS", "TOT_+/-", "FGM", "FGA", "FG%", "3PM", "3PA", "3P%", "FTM", "FTA", "FT%", "OREB", "DREB", "REB", "AST", "STL", "BLK", "TO", "PF"};
                JavascriptExecutor js = (JavascriptExecutor) driver;

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
                        System.out.println("homeTeam : " + homeTeamName);
                        System.out.println("awayTeamName : " + awayTeamName);

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

                        System.out.println("[TOTAL] >> ");
                        // Total
                        tableElement = wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector("div.ScoreBox_result__3atD0 table.ScoreBox_board_table__3V6uh")));

                        firstRow = tableElement.findElement(By.xpath(".//tbody/tr[1]"));
                        firstRowCells = firstRow.findElements(By.tagName("td"));

                        secondRow = tableElement.findElement(By.xpath(".//tbody/tr[2]"));
                        secondRowCells = secondRow.findElements(By.tagName("td"));

                        int homeTeamScore = 0;
                        int awayTeamScore = 0;
                        for (WebElement cell : firstRowCells) {
                            System.out.println(cell.getText());
                            homeTeamScore = Integer.parseInt(cell.getText());
                        }
                        for (WebElement cell : secondRowCells) {
                            System.out.println(cell.getText());
                            awayTeamScore = Integer.parseInt(cell.getText());
                        }
                        System.out.println("homeTeamScore : "+homeTeamScore);
                        System.out.println("homeTeamScore : "+awayTeamScore);


                        String finalHomeScore = (homeTeamScore > awayTeamScore) ? "1" : "0";
                        String finalAwayScore = (awayTeamScore > homeTeamScore) ? "1" : "0";
                        System.out.println("finalHomeScore : "+ finalHomeScore);
                        System.out.println("finalAwayScore : "+ finalAwayScore);

                        // KBL 선수 상세 데이터 [2개 Table]
                        System.out.println("[KBL 선수 상세 데이터 시작]");
                        // 모든 PlayerRecord_player_record_area__1HO0u div 요소 찾기
                        List<WebElement> playerRecordDivs = wait.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.className("PlayerRecord_player_record_area__1HO0u")));

//                        ArrayList<String[]> homeTeamPlayer = new ArrayList<>();
//                        ArrayList<String[]> awayTeamPlayer = new ArrayList<>();

                        String[] excelData;


                        boolean homeCheck = true;
                        // 각 playerRecordDiv 안에 있는 player_list를 처리
                        for (WebElement playerRecordDiv : playerRecordDivs) {
                            // player_list 요소 찾기
                            WebElement playerList = playerRecordDiv.findElement(By.className("player_list"));
                            List<WebElement> playerItems = playerList.findElements(By.className("PlayerRecord_player_item__16fJR"));

                            int number = 0;

                            // 각 player_item에서 필요한 정보 추출
                            for (WebElement playerItem : playerItems) {
                                excelData = new String[header.length];
                                String playerNumber = playerItem.findElement(By.className("PlayerRecord_number__koWes")).getText(); // 선수 번호
                                String playerName = playerItem.findElement(By.className("PlayerRecord_name__2aL7m")).getText(); // 선수 이름
                                String playerPosition = playerItem.findElement(By.className("PlayerRecord_position__1O96P")).getText(); // 선수 포지션
                                String playerUrl = playerItem.findElement(By.tagName("a")).getAttribute("href");    // 선수 URL

                                // 정보 출력
                                System.out.println("Number: " + playerNumber);
                                System.out.println("Name: " + playerName);
                                System.out.println("Position: " + playerPosition);
                                System.out.println("URL: " + playerUrl);
                                System.out.println("------");
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
//                        System.out.println("homeTeamData : "+ homeTeamPlayer);
//                        System.out.println("awayTeamData : "+ awayTeamPlayer);
                        // Loop through the list and print each String array
                        for (String[] dataArray : gameDataList) {
                            System.out.print("Array: ");
                            for (String data : dataArray) {
                                System.out.print(data + " ");
                            }
                            System.out.println();  // New line for each array
                        }


                        homeCheck = true;
                        System.out.println("[KBL 선수 이름 데이터 종료 ]");

                        ArrayList<String[]> homeTeamData = new ArrayList<>();
                        System.out.println("===> home Team Data size : "+homeTeamData.size());
                        ArrayList<String[]> awayTeamData = new ArrayList<>();

                        // TODO 선수 게임 데이터 가져오기
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

                        System.out.println("[Final Search KBL DATA]");
                        System.out.println("home data > " + homeTeamData);
                        System.out.println("away data > " + awayTeamData);

                        System.out.println("home data size : "+ homeTeamData.size());
                        System.out.println("away data size : "+ awayTeamData.size());

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

                        System.out.println("=> check >>>>>>>>>>>>>>>>>>>");
                        for (String[] dataArray : gameDataList) {
                            System.out.print("Array: ");
                            for (String data : dataArray) {
                                System.out.print(data + " ");
                            }
                            System.out.println();  // New line for each array
                        }
                        System.out.println("=> check >>>>>>>>>>>>>>>>>>>");

                        ArrayList<String> homeTotData = new ArrayList<>();
                        ArrayList<String> awayTotData = new ArrayList<>();
                        
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

                            System.out.println("Total PTS: " + totalPoints); // PTS 총 득점 ok
                            System.out.println("Total ORB: " + totalORB);       // 공격 리바운드의 개수 7
                            System.out.println("Total DRB: " + totalDRB);       // 수비 리바운드 개수 24
                            System.out.println("Total REB: " + totalREB);       // 총 리바운드의 개수 (공격 리바운드 + 수비 리바운드)
                            System.out.println("Total AST: " + totalAST);       // 어시스트의 개수 ok
                            System.out.println("Total STL: " + totalSTL);       // 스틸의 개수 ok
                            System.out.println("Total BLK: " + totalBLK);       // 블록의 개수 ok
                            System.out.println("Total TO: " + totalTO);         // 턴오버의 개수 ok
                            System.out.println("Total PF: " + totalPF);         // 개인 파울의 개수 ok
                            System.out.println("Total FGM: " + totalFGM);       // 성공한 야투의 개수 ok
                            System.out.println("Total FGA: " + totalFGA);       // 총 시도 야투 개수 ok
                            System.out.println("Total FG%: " + String.format("%.1f", totalFGPct));      // FG % ok
                            System.out.println("Total 3PM: " + total3PM);       // 성공 3점 슛 ok
                            System.out.println("Total 3PA: " + total3PA);       // 시도 3점 슛  ok
                            System.out.println("Total 3P%: " + String.format("%.1f", total3PPct));  // 3P% ok
                            System.out.println("Total FTM: " + totalFTM);       // 성공 자유투 개수  ok
                            System.out.println("Total FTA: " + totalFTA);       // 시도 자유투 개수  ok
                            System.out.println("Total FT%: " + String.format("%.1f", totalFTPct));  // ok

                            System.out.println(homeQ);
                            System.out.println(homeQ.length);
                            System.out.println(homeQ[0]);
                            System.out.println(homeQ[4]);
                            System.out.println(homeQ[5]);

                            if(homeQ[5] != null) {
                                System.out.println("is not null");
                            } else {
                                System.out.println("is null");
                            }



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

                            System.out.println("Total PTS: " + totalPoints); // PTS 총 득점 ok
                            System.out.println("Total ORB: " + totalORB);       // 공격 리바운드의 개수 7
                            System.out.println("Total DRB: " + totalDRB);       // 수비 리바운드 개수 24
                            System.out.println("Total REB: " + totalREB);       // 총 리바운드의 개수 (공격 리바운드 + 수비 리바운드)
                            System.out.println("Total AST: " + totalAST);       // 어시스트의 개수 ok
                            System.out.println("Total STL: " + totalSTL);       // 스틸의 개수 ok
                            System.out.println("Total BLK: " + totalBLK);       // 블록의 개수 ok
                            System.out.println("Total TO: " + totalTO);         // 턴오버의 개수 ok
                            System.out.println("Total PF: " + totalPF);         // 개인 파울의 개수 ok
                            System.out.println("Total FGM: " + totalFGM);       // 성공한 야투의 개수 ok
                            System.out.println("Total FGA: " + totalFGA);       // 총 시도 야투 개수 ok
                            System.out.println("Total FG%: " + String.format("%.1f", totalFGPct));      // FG % ok
                            System.out.println("Total 3PM: " + total3PM);       // 성공 3점 슛 ok
                            System.out.println("Total 3PA: " + total3PA);       // 시도 3점 슛  ok
                            System.out.println("Total 3P%: " + String.format("%.1f", total3PPct));  // 3P% ok
                            System.out.println("Total FTM: " + totalFTM);       // 성공 자유투 개수  ok
                            System.out.println("Total FTA: " + totalFTA);       // 시도 자유투 개수  ok
                            System.out.println("Total FT%: " + String.format("%.1f", totalFTPct));  // ok

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
                        System.out.println("=> last check >>>>>>>>>>>>>>>>>>>");
                        for (String[] dataArray : gameDataList) {
                            System.out.print("Array: ");
                            for (String data : dataArray) {
                                System.out.print(data + " ");
                            }
                            System.out.println();  // New line for each array
                        }
                        System.out.println("=> last check >>>>>>>>>>>>>>>>>>>");

                        excelList.addAll(gameDataList);

                        driver.navigate().back();

                        try {
                            Thread.sleep(8000);
                        } catch (InterruptedException e) {
                            e.printStackTrace();
                        }

                        System.out.println(" back check 1");
                        // 페이지 네비게이션 후, thirdCardContainer와 liElements를 다시 찾아 업데이트
                        System.out.println("==> secondCardContainer => "+ secondCardContainer);
                        gameCardsContainer = wait.until(ExpectedConditions.visibilityOfElementLocated(By.className("ScheduleLeagueType_match_list_container__1v4b0")));
                        secondCardContainer = gameCardsContainer.findElements(By.className("ScheduleLeagueType_match_list_group__18ML9")).get(i);
                        thirdCardContainer = secondCardContainer.findElement(By.className("ScheduleLeagueType_match_list__1-n6x"));
                        System.out.println(" back check 2");
                        liElements = thirdCardContainer.findElements(By.xpath("./li[contains(@class, 'MatchBox_match_item__3_D0Q') and contains(@class, 'type_end')]"));
                        System.out.println(" back check 3");
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

    @GetMapping("/gameMonthData")
    @ResponseBody
    public ResponseEntity<InputStreamResource> getMonthDownloadGameData(@RequestParam("year5") String year, @RequestParam("day") int lastDay, HttpServletResponse res) throws IOException {
        System.out.println("year : " + year);
        System.out.println("lastDay : " + lastDay);
        String excelName = "nba error";
        HttpHeaders headers = new HttpHeaders();
        String[] day = year.split("-");
        Workbook workbook = null;
//        String fontPath = "C:\\A_ThinkTree-Project\\KMU-Project\\MINA\\chromedriver-win64\\chromedriver.exe";
//        String fontPath = "driver/chromedriver.exe";
//        String driverPath = new File("src/main/resources/driver/chromedriver.exe").getAbsolutePath();

//        System.setProperty("webdriver.chrome.driver", fontPath);
//        String os = System.getProperty("os.name").toLowerCase();
//        String chromDriver = new File("src/main/resources/driver/chromedriver.exe").getAbsolutePath();
        String chromDriver = new File("src/main/resources/driver/chromedriver.exe").getAbsolutePath();
        System.setProperty("webdriver.chrome.driver", chromDriver);
        ChromeOptions options = new ChromeOptions();
//        if (os.contains("win")) {
//            System.out.println("Windows");
//            chromDriver = new File("src/main/resources/driver/chromedriver.exe").getAbsolutePath();
//        } else if (os.contains("mac")) {
//            System.out.println("Mac");
//
//        } else if (os.contains("nix") || os.contains("nux") || os.contains("aix")) {
//            System.out.println("Unix");
//            chromDriver = new File("/home/ubuntu/chromedriver-linux64/chromedriver").getAbsolutePath();
//            options.setBinary("/usr/bin/google-chrome");
//
//        } else if (os.contains("linux")) {
//            System.out.println("OS is linux");
//            chromDriver = new File("/home/ubuntu/chromedriver-linux64/chromedriver").getAbsolutePath();
//            options.setBinary("/usr/bin/google-chrome");
//        }

        options.addArguments("--headless"); // 헤드리스 모드
        options.addArguments("--remote-allow-origins=*");
        options.addArguments("--disable-popup-blocking");       // 팝업안띄움
        options.addArguments("--disable-gpu");                  // gpu 비활성화
        options.addArguments("--disable-images");
        options.addArguments("headless");                       // 브라우저 안띄움
        options.addArguments("--no-sandbox");
        options.addArguments("--blink-settings=imagesEnabled=false"); // 이미지 다운 안받음

        String nbaDate = year + "-" + String.format("%02d", lastDay);

        WebDriver driver = new ChromeDriver(options);
        boolean errorFlagAndStopFlag = false;
        int successDivCheck = 1;
        int divSize = 0;
        int errorCount = 0;
        while (!errorFlagAndStopFlag) {
//            driver.get("https://www.nba.com/games?date=" + year); // ex)1946-11-02
            System.out.println("day :: "+year + "-" + String.format("%02d", lastDay));
            driver.get("https://www.nba.com/games?date="+nbaDate); // ex)1946-11-02

            WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
            workbook = new XSSFWorkbook();

            /** 클래스 값이 NoDataMessage_base__xUA61 인 요소가 있는지 여부 확인 */
            boolean isPresent = isElementPresent(driver, By.className("NoDataMessage_base__xUA61"));
            System.out.println("NoDataMessage_base__xUA61 클래스를 가진 요소가 존재 유무 " + isPresent);
            try {
                if (!isPresent) {
                    try {
                        Thread.sleep(3000);
                    } catch (InterruptedException e) {
                        e.printStackTrace();
                    }
                    WebElement container = wait.until(ExpectedConditions.visibilityOfElementLocated(By.className("GamesView_gameCardsContainer__c_9fB")));
                    /** 해당 요소 바로 아래에 있는 div 요소 가져오기 */
                    List<WebElement> divElements = container.findElements(By.xpath("./div"));

                    /** TEST */
                    // GamesView_gameCardsContainer__c_9fB 클래스로 식별되는 요소의 바로 아래에 있는 div 태그를 모두 선택하는 XPath
                    String xpathExpression = "//div[@class='GamesView_gameCardsContainer__c_9fB']/div";
                    List<WebElement> divElementsTest = driver.findElements(By.xpath(xpathExpression));


                    /** div 요소의 개수 출력 */
                    System.out.println("GAME 개수: " + divElements.size());
                    System.out.println("GAME TEST 개수: " + divElementsTest.size());
                    divSize = divElements.size();

                    /** excel header :: 쿼터수는 변동이 있을 수 있음 */
                    String[] header = {"Year", "Month", "Day", "Attendance", "TEAM", "HOME (홈)", "AWAY (어웨이)", "L(0)/W(1)", "PLAYER", "MIN", "Q1", "Q2", "Q3", "Q4", "OT1", "OT2", "FINAL", "TOT_FGM", "TOT_FGA", "TOT_FG%", "TOT_3PM", "TOT_3PA", "TOT_3P%", "TOT_FTM", "TOT_FTA", "TOT_FT%", "TOT_OREB", "TOT_DREB", "TOT_REB", "TOT_AST", "TOT_STL", "TOT_BLK", "TOT_TO", "TOT_PF", "TOT_PTS", "TOT_+/-", "FGM", "FGA", "FG%", "3PM", "3PA", "3P%", "FTM", "FTA", "FT%", "OREB", "DREB", "REB", "AST", "STL", "BLK", "TO", "PF", "PTS", "+/-"};
                    JavascriptExecutor js = (JavascriptExecutor) driver;

                    Sheet sheet = workbook.createSheet("NBA GAMES");
                    Row excelRow = sheet.createRow(0);
                    for (int i = 0; i < header.length; i++) {
                        Cell cell = excelRow.createCell(i);
                        cell.setCellValue(header[i]);
                    }
                    List<String[]> excelList = new ArrayList<>();
                    System.out.println("<<<< GAME SEARCH >>>>");

                    for (int i = 0; i < divElementsTest.size(); i++) {
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
                                    num = 36;
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
                                    num = 36;
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
                            array[2] = String.format("%02d", lastDay);
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
                            if (isFirstCharacterDigit(gameDataList.get(j)[9])) {
//                    if (!gameDataList.get(j)[9].startsWith("D")) {
                                switch (firstTable.size()) {
                                    case 8:
                                        gameDataList.get(j)[15] = firstTable.get(6); // OT2
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
                                gameDataList.get(j)[16] = firstTable.get(firstTable.size() - 1);    // FINAL
                                gameDataList.get(j)[17] = firstTotalData.get(2);        // TOTO_FGM
                                gameDataList.get(j)[18] = firstTotalData.get(3);        // TOT_EGA
                                gameDataList.get(j)[19] = firstTotalData.get(4);        // TOT_FG%
                                gameDataList.get(j)[20] = firstTotalData.get(5);        //TOT_3PM
                                gameDataList.get(j)[21] = firstTotalData.get(6);        //
                                gameDataList.get(j)[22] = firstTotalData.get(7);
                                gameDataList.get(j)[23] = firstTotalData.get(8);
                                gameDataList.get(j)[24] = firstTotalData.get(9);
                                gameDataList.get(j)[25] = firstTotalData.get(10);
                                gameDataList.get(j)[26] = firstTotalData.get(11);
                                gameDataList.get(j)[27] = firstTotalData.get(12);
                                gameDataList.get(j)[28] = firstTotalData.get(13);
                                gameDataList.get(j)[29] = firstTotalData.get(14);
                                gameDataList.get(j)[30] = firstTotalData.get(15);
                                gameDataList.get(j)[31] = firstTotalData.get(16);
                                gameDataList.get(j)[32] = firstTotalData.get(17);
                                gameDataList.get(j)[33] = firstTotalData.get(18);
                                gameDataList.get(j)[34] = firstTotalData.get(19);
                                gameDataList.get(j)[35] = firstTotalData.get(20);
                            } else {
                                gameDataList.get(j)[9] = "-";
                            }
                        }

                        for (int j = firstRowNumber; j < gameDataList.size(); j++) {
                            gameDataList.get(j)[4] = secondTable.get(0);
                            gameDataList.get(j)[7] = String.valueOf(homeTeamScore);
                            if (isFirstCharacterDigit(gameDataList.get(j)[9])) {
//                    if (!gameDataList.get(j)[9].startsWith("D")) {
                                switch (secondTable.size()) {
                                    case 8:
                                        gameDataList.get(j)[15] = secondTable.get(6); // OT2
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
                                gameDataList.get(j)[16] = secondTable.get(secondTable.size() - 1);      // TOTAL

                                gameDataList.get(j)[17] = secondTotalData.get(2);
                                gameDataList.get(j)[18] = secondTotalData.get(3);
                                gameDataList.get(j)[19] = secondTotalData.get(4);
                                gameDataList.get(j)[20] = secondTotalData.get(5);
                                gameDataList.get(j)[21] = secondTotalData.get(6);
                                gameDataList.get(j)[22] = secondTotalData.get(7);
                                gameDataList.get(j)[23] = secondTotalData.get(8);
                                gameDataList.get(j)[24] = secondTotalData.get(9);
                                gameDataList.get(j)[25] = secondTotalData.get(10);
                                gameDataList.get(j)[26] = secondTotalData.get(11);
                                gameDataList.get(j)[27] = secondTotalData.get(12);
                                gameDataList.get(j)[28] = secondTotalData.get(13);
                                gameDataList.get(j)[29] = secondTotalData.get(14);
                                gameDataList.get(j)[30] = secondTotalData.get(15);
                                gameDataList.get(j)[31] = secondTotalData.get(16);
                                gameDataList.get(j)[32] = secondTotalData.get(17);
                                gameDataList.get(j)[33] = secondTotalData.get(18);
                                gameDataList.get(j)[34] = secondTotalData.get(19);
                                gameDataList.get(j)[35] = secondTotalData.get(20);
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
                        driver.get("https://www.nba.com/games?date=" + nbaDate);
                        excelList.addAll(gameDataList);
                        System.out.println("====> success div size :: " + successDivCheck);
                        successDivCheck++;
                    }   /** div for 문 종료 */
                    System.out.println("final div size :: " + successDivCheck);

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
                errorFlagAndStopFlag = true;
            } catch (NoSuchElementException e) {
                System.out.println("NoSuchElementException : " + e.getMessage());
                errorCount++;
                if (errorCount == 3) {
                    errorFlagAndStopFlag = true;
                    return null;
                }
            } catch (Exception e) {
                System.out.println("Exception : " + e.getMessage());
                errorCount++;
                if (errorCount == 3) {
                    errorFlagAndStopFlag = true;
                    return null;
                }
            }
        }

        excelName = year + "-" + String.format("%02d", lastDay);
        if(divSize == 1) {
            excelName = year + "-" + String.format("%02d", lastDay) +" [1]";
        }

        System.out.println(" !!! excel insert end & download start !!! ");
        headers.add("Content-Disposition", "attachment; filename="+excelName+".xlsx");

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

    @PostMapping("/gameData")
    @ResponseBody
    public String getDownloadGameData(@RequestParam("year4") String year, HttpServletResponse res) {
        System.out.println("year : " + year);
        String[] day = year.split("-");
//        String fontPath = "C:\\A_ThinkTree-Project\\KMU-Project\\MINA\\chromedriver-win64\\chromedriver.exe";
//        String chromDriver = new File("/home/ubuntu/chromedriver-linux64/chromedriver").getAbsolutePath();    //서버
        String chromDriver = new File("src/main/resources/driver/chromedriver.exe").getAbsolutePath();

        System.setProperty("webdriver.chrome.driver", chromDriver);
        ChromeOptions options = new ChromeOptions();
        String os = System.getProperty("os.name").toLowerCase();
        if (os.contains("nix") || os.contains("nux") || os.contains("aix")) {
            System.out.println("Unix");
            options.setBinary("/usr/bin/google-chrome");  //서버
        } else if (os.contains("linux")) {
            System.out.println("OS is linux");
            options.setBinary("/usr/bin/google-chrome");
        }
        options.addArguments("--headless"); // 헤드리스 모드
        options.addArguments("--remote-allow-origins=*");
        options.addArguments("--disable-popup-blocking");       // 팝업안띄움
        options.addArguments("--disable-gpu");                  // gpu 비활성화
        options.addArguments("--disable-images");
        options.addArguments("headless");                       // 브라우저 안띄움
        options.addArguments("--no-sandbox");
        options.addArguments("--blink-settings=imagesEnabled=false"); // 이미지 다운 안받음

        WebDriver driver = new ChromeDriver(options);
        boolean errorFlagAndStopFlag = false;
        int successDivCheck = 1;
        while (!errorFlagAndStopFlag) {
            driver.get("https://www.nba.com/games?date=" + year); // ex)1946-11-02

            WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
            Workbook workbook = new XSSFWorkbook();

            /** 클래스 값이 NoDataMessage_base__xUA61 인 요소가 있는지 여부 확인 */
            boolean isPresent = isElementPresent(driver, By.className("NoDataMessage_base__xUA61"));
            System.out.println("NoDataMessage_base__xUA61 클래스를 가진 요소가 존재 유무 " + isPresent);
            try {
                if (!isPresent) {
                    try {
                        Thread.sleep(3000);
                    } catch (InterruptedException e) {
                        e.printStackTrace();
                    }
                    WebElement container = wait.until(ExpectedConditions.visibilityOfElementLocated(By.className("GamesView_gameCardsContainer__c_9fB")));
                    /** 해당 요소 바로 아래에 있는 div 요소 가져오기 */
                    List<WebElement> divElements = container.findElements(By.xpath("./div"));

                    /** TEST */
                    // GamesView_gameCardsContainer__c_9fB 클래스로 식별되는 요소의 바로 아래에 있는 div 태그를 모두 선택하는 XPath
                    String xpathExpression = "//div[@class='GamesView_gameCardsContainer__c_9fB']/div";
                    List<WebElement> divElementsTest = driver.findElements(By.xpath(xpathExpression));


                    /** div 요소의 개수 출력 */
                    System.out.println("GAME 개수: " + divElements.size());
                    System.out.println("GAME TEST 개수: " + divElementsTest.size());

                    /** excel header :: 쿼터수는 변동이 있을 수 있음 */
                    String[] header = {"Year", "Month", "Day", "Attendance", "TEAM", "HOME (홈)", "AWAY (어웨이)", "L(0)/W(1)", "PLAYER", "MIN", "Q1", "Q2", "Q3", "Q4", "OT1", "OT2", "FINAL", "TOT_FGM", "TOT_FGA", "TOT_FG%", "TOT_3PM", "TOT_3PA", "TOT_3P%", "TOT_FTM", "TOT_FTA", "TOT_FT%", "TOT_OREB", "TOT_DREB", "TOT_REB", "TOT_AST", "TOT_STL", "TOT_BLK", "TOT_TO", "TOT_PF", "TOT_PTS", "TOT_+/-", "FGM", "FGA", "FG%", "3PM", "3PA", "3P%", "FTM", "FTA", "FT%", "OREB", "DREB", "REB", "AST", "STL", "BLK", "TO", "PF", "PTS", "+/-"};
                    JavascriptExecutor js = (JavascriptExecutor) driver;

                    Sheet sheet = workbook.createSheet("NBA GAMES");
                    Row excelRow = sheet.createRow(0);
                    for (int i = 0; i < header.length; i++) {
                        Cell cell = excelRow.createCell(i);
                        cell.setCellValue(header[i]);
                    }
                    List<String[]> excelList = new ArrayList<>();
                    System.out.println("<<<< GAME SEARCH >>>>");

                    for (int i = 0; i < divElementsTest.size(); i++) {
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
                                    num = 36;
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
                                    num = 36;
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
                            if (isFirstCharacterDigit(gameDataList.get(j)[9])) {
//                    if (!gameDataList.get(j)[9].startsWith("D")) {
                                switch (firstTable.size()) {
                                    case 8:
                                        gameDataList.get(j)[15] = firstTable.get(6); // OT1
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
                                gameDataList.get(j)[16] = firstTable.get(firstTable.size() - 1);    // FINAL
                                gameDataList.get(j)[17] = firstTotalData.get(2);        // TOTO_FGM
                                gameDataList.get(j)[18] = firstTotalData.get(3);        // TOT_EGA
                                gameDataList.get(j)[19] = firstTotalData.get(4);        // TOT_FG%
                                gameDataList.get(j)[20] = firstTotalData.get(5);        //TOT_3PM
                                gameDataList.get(j)[21] = firstTotalData.get(6);        //
                                gameDataList.get(j)[22] = firstTotalData.get(7);
                                gameDataList.get(j)[23] = firstTotalData.get(8);
                                gameDataList.get(j)[24] = firstTotalData.get(9);
                                gameDataList.get(j)[25] = firstTotalData.get(10);
                                gameDataList.get(j)[26] = firstTotalData.get(11);
                                gameDataList.get(j)[27] = firstTotalData.get(12);
                                gameDataList.get(j)[28] = firstTotalData.get(13);
                                gameDataList.get(j)[29] = firstTotalData.get(14);
                                gameDataList.get(j)[30] = firstTotalData.get(15);
                                gameDataList.get(j)[31] = firstTotalData.get(16);
                                gameDataList.get(j)[32] = firstTotalData.get(17);
                                gameDataList.get(j)[33] = firstTotalData.get(18);
                                gameDataList.get(j)[34] = firstTotalData.get(19);
                                gameDataList.get(j)[35] = firstTotalData.get(20);
                            } else {
                                gameDataList.get(j)[9] = "-";
                            }
                        }

                        for (int j = firstRowNumber; j < gameDataList.size(); j++) {
                            gameDataList.get(j)[4] = secondTable.get(0);
                            gameDataList.get(j)[7] = String.valueOf(homeTeamScore);
                            if (isFirstCharacterDigit(gameDataList.get(j)[9])) {
//                    if (!gameDataList.get(j)[9].startsWith("D")) {
                                switch (secondTable.size()) {
                                    case 8:
                                        gameDataList.get(j)[15] = secondTable.get(6); // OT2
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
                                gameDataList.get(j)[16] = secondTable.get(secondTable.size() - 1);      // TOTAL

                                gameDataList.get(j)[17] = secondTotalData.get(2);
                                gameDataList.get(j)[18] = secondTotalData.get(3);
                                gameDataList.get(j)[19] = secondTotalData.get(4);
                                gameDataList.get(j)[20] = secondTotalData.get(5);
                                gameDataList.get(j)[21] = secondTotalData.get(6);
                                gameDataList.get(j)[22] = secondTotalData.get(7);
                                gameDataList.get(j)[23] = secondTotalData.get(8);
                                gameDataList.get(j)[24] = secondTotalData.get(9);
                                gameDataList.get(j)[25] = secondTotalData.get(10);
                                gameDataList.get(j)[26] = secondTotalData.get(11);
                                gameDataList.get(j)[27] = secondTotalData.get(12);
                                gameDataList.get(j)[28] = secondTotalData.get(13);
                                gameDataList.get(j)[29] = secondTotalData.get(14);
                                gameDataList.get(j)[30] = secondTotalData.get(15);
                                gameDataList.get(j)[31] = secondTotalData.get(16);
                                gameDataList.get(j)[32] = secondTotalData.get(17);
                                gameDataList.get(j)[33] = secondTotalData.get(18);
                                gameDataList.get(j)[34] = secondTotalData.get(19);
                                gameDataList.get(j)[35] = secondTotalData.get(20);
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
                        System.out.println("====> success div size :: " + successDivCheck);
                        successDivCheck++;
                    }   /** div for 문 종료 */
                    System.out.println("final div size :: " + successDivCheck);

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

                errorFlagAndStopFlag = true;
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
            } catch (NoSuchElementException e) {
                System.out.println("NoSuchElementException : " + e.getMessage());
            } catch (Exception e) {
                System.out.println("Exception : " + e.getMessage());
            }
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

//        String fontPath = "C:\\A_ThinkTree-Project\\KMU-Project\\MINA\\chromedriver-win64\\chromedriver.exe";

        System.setProperty("webdriver.chrome.driver", driverPath);
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

    @PostMapping("/team")
    @ResponseBody
    public String getExcelDownTeam(@RequestParam("team") String year, HttpServletResponse res) {
        System.out.println("== team : ");

//        String fontPath = "C:\\A_ThinkTree-Project\\KMU-Project\\MINA\\chromedriver-win64\\chromedriver.exe";

        System.setProperty("webdriver.chrome.driver", driverPath);

        ChromeOptions options = new ChromeOptions();
        options.addArguments("--remote-allow-origins=*");

//        options.addArguments("--disable-popup-blocking");       // 팝업안띄움
//        options.addArguments("--disable-gpu");                  // gpu 비활성화
//        options.addArguments("--disable-images");
//        options.addArguments("headless");                       // 브라우저 안띄움
//        options.addArguments("--no-sandbox");
//        options.addArguments("--blink-settings=imagesEnabled=false"); // 이미지 다운 안받음

        WebDriver driver = new ChromeDriver(options);
//        driver.get("https://www.nba.com/stats/players/boxscores-traditional?Season=" + year);
        driver.get("https://www.nba.com/stats/teams/traditional?SeasonType=Regular+Season&PerMode=Totals&Season=" + year);

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("NBA DATA");

//        final String[] header = {"PLAYER", "TEAM", "MATCH UP", "GAME DATE", "W/L", "MIN", "PTS", "FGM", "FGA", "FG%", "3PM", "3PA", "3P%", "FTM", "FTA", "FT%", "OREB", "DREB", "REB", "AST", "STL", "BLK", "TOV", "PF", "+/-"};
        final String[] header = {"No", "TEAM", "GP", "W", "L", "WIN%", "MIN", "PTS", "FGM", "FGA", "FG%", "3PM", "3PA", "3P%", "FTM", "FTA", "FT%", "OREB", "DREB", "REB", "AST", "TOV", "STL", "BLK", "BLKA", "PF", "PFD", "+/-"};
        Row row = sheet.createRow(0);
        for (int i = 0; i < header.length; i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue(header[i]);
        }

        try {
            Thread.sleep(3000);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }

        System.out.println("====> header :: ");
//        WebElement tbody = driver.findElement(By.xpath("//tbody[contains(@class, 'Crom_body__UYOcU')]"));
        WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
        WebElement tbody = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//tbody[contains(@class, 'Crom_body__UYOcU')]")));
        List<WebElement> rows = tbody.findElements(By.tagName("tr"));

        System.out.println("====> rows  :: ");

        try {
            List<List<String>> dataList = new ArrayList<>();
            for (WebElement webRow : rows) {
                List<WebElement> columns = webRow.findElements(By.tagName("td"));
                List<String> rowData = new ArrayList<>();
                for (WebElement column : columns) {
                    rowData.add(column.getText());
                }
                dataList.add(rowData);
            }

            System.out.println("excel insert : ");
            int rowNum = 1;
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

    @PostMapping("/mina2")
    @ResponseBody
    public String getExcelDownLocation(@RequestParam("year2") String year, HttpServletResponse res) {
        System.out.println("==> test 2");

//        String fontPath = "C:\\A_ThinkTree-Project\\KMU-Project\\MINA\\chromedriver-win64\\chromedriver.exe";

        System.setProperty("webdriver.chrome.driver", driverPath);

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

//        String fontPath = "C:\\A_ThinkTree-Project\\KMU-Project\\MINA\\chromedriver-win64\\chromedriver.exe";

        System.setProperty("webdriver.chrome.driver", driverPath);

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
//        String fontPath = "C:\\A_ThinkTree-Project\\KMU-Project\\MINA\\chromedriver-win64\\chromedriver.exe";
        String fontPath = "";

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


