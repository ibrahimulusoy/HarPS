import com.sun.corba.se.spi.ior.Writeable;
import com.sun.net.httpserver.Authenticator;
import com.sun.org.apache.xpath.internal.operations.Bool;
import com.sun.org.apache.xpath.internal.operations.Equals;
import jxl.Sheet;
import jxl.Workbook;
import jxl.format.CellFormat;
import jxl.read.biff.BiffException;
import java.io.File;
import jxl.write.*;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedCondition;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.IRetryAnalyzer;
import org.testng.ITestResult;
import org.testng.annotations.Test;
import org.openqa.selenium.support.ui.ExpectedConditions;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;



public class CalenderDays implements IRetryAnalyzer {
    //this one is the last version

    int counter = 0;
    int retryLimit = 20;
    @Test(retryAnalyzer = Authenticator.Retry.class)
    public void test() throws InterruptedException, IOException, BiffException, ParseException, WriteException, InvalidFormatException {
        File file=new File("C:\\Users\\iulusoy\\Desktop\\CAMPUSES\\011.xls");
        FileInputStream fs =new FileInputStream(file);
        org.apache.poi.ss.usermodel.Workbook wb =new HSSFWorkbook(fs);
        org.apache.poi.ss.usermodel.Sheet sh ;

       // String FilePath = "C:\\Users\\iulusoy\\Desktop\\CAMPUSES\\060.xls";
       // FileInputStream fs = new FileInputStream(FilePath);
       // Workbook wb = Workbook.getWorkbook(fs);
       // Sheet sh;
       // WritableWorkbook wbCopy=Workbook.createWorkbook(new File("C:\\Users\\iulusoy\\Desktop\\CAMPUSES\\001_.xls"),wb);
       // WritableSheet wSh;


        System.setProperty("webdriver.chrome.driver","C:\\Program Files\\Driver\\chromedriver.exe");
       // ChromeOptions options =new ChromeOptions();
       // options.addArguments("headless");
       // options.addArguments("window-size=1920x1080");
        WebDriver driver =new ChromeDriver();
        driver.get("https://skyward.iscorp.com/HarmonyTXStuSTS");
        //login
        driver.manage().window().maximize();
        driver.findElement(By.id("UserName")).sendKeys("Automation.1");
        driver.findElement(By.id("Password")).sendKeys("Auto1.!@#");
        driver.findElement(By.cssSelector("button[type='submit']")).click();
        //Choose the district
        WebDriverWait dr=new WebDriverWait(driver,20);
        // Step 6: Locate the search area, and wait till it appears.
        dr.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id='autoid1']/a/span")));
        Thread.sleep(1000);
        driver.findElement(By.xpath("//*[@id='autoid1']/a/span")).click();
        dr.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id='_code']")));
        Thread.sleep(1000);
        driver.findElement(By.xpath("//*[@id='_code']")).sendKeys("011");
        Thread.sleep(200);
        driver.findElement(By.xpath("//*[@id='_code']")).sendKeys(Keys.ARROW_DOWN);
        Thread.sleep(200);
        driver.findElement(By.xpath("//*[@id='_code']")).sendKeys(Keys.ENTER);
        //navigate to the calender
        Thread.sleep(1000);
        driver.get("https://skyward.iscorp.com/HarmonyTXStu/Attendance/Calendar/List?w=4c4dfbbc70d94999a090e57d689e907f&p=a2d6698bb58c44e1ad0a6ec213c2238b");
        //

        Thread.sleep(1000);
        //driver.findElement(By.linkText("Austin-HSA-K-8 Main")).click();
        //Select s = new Select(driver.findElement(By.id("//*[@id=\"CalendarCalendarDays_footer\"]/div[1]/select")));
        //s.selectByValue("200");
        //iver.findElement(By.xpath("//*[@id='CalendarCalendarDays_footer']/div[1]/select"));
        int countCalender= driver.findElements(By.xpath("//*[contains(@id,'browse_lockedRow')]")).size();
        Thread.sleep(1000);
        int CountDays ;   //= driver.findElements(By.xpath("//*[contains(@id,'CalendarCalendarDays_lockedRow')]")).size();
        Thread.sleep(1000);
       // System.out.println("sayi:" +countCalender);
       // System.out.println("sayi:" +CountDays);
        String lastXpathCalender;
        String lastXpathDays;
        String lastDate=null;
        for (int i=1;i<countCalender;i++){
            lastXpathCalender= "//*[@id='"+"browse_lockedRow"+i+"']";
            Thread.sleep(1000);
            dr.until(ExpectedConditions.presenceOfElementLocated(By.xpath(lastXpathCalender)));
            dr.until(ExpectedConditions.elementToBeClickable(By.xpath(lastXpathCalender)));
            String calendarCode= driver.findElement(By.xpath("//*[@id='"+"browse_row"+i+"_col0']/div")).getText();
            //*[@id="browse_row0_col0"]/div
            sh = wb.getSheet(calendarCode);
            int totalRows= sh.getLastRowNum();

            CountDays=totalRows-7;
            System.out.println("CountDays:"+CountDays);
            String remainingStartPoint;
            int excelPoint=-1;
            int k=0;
            Thread.sleep(1000);
//          dr.until(ExpectedConditions.visibilityOfAllElements());
            System.out.println("1");
            for(k =7;k<totalRows;k++){
              //  System.out.println(totalRows);
               // System.out.println(sh.getCell(6,k).getContents());

                Row row =sh.getRow(k);

                Cell cell6 = row.getCell(6);

                Cell cell1 = row.getCell(1);

                if(!cell6.getStringCellValue().trim().equals("Done")){

                    DateFormat format = new SimpleDateFormat("dd/MM/yyyy");
                    Date date=cell1.getDateCellValue();

              //      System.out.println(">> "+date);
                    lastDate= new SimpleDateFormat("MM/dd/yyyy").format(date);
                //    System.out.println("ana"+lastDate);
                    excelPoint=k;
                  //  remainingStartPoint =cell1.getStringCellValue();
                //    System.out.println(remainingStartPoint);
                    break;
                }
            }
          System.out.println("4");
            dr.until(ExpectedConditions.elementToBeClickable(By.xpath(lastXpathCalender)));
            driver.findElement(By.xpath(lastXpathCalender)).click();
            dr.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id='browse_previewDetails']")));
            driver.findElement(By.xpath("//*[@id='browse_previewDetails']")).click();
            int j=0;
            for( j =k-7;j<CountDays;j++){
                Thread.sleep(1000);
                lastXpathDays= "//*[@id='"+"CalendarCalendarDays_lockedRow"+j+"']";
                Thread.sleep(1000);
                dr.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//*[@id='CalendarCalendarDays_row"+j+"_col0']/div")));
                dr.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id='CalendarCalendarDays_row"+j+"_col0']/div")));
                dr.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id='CalendarCalendarDays_row"+j+"_col0']/div")));
                String cDate= driver.findElement(By.xpath("//*[@id='CalendarCalendarDays_row"+j+"_col0']/div")).getText().trim();
               // System.out.println("--");
                  System.out.println("cdate"+cDate+" / lastDate"+lastDate);
               // System.out.println("lastDate"+lastDate);
               // System.out.println("++");
                if (lastDate.trim().equals(cDate.trim())) {
                    //    System.out.println("girdi"+cDate);

                    if (!cDate.equals("05/22/2020")) {
                        int next = j + 1;
                        lastDate = driver.findElement(By.xpath("//*[@id='CalendarCalendarDays_row" + next + "_col0']/div")).getText().trim();
                    }
                    int exCount = 7 + j;
                    dr.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(lastXpathDays)));
                    dr.until(ExpectedConditions.elementToBeClickable(By.xpath(lastXpathDays)));
                    driver.findElement(By.xpath(lastXpathDays)).click();
                    Thread.sleep(500);

                        //System.out.println(sh.getCell(4,exCount).getContents());
                        dr.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id='editCalendarDay_CountAs']")));
                        driver.findElement(By.xpath("//*[@id='editCalendarDay_CountAs']")).clear();
                        Row rexCount=sh.getRow(exCount);
                        //System.out.println("5");
                        driver.findElement(By.xpath("//*[@id='editCalendarDay_CountAs']")).sendKeys(String.valueOf(rexCount.getCell(4).getNumericCellValue()));
                        /*
                        if (sh.getCell(2,exCount).getContents().equals("M")) {
                            driver.findElement(By.xpath("//*[@id='editCalendarDay_DayRotationID_code']")).clear();
                            driver.findElement(By.xpath("//*[@id='editCalendarDay_DayRotationID_code']")).sendKeys(sh.getCell(2, exCount).getContents());
                            dr.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id='popup0_browse_lockedRow1']/td/div/div/span")));
                            dr.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//*[@id='popup0_browse_lockedRow1']/td/div/div/span")));
                            driver.findElement(By.xpath("//*[@id='popup0_browse_lockedRow1']/td/div/div/span")).click();
                        }
                        else if (sh.getCell(2,exCount).getContents().equals("T")) {
                            driver.findElement(By.xpath("//*[@id='editCalendarDay_DayRotationID_code']")).clear();
                            driver.findElement(By.xpath("//*[@id='editCalendarDay_DayRotationID_code']")).sendKeys(sh.getCell(2, exCount).getContents());
                            dr.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id='popup0_browse_lockedRow2']/td/div/div/span")));
                            dr.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//*[@id='popup0_browse_lockedRow2']/td/div/div/span")));
                            driver.findElement(By.xpath("//*[@id='popup0_browse_lockedRow2']/td/div/div/span")).click();
                        }
                        else if (sh.getCell(2,exCount).getContents().equals("W")) {
                            driver.findElement(By.xpath("//*[@id='editCalendarDay_DayRotationID_code']")).clear();
                            driver.findElement(By.xpath("//*[@id='editCalendarDay_DayRotationID_code']")).sendKeys(sh.getCell(2, exCount).getContents());
                            dr.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id='popup0_browse_lockedRow1']/td/div/div/span")));
                            dr.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//*[@id='popup0_browse_lockedRow1']/td/div/div/span")));
                            driver.findElement(By.xpath("//*[@id='popup0_browse_lockedRow1']/td/div/div/span")).click();
                        }
                        else if (sh.getCell(2,exCount).getContents().equals("R")) {
                            driver.findElement(By.xpath("//*[@id='editCalendarDay_DayRotationID_code']")).clear();
                            driver.findElement(By.xpath("//*[@id='editCalendarDay_DayRotationID_code']")).sendKeys(sh.getCell(2, exCount).getContents());
                            dr.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id='_lockedMirror']/span")));
                            dr.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//*[@id='_lockedMirror']/span")));
                            driver.findElement(By.xpath("//*[@id='_lockedMirror']/span")).click();
                        }
                        else if (sh.getCell(2,exCount).getContents().equals("F")) {
                            driver.findElement(By.xpath("//*[@id='editCalendarDay_DayRotationID_code']")).clear();
                            driver.findElement(By.xpath("//*[@id='editCalendarDay_DayRotationID_code']")).sendKeys(sh.getCell(2, exCount).getContents());
                            dr.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id='popup0_browse_lockedRow2']/td/div/div/span")));
                            dr.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//*[@id='popup0_browse_lockedRow2']/td/div/div/span")));
                            driver.findElement(By.xpath("//*[@id='popup0_browse_lockedRow2']/td/div/div/span")).click();
                        }

                         */

                    try{
                        driver.findElement(By.xpath("//*[@id='editCalendarDay_DayRotationID_code']")).clear();
                        Row rw=sh.getRow(exCount);
                        driver.findElement(By.xpath("//*[@id='editCalendarDay_DayRotationID_code']")).sendKeys(String.valueOf(rexCount.getCell(2).getStringCellValue()));
                        Thread.sleep(500);
                        dr.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id='_lockedMirror']/span")));
                        dr.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//*[@id='_lockedMirror']/span")));
                        driver.findElement(By.xpath("//*[@id='_lockedMirror']/span")).click();
                    }
                    catch (org.openqa.selenium.StaleElementReferenceException e)
                    {
                        Thread.sleep(2000);
                        driver.findElement(By.xpath("//*[@id='editCalendarDay_DayRotationID_code']")).clear();
                        driver.findElement(By.xpath("//*[@id='editCalendarDay_DayRotationID_code']")).sendKeys(String.valueOf(rexCount.getCell(2).getStringCellValue()));
                        Thread.sleep(500);
                        dr.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id='_lockedMirror']/span")));
                        dr.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//*[@id='_lockedMirror']/span")));
                        driver.findElement(By.xpath("//*[@id='_lockedMirror']/span")).click();
                    }


                   // Thread.sleep(300);
                   // dr.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//*[@id='_lockedMirror']/span")));
                     //   dr.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id='_lockedMirror']/span")));

                      //  d
                         //   }
                        //catch(org.openqa.selenium.StaleElementReferenceException ex)
                       // {
                            //driver.findElement(By.xpath("//*[@id='editCalendarDay_DayRotationID_code']")).clear();
                            //driver.findElement(By.xpath("//*[@id='editCalendarDay_DayRotationID_code']")).sendKeys(sh.getCell(2,exCount).getContents());
                            // dr.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//*[@id='_lockedMirror']/span")));
                           // dr.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//*[@id='_lockedMirror']/span")));
                           // dr.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id='_lockedMirror']/span")));
                        //    driver.findElement(By.xpath("//*[@id='_lockedMirror']/span")).click();
                        //}
                        driver.findElement(By.xpath("//*[@id='editCalendarDay_Comment']")).sendKeys(".");
                        driver.findElement(By.xpath("//*[@id='editCalendarDay_AttendancePeriodIDFundingPeriod_code']")).clear();
                        driver.findElement(By.xpath("//*[@id='editCalendarDay_AttendancePeriodIDFundingPeriod_code']")).sendKeys("2");
                    //*[@id="_lockedMirror"]/span

                        dr.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id='_lockedMirror']/span")));
                        driver.findElement(By.xpath("//*[@id='_lockedMirror']/span")).click();
                        dr.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id='editCalendarDay_save']/span")));
                        driver.findElement(By.xpath("//*[@id='editCalendarDay_save']/span")).click();
                        dr.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@id='editCalendarDay_save']/span")));
                        Thread.sleep(1000);
                        dr.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(lastXpathCalender)));
                        dr.until(ExpectedConditions.elementToBeClickable(By.xpath(lastXpathCalender)));
                        //Thread.sleep(500);

                         //bell schedule seciliyor

/*
                        try
                        {
                            driver.findElement(By.xpath("//*[@id='modalLayerDialog0']/div[3]/button[2]/div/span")).click();
                        }
                        catch (NoSuchElementException e)
                        {

                        }

                    Thread.sleep(1000);
                    dr.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id='tabCalendarDayBellScheduleGroupBellSchedules']")));
                    driver.findElement(By.xpath("//*[@id='tabCalendarDayBellScheduleGroupBellSchedules']")).click();
                    Thread.sleep(1000);

                    dr.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//*[@id='detailsPanel_CalendarDayBellScheduleGroupBellSchedules_newButton']/span[1]")));
                    dr.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id='detailsPanel_CalendarDayBellScheduleGroupBellSchedules_newButton']/span[1]")));
                    driver.findElement(By.xpath("//*[@id='detailsPanel_CalendarDayBellScheduleGroupBellSchedules_newButton']/span[1]")).click();
                    Thread.sleep(1000);
                    //Schedule 1 //*[@id="SelectedBellSchedule-715"]/div/a
                    dr.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//*[@id='SelectedBellSchedule-723_code']")));
                    dr.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id='SelectedBellSchedule-723_code']")));
                    driver.findElement(By.xpath("//*[@id='SelectedBellSchedule-723']/div/a")).click();
                    dr.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id='_lockedMirror']/span")));
                    driver.findElement(By.xpath("//*[@id='_lockedMirror']/span")).click();

                    dr.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//*[@id='SelectedBellSchedule-724_code']")));
                    dr.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id='SelectedBellSchedule-724_code']")));
                    driver.findElement(By.xpath("//*[@id='SelectedBellSchedule-724']/div/a")).click();
                    dr.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id='_lockedMirror']/span")));
                    driver.findElement(By.xpath("//*[@id='_lockedMirror']/span")).click();

                    //dr.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//*[@id='SelectedBellSchedule-721_code']")));
                    // dr.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id='SelectedBellSchedule-721_code']")));
                    // driver.findElement(By.xpath("//*[@id='SelectedBellSchedule-721']/div/a")).click();
                    // dr.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id='_lockedMirror']/span")));
                    // driver.findElement(By.xpath("//*[@id='_lockedMirror']/span")).click();

                        /*Schedule 2 //*[@id="SelectedBellSchedule-715"]/div/a
                        dr.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//*[@id='SelectedBellSchedule-721_code']")));
                        dr.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id='SelectedBellSchedule-721_code']")));
                        driver.findElement(By.xpath("//*[@id='SelectedBellSchedule-721']/div/a")).click();
                        dr.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id='popup1_browse_lockedRow1']/td/div/div/span")));
                        driver.findElement(By.xpath("//*[@id='popup1_browse_lockedRow1']/td/div/div/span")).click();

                    if (driver.findElement(By.xpath("//*[@id='SelectedBellSchedule-667_code']")).getAttribute("data-original-value").equals("A"))
                    {
                        driver.findElement(By.xpath("//*[@id='SelectedBellSchedule-667_code']")).click();
                        driver.findElement(By.xpath("//*[@id='SelectedBellSchedule-667_code']")).sendKeys(Keys.BACK_SPACE);
                        //driver.findElement(By.xpath("//*[@id='SelectedBellSchedule-643_code']")).sendKeys(Keys.BACK_SPACE);
                        //driver.findElement(By.xpath("//*[@id='SelectedBellSchedule-643_code']")).sendKeys(Keys.DELETE);
                        Thread.sleep(1000);
                        dr.until(ExpectedConditions.visibilityOfElementLocated(By.id("modalLayerGlass1")));
                        Actions act = new Actions(driver);
                        act.moveByOffset(1, 1).click().build().perform();
                        act.moveByOffset(1, 1).click().build().perform();
                        act.moveByOffset(1, 1).click().build().perform();
                        dr.until(ExpectedConditions.invisibilityOfElementLocated(By.id("modalLayerGlass1")));
                    }


                        dr.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id='popup0_next']/span")));
                        driver.findElement(By.xpath("//*[@id='popup0_next']/span")).click();

                        boolean invisible= dr.until(ExpectedConditions.invisibilityOfElementLocated(By.xpath("//*[@id='popup0_next']/span")));

                        if (invisible)
                        {
                            Row rw=sh.getRow(j+7);
                            Cell cell = rw.createCell(6);
                            cell.setCellValue("Done");
                            fs.close();
                            //System.out.println("6" );
                            FileOutputStream fos=new FileOutputStream(file);
                            //System.out.println("7" );
                            wb.write(fos);
                            fos.close();
                            // wb.close();

                            dr.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//*[@id='tabGeneral']")));
                            dr.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id='tabGeneral']")));
                            Thread.sleep(1000);
                            driver.findElement(By.xpath("//*[@id='tabGeneral']")).click();
                        }
*/
                    Row rw=sh.getRow(j+7);
                    Cell cell = rw.createCell(6);
                    cell.setCellValue("Done");
                    fs.close();
                    //System.out.println("6" );
                    FileOutputStream fos=new FileOutputStream(file);
                    //System.out.println("7" );
                    wb.write(fos);
                    fos.close();
                    System.out.println("CalenderCode /"+cDate.trim() );
                    System.out.println("--" );
                }
            }
        }
    }

    public boolean retry(ITestResult result) {

        if(counter < retryLimit)
        {
            counter++;
            return true;
        }
        return false;

    }
    public void call(){



    }

}
