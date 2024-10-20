package STTCA.STTCA;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Action;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;

import io.github.bonigarcia.wdm.WebDriverManager;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import java.util.Date;
import java.util.concurrent.TimeUnit;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class TestingJava {

    WebDriver driver;
    XSSFWorkbook workbook;
    XSSFSheet sheet;
    
    
    @BeforeTest
    public void setup() throws IOException, InterruptedException, InvalidFormatException {
    	
        WebDriverManager.chromedriver().setup();
        driver = new ChromeDriver();
        driver.get("https://demo.wpeverest.com/user-registration/customer-account-opening-form/");
        driver.manage().window().maximize();

        // Load Excel file
        File scrfile = new File("C:\\Users\\dell\\eclipse-workspace\\STTCA\\target\\Stt.xlsx");
       
        workbook = new XSSFWorkbook(scrfile);
        sheet = workbook.getSheetAt(0); // Access the first sheet   
    }

    @Test
    public void formSubmissionTest() throws IOException, InterruptedException {
    	HashMap<String, String> monthMap = new HashMap<>();

        // Add key-value pairs for months
        monthMap.put("01", "January");
        monthMap.put("02", "February");
        monthMap.put("03", "March");
        monthMap.put("04", "April");
        monthMap.put("05", "May");
        monthMap.put("06", "June");
        monthMap.put("07", "July");
        monthMap.put("08", "August");
        monthMap.put("09", "September");
        monthMap.put("10", "October");
        monthMap.put("11", "November");
        monthMap.put("12", "December");
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            XSSFRow row = sheet.getRow(i);

            // Fetch data from each column
            String salutation = row.getCell(1).getStringCellValue();
            String firstName = row.getCell(2).getStringCellValue();
            String lastName = row.getCell(3).getStringCellValue();
            String userEmail = row.getCell(4).getStringCellValue();
            String userPassword = row.getCell(5).getStringCellValue();
            //String mobileNumber = String.valueOf((long) row.getCell(6).getNumericCellValue());
            Date dobDate = row.getCell(7).getDateCellValue();
            DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
            String dob = dateFormat.format(dobDate);
            
            String[] dateParts = dob.split("-");
            
            String dobYear = dateParts[0];
            String dobMonth = dateParts[1];
            String dobDay = dateParts[2];

            String nationality = row.getCell(8).getStringCellValue();
            String panNo = row.getCell(9).getStringCellValue();
            String citizenshipNo = String.valueOf((long) row.getCell(10).getNumericCellValue());;
            Date issueDate = row.getCell(11).getDateCellValue();
            String issue = dateFormat.format(issueDate);
            String[] issueDateParts = issue.split("-");
            
            String issueYear = issueDateParts[0];
            String issueMonth = issueDateParts[1];
            String issueDay = issueDateParts[2];
            
            String accountType = row.getCell(12).getStringCellValue();
            String grandFathersName = row.getCell(13).getStringCellValue();
            String fathersName = row.getCell(14).getStringCellValue();
            String mothersName = row.getCell(15).getStringCellValue();
            String spouseName = row.getCell(16).getStringCellValue();
            String houseNumber = row.getCell(17).getStringCellValue();
            String wardNumber = row.getCell(18).getStringCellValue();
            String streetName = row.getCell(19).getStringCellValue();
            
            String district = row.getCell(20).getStringCellValue();
            String vdcRuralMunicipality = row.getCell(21).getStringCellValue();
            String location = row.getCell(22).getStringCellValue();
            String workStatus = row.getCell(23).getStringCellValue();
            String purposeOfAccount = row.getCell(24).getStringCellValue();
            String sourceOfIncome = row.getCell(25).getStringCellValue();
            Double annualSalaryIncome = row.getCell(26).getNumericCellValue();
            String salary = String.valueOf(annualSalaryIncome.intValue());
            int electronicBanking = (int) row.getCell(27).getNumericCellValue();
            int electronicStatement = (int) row.getCell(28).getNumericCellValue();
            int smsAlertService = (int) row.getCell(29).getNumericCellValue();
            String citizenshipFront = row.getCell(30).getStringCellValue();

            /*
            */
            // Fill in the form on the web page using the field names (adjust locators as necessary)
            WebElement salutationField = driver.findElement(By.xpath("//*[@id=\"select_1665559734\"]"));
            WebElement firstNameField = driver.findElement(By.xpath("//*[@id=\"first_name\"]"));
            WebElement lastNameField = driver.findElement(By.xpath("//*[@id=\"last_name\"]"));
            WebElement emailField = driver.findElement(By.xpath("//*[@id=\"user_email\"]"));
            WebElement passwordField = driver.findElement(By.xpath("//*[@id=\"user_pass\"]"));
            //WebElement mobileNumberField = driver.findElement(By.xpath("//*[@id=\"phone_1665559881\"]"));
            WebElement dobField = driver.findElement(By.xpath("//*[@id=\"date_box_1665559930_field\"]/span/input[1]"));
            WebElement nationalityField = driver.findElement(By.xpath("//*[@id='input_box_1665559991']"));
            WebElement panField = driver.findElement(By.xpath("//*[@id='input_box_1665560012148']"));
            WebElement citizenshipField = driver.findElement(By.xpath("//*[@id='input_box_1665560036416']"));
            WebElement issueDateField = driver.findElement(By.xpath("//*[@id='date_box_1665560118558_field']/span/input[1]"));
            WebElement accountTypeField = driver.findElement(By.xpath("//*[@id='select_1665560156']"));
            WebElement grandFathersNameField = driver.findElement(By.xpath("//*[@id='input_box_1665560420']"));
            WebElement fathersNameField = driver.findElement(By.xpath("//*[@id='input_box_1665560500990']"));
            WebElement mothersNameField = driver.findElement(By.xpath("//*[@id='input_box_1665560501718']"));
            WebElement spouseNameField = driver.findElement(By.xpath("//*[@id='input_box_1665983718']"));
            WebElement houseNumberField = driver.findElement(By.xpath("//*[@id='input_box_1665561122']"));
            WebElement wardNumberField = driver.findElement(By.xpath("//*[@id='input_box_1665561204560']"));
            WebElement streetNameField = driver.findElement(By.xpath("//*[@id='input_box_1665561212656']"));
            
            WebElement districtField = driver.findElement(By.xpath("//*[@id=\"input_box_1665561289315\"]"));
            WebElement vdcField = driver.findElement(By.xpath("//*[@id=\"input_box_1665561216434\"]"));
            WebElement locationField = driver.findElement(By.xpath("//*[@id=\"input_box_1665561291327\"]"));
            WebElement workStatusField = driver.findElement(By.xpath("//*[@id=\"select_1665561503\"]"));
            WebElement purposeField = driver.findElement(By.xpath("//*[@id=\"select_1665561539775\"]"));
            WebElement incomeField = driver.findElement(By.xpath("//*[@id=\"select_1665561549284\"]"));
            WebElement salaryField = driver.findElement(By.xpath("//*[@id=\"number_box_1665561942\"]"));
            WebElement internetBankingCheckbox = driver.findElement(By.xpath("//*[@id=\"check_box_1665562220_Internet Banking\"]"));
            WebElement electronicStatementCheckbox = driver.findElement(By.xpath("//*[@id=\"check_box_1665562296_Electronic Statement\"]"));
            WebElement smsAlertServiceCheckbox = driver.findElement(By.xpath("//*[@id=\"check_box_1665562296_SMS Alert Service\"]"));
            WebElement citizenshipFrontField = driver.findElement(By.xpath("//*[@id=\"check_box_1665562296_SMS Alert Service\"]"));

            /*
*/
            // Add time intervals after filling each field
            Select selectSalutation = new Select(salutationField);
            selectSalutation.selectByVisibleText(salutation);
         

            firstNameField.clear();
            firstNameField.sendKeys(firstName);
      

            lastNameField.clear();
            lastNameField.sendKeys(lastName);

            emailField.clear();
            emailField.sendKeys(userEmail);

            passwordField.clear();
            passwordField.sendKeys(userPassword);
            
            
            dobField.click();
           
            WebElement dobYearField = driver.findElement(By.xpath("/html/body/div[6]/div[1]/div/div/div/input"));
            dobYearField.clear();
            dobYearField.sendKeys(dobYear);
            
            Thread.sleep(200);
            
            WebElement dobMonthField = driver.findElement(By.xpath("/html/body/div[6]/div[1]/div/div/select"));
            Select selectDobMonth = new Select(dobMonthField);
            selectDobMonth.selectByVisibleText(monthMap.get(dobMonth));
            
            Thread.sleep(200);
            WebElement dobDateField = driver.findElement(By.xpath("/html/body/div[6]/div[2]/div/div[2]/div/span["+dobDay+"]"));
            dobDateField.click();
         
            

            
            
            

            
            nationalityField.clear();
            nationalityField.sendKeys(nationality);
            
            panField.clear();
            panField.sendKeys(panNo);
        

            citizenshipField.clear();
            citizenshipField.sendKeys(citizenshipNo);
            
            issueDateField.click();
            WebElement issueYearField = driver.findElement(By.xpath("/html/body/div[7]/div[1]/div/div/div/input"));
            issueYearField.clear();
            issueYearField.sendKeys(issueYear);
            
            Thread.sleep(200);
            
            WebElement issueMonthField = driver.findElement(By.xpath("/html/body/div[7]/div[1]/div/div/select"));
            Select selectIssueMonth = new Select(issueMonthField);
            selectIssueMonth.selectByVisibleText(monthMap.get(issueMonth));
            
            Thread.sleep(200);
            WebElement issueDayField = driver.findElement(By.xpath("/html/body/div[7]/div[2]/div/div[2]/div/span["+issueDay+"]"));
            issueDayField.click();
            
            Select selectAccountType = new Select(accountTypeField);
            selectAccountType.selectByVisibleText(accountType);
           

            grandFathersNameField.clear();
            grandFathersNameField.sendKeys(grandFathersName);

            fathersNameField.clear();
            fathersNameField.sendKeys(fathersName);

            mothersNameField.clear();
            mothersNameField.sendKeys(mothersName);

            spouseNameField.clear();
            spouseNameField.sendKeys(spouseName);

            houseNumberField.clear();
            houseNumberField.sendKeys(houseNumber);
         

            wardNumberField.clear();
            wardNumberField.sendKeys(wardNumber);

            streetNameField.clear();
            streetNameField.sendKeys(streetName);
   

            districtField.clear();
            districtField.sendKeys(district);
       

            vdcField.clear();
            vdcField.sendKeys(vdcRuralMunicipality);
        

            locationField.clear();
            locationField.sendKeys(location);
       

            
            Select selectWorkStatus = new Select(workStatusField);
            selectWorkStatus.selectByVisibleText(workStatus);
            
            Select selectPurpose = new Select(purposeField);
            selectPurpose.selectByVisibleText(purposeOfAccount);
            
            Select selectIncome = new Select(incomeField);
            selectIncome.selectByVisibleText(sourceOfIncome);
            
            salaryField.clear();
            salaryField.sendKeys(salary);
      
            
            if (electronicBanking == 1) {
                if (!internetBankingCheckbox.isSelected()) {
                    internetBankingCheckbox.click();  // Check the checkbox
                }
            } else {
                if (internetBankingCheckbox.isSelected()) {
                    internetBankingCheckbox.click();  // Uncheck the checkbox
                }
            }
            
            if (electronicStatement == 1) {
                if (!electronicStatementCheckbox.isSelected()) {
                    electronicStatementCheckbox.click();  // Check the checkbox
                }
            } else {
                if (electronicStatementCheckbox.isSelected()) {
                    electronicStatementCheckbox.click();  // Uncheck the checkbox
                }
            }
            
            if (smsAlertService == 1) {
                if (!smsAlertServiceCheckbox.isSelected()) {
                    smsAlertServiceCheckbox.click();  // Check the checkbox
                }
            } else {
                if (smsAlertServiceCheckbox.isSelected()) {
                    smsAlertServiceCheckbox.click();  // Uncheck the checkbox
                }
            }
            
            
            /*
            mobileNumberField.clear();
            mobileNumberField.sendKeys(mobileNumber);
            Thread.sleep(2000);

            dobField.clear();
            dobField.sendKeys(dob);
            Thread.sleep(2000);


            
            issueDateField.clear();
            issueDateField.sendKeys(issueDate);
            Thread.sleep(2000);
             */
            // Click the submit button
            WebElement submitButton = driver.findElement(By.xpath("//*[@id=\"user-registration-form-781\"]/form/div[13]/button"));
            Thread.sleep(2000); // Add delay before clicking submit
            submitButton.click();
        }
    }


    @AfterTest
    public void tearDown() throws IOException {
        if (workbook != null) {
            workbook.close();
        }
        if (driver != null) {
            driver.quit();
        }
    }
}
