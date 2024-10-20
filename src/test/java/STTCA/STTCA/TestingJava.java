package STTCA.STTCA;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import io.github.bonigarcia.wdm.WebDriverManager;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import java.io.File;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.DataFormatter;

public class TestingJava {

    // Create a logger instance
    private static final Logger logger = LogManager.getLogger(TestingJava.class);
    WebDriver driver;
    XSSFWorkbook workbook;
    XSSFSheet sheet;

    @BeforeTest
    public void setup() throws IOException, InvalidFormatException {
        logger.info("Setting up the WebDriver...");
        WebDriverManager.chromedriver().setup();
        driver = new ChromeDriver();
        driver.get("https://demo.wpeverest.com/user-registration/customer-account-opening-form/");
        driver.manage().window().maximize();
        
        // Load Excel file
        File scrfile = new File("C:\\Users\\dell\\eclipse-workspace\\STTCA\\target\\Stt.xlsx");
        workbook = new XSSFWorkbook(scrfile);
        sheet = workbook.getSheetAt(0); // Access the first sheet
        logger.info("Excel file loaded successfully.");
    }

    @Test
    public void formSubmissionTest() throws IOException {
        logger.info("Starting form submission test...");
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
        
        DataFormatter dataFormatter = new DataFormatter();

        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            XSSFRow row = sheet.getRow(i);

            // Fetch data from each column
            String salutation = row.getCell(1).getStringCellValue();
            String firstName = row.getCell(2).getStringCellValue();
            String lastName = row.getCell(3).getStringCellValue();
            String userEmail = row.getCell(4).getStringCellValue();
            String userPassword = row.getCell(5).getStringCellValue();
            String mobileNumber = dataFormatter.formatCellValue(row.getCell(6));
            Date dobDate = row.getCell(7).getDateCellValue();
            DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
            String dob = dateFormat.format(dobDate);
            
            String[] dateParts = dob.split("-");
            String dobYear = dateParts[0];
            String dobMonth = dateParts[1];
            String dobDay = dateParts[2];

            String nationality = row.getCell(8).getStringCellValue();
            String panNo = row.getCell(9).getStringCellValue();
            String citizenshipNo = String.valueOf((long) row.getCell(10).getNumericCellValue());
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

            // Fill in the form on the web page using the field names (adjust locators as necessary)
            try {
                WebElement salutationField = driver.findElement(By.xpath("//*[@id=\"select_1665559734\"]"));
                Select selectSalutation = new Select(salutationField);
                selectSalutation.selectByVisibleText(salutation);
                logger.info("Selected salutation: " + salutation);
                
                driver.findElement(By.xpath("//*[@id=\"first_name\"]")).sendKeys(firstName);
                logger.info("Entered first name: " + firstName);
                
                driver.findElement(By.xpath("//*[@id=\"last_name\"]")).sendKeys(lastName);
                logger.info("Entered last name: " + lastName);
                
                driver.findElement(By.xpath("//*[@id=\"user_email\"]")).sendKeys(userEmail);
                logger.info("Entered email: " + userEmail);
                
                driver.findElement(By.xpath("//*[@id=\"user_pass\"]")).sendKeys(userPassword);
                logger.info("Entered password.");

                // Fill the date of birth
                WebElement dobField = driver.findElement(By.xpath("//*[@id=\"date_box_1665559930_field\"]/span/input[1]"));
                dobField.click();
                driver.findElement(By.xpath("/html/body/div[6]/div[1]/div/div/div/input")).sendKeys(dobYear);
                logger.info("Entered date of birth year: " + dobYear);
                Select dobMonthSelect = new Select(driver.findElement(By.xpath("/html/body/div[6]/div[1]/div/div/select")));
                dobMonthSelect.selectByVisibleText(monthMap.get(dobMonth));
                logger.info("Selected date of birth month: " + monthMap.get(dobMonth));
                driver.findElement(By.xpath("/html/body/div[6]/div[2]/div/div[2]/div/span[" + dobDay + "]")).click();
                logger.info("Selected date of birth day: " + dobDay);

                driver.findElement(By.name("phone_1665559881")).sendKeys(mobileNumber);
                logger.info("Entered mobile number: " + mobileNumber);
                
                driver.findElement(By.xpath("//*[@id='input_box_1665569991']")).sendKeys(nationality);
                logger.info("Entered nationality: " + nationality);
                
                driver.findElement(By.xpath("//*[@id='input_box_1665560012148']")).sendKeys(panNo);
                logger.info("Entered PAN: " + panNo);
                
                driver.findElement(By.xpath("//*[@id='input_box_1665560036416']")).sendKeys(citizenshipNo);
                logger.info("Entered citizenship number: " + citizenshipNo);
                
                // Fill the issue date
                WebElement issueDateField = driver.findElement(By.xpath("//*[@id='date_box_1665560118558_field']/span/input[1]"));
                issueDateField.click();
                driver.findElement(By.xpath("/html/body/div[7]/div[1]/div/div/div/input")).sendKeys(issueYear);
                logger.info("Entered issue date year: " + issueYear);
                Select issueMonthSelect = new Select(driver.findElement(By.xpath("/html/body/div[7]/div[1]/div/div/select")));
                issueMonthSelect.selectByVisibleText(monthMap.get(issueMonth));
                logger.info("Selected issue date month: " + monthMap.get(issueMonth));
                driver.findElement(By.xpath("/html/body/div[7]/div[2]/div/div[2]/div/span[" + issueDay + "]")).click();
                logger.info("Selected issue date day: " + issueDay);

                Select selectAccountType = new Select(driver.findElement(By.xpath("//*[@id='select_1665560156']")));
                selectAccountType.selectByVisibleText(accountType);
                logger.info("Selected account type: " + accountType);

                driver.findElement(By.xpath("//*[@id='grandfather_name']")).sendKeys(grandFathersName);
                logger.info("Entered grandfather's name: " + grandFathersName);
                
                driver.findElement(By.xpath("//*[@id='father_name']")).sendKeys(fathersName);
                logger.info("Entered father's name: " + fathersName);
                
                driver.findElement(By.xpath("//*[@id='mother_name']")).sendKeys(mothersName);
                logger.info("Entered mother's name: " + mothersName);
                
                driver.findElement(By.xpath("//*[@id='spouse_name']")).sendKeys(spouseName);
                logger.info("Entered spouse's name: " + spouseName);
                
                driver.findElement(By.xpath("//*[@id='house_no']")).sendKeys(houseNumber);
                logger.info("Entered house number: " + houseNumber);
                
                driver.findElement(By.xpath("//*[@id='ward_no']")).sendKeys(wardNumber);
                logger.info("Entered ward number: " + wardNumber);
                
                driver.findElement(By.xpath("//*[@id='street_name']")).sendKeys(streetName);
                logger.info("Entered street name: " + streetName);
                
                driver.findElement(By.xpath("//*[@id='district']")).sendKeys(district);
                logger.info("Entered district: " + district);
                
                driver.findElement(By.xpath("//*[@id='vdc_or_rural_municipality']")).sendKeys(vdcRuralMunicipality);
                logger.info("Entered VDC/Rural Municipality: " + vdcRuralMunicipality);
                
                driver.findElement(By.xpath("//*[@id='location']")).sendKeys(location);
                logger.info("Entered location: " + location);
                
                Select selectWorkStatus = new Select(driver.findElement(By.xpath("//*[@id='select_1665560795']")));
                selectWorkStatus.selectByVisibleText(workStatus);
                logger.info("Selected work status: " + workStatus);
                
                Select selectPurposeOfAccount = new Select(driver.findElement(By.xpath("//*[@id='select_1665560807']")));
                selectPurposeOfAccount.selectByVisibleText(purposeOfAccount);
                logger.info("Selected purpose of account: " + purposeOfAccount);
                
                Select selectSourceOfIncome = new Select(driver.findElement(By.xpath("//*[@id='select_1665560826']")));
                selectSourceOfIncome.selectByVisibleText(sourceOfIncome);
                logger.info("Selected source of income: " + sourceOfIncome);
                
                driver.findElement(By.xpath("//*[@id='input_box_1665560835']")).sendKeys(salary);
                logger.info("Entered salary: " + salary);
                
                // Checkboxes for services
                if (electronicBanking == 1) {
                    driver.findElement(By.xpath("//*[@id='electronic_banking']")).click();
                    logger.info("Selected electronic banking.");
                }
                if (electronicStatement == 1) {
                    driver.findElement(By.xpath("//*[@id='electronic_statement']")).click();
                    logger.info("Selected electronic statement.");
                }
                if (smsAlertService == 1) {
                    driver.findElement(By.xpath("//*[@id='sms_alert_service']")).click();
                    logger.info("Selected SMS alert service.");
                }
                
                // Upload citizenship front image
                driver.findElement(By.xpath("//*[@id='input_box_1665560967']")).sendKeys(citizenshipFront);
                logger.info("Uploaded citizenship front image: " + citizenshipFront);
                
                // Submit the form
                driver.findElement(By.xpath("//*[@id='user-registration-form-781']/form/div[13]/button")).click();
                logger.info("Form submitted successfully for: " + firstName + " " + lastName);

                // Wait for some confirmation element to appear (adjust as necessary)
                WebElement confirmationMessage = driver.findElement(By.xpath("//your-confirmation-xpath-here"));
                if (confirmationMessage.isDisplayed()) {
                    logger.info("Confirmation message displayed: " + confirmationMessage.getText());
                }

            } catch (Exception e) {
                logger.error("An error occurred while processing row " + (i + 1), e);
            }
        }
    }

    @AfterTest
    public void tearDown() {
        if (driver != null) {
            driver.quit();
            logger.info("Closed the WebDriver.");
        }
        try {
            if (workbook != null) {
                workbook.close();
                logger.info("Workbook closed.");
            }
        } catch (IOException e) {
            logger.error("Error while closing the workbook.", e);
        }
    }

    public static void main(String[] args) {
        logger.info("Starting the application...");
        // The rest of your main logic (if any)...
    }
}
