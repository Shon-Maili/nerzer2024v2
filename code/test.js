// Import necessary libraries
const { Builder, By, Key, until } = require('selenium-webdriver');
const chrome = require('selenium-webdriver/chrome');
const xlsx = require('xlsx');

async function fillOutForm() {
    // Set path to ChromeDriver executable
    const chromeDriverPath = './node_modules/chrome/chromedriver.exe';
    // Adjust the path as needed

    // Set Chrome options
    const chromeOptions = new chrome.Options();

    // Initialize WebDriver with Chrome
    const driver = await new Builder()
        .forBrowser('chrome')
        .setChromeOptions(chromeOptions)
        .setChromeService(new chrome.ServiceBuilder(chromeDriverPath))
        .build();

    try {
        // Navigate to the form page
        await driver.get('https://mushlam-frontend.wiz.digital.idf.il/m/SWYX5GMB4U');

        // Read data from Excel file
        const workbook = xlsx.readFile('C:/Users/Admin/Downloads/test.xlsx');
        const firstSheetName = workbook.SheetNames[0]; // Get the name of the first sheet
        const worksheet = workbook.Sheets[firstSheetName]; // Get the first sheet by name
        const column = worksheet['A']; // Assuming column A contains the "מספר אישי" data

        // Get all cell addresses in column A
        const columnAddresses = Object.keys(worksheet).filter(address => address.startsWith('A') && address !== 'A1');

        // Extract values from column A
        const columnValues = columnAddresses.map(address => worksheet[address].v);

        // Print all values in the column
        console.log('Values in column A:');
        console.log(columnValues);

        // Iterate through each cell in the column and fill out the form fields
        for (let rowNum = 2; rowNum <= columnValues.length + 1; rowNum++) {
            const misparIshi = columnValues[rowNum - 2]; // Adjusted to use zero-based index
            // Fill out the form fields
            await driver.findElement(By.css('[fieldid="fld_fname"]')).sendKeys('א');
            await driver.findElement(By.css('[fieldid="fld_firstLetter"]')).sendKeys('א');
            await driver.findElement(By.css('[fieldid="fld_phoneNum"]')).sendKeys('0505555555');
            await driver.findElement(By.css('[fieldid="fld_misparIshi"]')).sendKeys(misparIshi.toString());
            await driver.findElement(By.css('[fieldid="fld_hfname"]')).sendKeys('יוסי');
            await driver.findElement(By.css('[fieldid="fld_hlname"]')).sendKeys('יוסי');

            // Submit the form
            await driver.findElement(By.id('finishForm')).click();

            await driver.wait(async () => {
                const elements = await driver.findElements(By.css('.summaryPage'));
                for (const element of elements) {
                    const display = await element.getCssValue('display');
                    if (display === 'block') {
                        // Element found with display: block property
                        return true;
                    }
                }
                return false; // Element not found with display: block property
            });
            
            // Refresh the page
            await driver.navigate().refresh();
            


            console.log(`Form submitted for row ${rowNum - 1}`);

            // Navigate back to the form page URL
           // await driver.get('https://mushlam-frontend.wiz.digital.idf.il/m/SWYX5GMB4U');
        }
    } catch (error) {
        console.error('An error occurred:', error);
    } finally {
        // Close the browser
        await driver.quit();
    }
}

fillOutForm();