using AventStack.ExtentReports;
using AventStack.ExtentReports.MarkupUtils;
using AventStack.ExtentReports.Reporter;
using NUnit.Framework;
using System;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.PageObjects;
using System.Threading;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Interactions;
using System.Drawing;

namespace NUnitSample
{
    [TestFixture]  
    class SampleTestforLaunchBrowse
    {
        public static ExtentReports extent;
        IWebDriver driver = new ChromeDriver();
        ExtentTest test;
        ExtentTest test2;
        public static ExtentHtmlReporter htmlreporter;
        private static ExcelFile ExcelFileInstance = null;
        private ExcelFile excelFile = getExcelFile();
        public static ExcelFile getExcelFile()
        {
            if (ExcelFileInstance == null)
            {
                ExcelFileInstance = ExcelUtility.LoadExcel();
            }
            return ExcelFileInstance;
        }
     

        [Test]
        public void inapplication()
        {
            Console.WriteLine("Time " + String.Format("{0:MM}_{0:D}", DateTime.Now));
            String datetoday = String.Format("{0:D}", DateTime.Now).Replace(",","_").Replace(" ","").Replace("Wednesday","Wed");
            Console.WriteLine(datetoday);
            excelFile.getrowcount("");
            string Spec_CustName = excelFile.GetTestInputValue("SPECIFICATION", "Create_Specification", "Create_RSC_makeBoard", "SPECCUSTNAME");
            excelFile.getrowcount("");
            Console.WriteLine(Spec_CustName.ToLower());    
            htmlreporter = new ExtentHtmlReporter("D:\\Nunit_Workspace\\NUnitSamplec\\NUnitSample\\Reports\\"+datetoday+".html");
            htmlreporter.AppendExisting = true;
            extent = new ExtentReports();
            extent.AttachReporter(htmlreporter);
            Console.WriteLine("Started the test test 1 ");
            extent.AddSystemInfo("Host Name", "Manoj");
            extent.AddSystemInfo("Environment","QA");
            extent.AddSystemInfo("User Name", "Manoj");
            Console.WriteLine("Started the test test 2 ");
            test = extent.CreateTest("Create test wrap", "");
            test2 = extent.CreateTest("Second test", "");
            IMarkup m = MarkupHelper.CreateLabel("Create test wrap", ExtentColor.Red);
            test.Log(Status.Warning, "This is the final Log");
            test.Log(Status.Pass, MarkupHelper.CreateLabel("Manoj" + " Test Case PASSED", ExtentColor.Green));
            Console.WriteLine("Output1 : " + Property.ProductMenu_ID);
             driver.Url=PropertyReader.GetProperty(Property.ProductMenu_ID, Locators.Specification.Default);
            driver.Manage().Window.Maximize();
            PropertyReader.GetProperty(Property.BASEURL, Locators.Specification.Default);
            Console.WriteLine("Output : "+ PropertyReader.GetProperty(Property.ProductMenu_ID, Locators.Specification.Default));
            extent.AddTestRunnerLogs("Navigated to Gogle ");
            driver.FindElement(By.XPath(PropertyReader.GetProperty(Property.Product_Management, Locators.Specification.Default))).Click();
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            Thread.Sleep(1000);
            // wait.Until(ExpectedConditions.ElementToBeClickable(By.Id(PropertyReader.GetProperty(Property.Specifications, Specification.Default))));
           // var coomon = new Commonfunctions();
            driver.FindElement(By.Id(PropertyReader.GetProperty(Property.Specifications, Locators.Specification.Default))).Click();
         //   coomon.findelement(driver, "id", "Specifications").Click();
           // driver.FindElement(By.Id(PropertyReader.GetProperty(Property.Specifications, Specification.Default))).Click();
            extent.AddTestRunnerLogs("Launched the application");
            test.Info("Now i am in the test");
           extent.Flush();
            
        }
        
        [Test]
        public void mainbrowser()
        {
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(PropertyReader.GetProperty(Property.Createnewbutton, Locators.Specification.Default))));
            Thread.Sleep(2000);
            driver.FindElement(By.XPath(PropertyReader.GetProperty(Property.Createnewbutton, Locators.Specification.Default))).Click();
            Thread.Sleep(2000);
            driver.FindElement(By.XPath(PropertyReader.GetProperty(Property.customerSearch, Locators.Specification.Default))).Click(); Thread.Sleep(2000);
            //*[@id='createNew']
            Actions builder=new Actions(driver);
            builder.MoveToElement(driver.FindElement(By.XPath(PropertyReader.GetProperty(Property.customerSearch, Locators.Specification.Default)))).Click().Build().Perform();
           // driver.FindElement(By.XPath(PropertyReader.GetProperty(Property.customerSearch, Specification.Default))).Click();
            Thread.Sleep(2000);
            driver.FindElement(By.XPath(PropertyReader.GetProperty(Property.customerSearch, Locators.Specification.Default))).SendKeys("Activision/Blizzard");
            Thread.Sleep(2000);
            builder.MoveToElement(driver.FindElement(By.XPath(PropertyReader.GetProperty(Property.Customer, Locators.Specification.Default)))).DoubleClick().Build().Perform();
            //driver.FindElement(By.XPath(PropertyReader.GetProperty(Property.customerSearch, Specification.Default))).SendKeys("Activision/Blizzard");
            driver.FindElement(By.XPath(PropertyReader.GetProperty(Property.SpecId, Locators.Specification.Default))).SendKeys("reg126");
            Thread.Sleep(2000);
            driver.FindElement(By.XPath(PropertyReader.GetProperty(Property.ProductStyle, Locators.Specification.Default))).SendKeys("RSC JI TWLWL");
            Thread.Sleep(2000);
            String value1 = "//*[text()='RSC JI TWLWL']";
            driver.FindElement(By.XPath(value1)).Click();
            // builder.MoveToElement(driver.FindElement(By.XPath(PropertyReader.GetProperty(Property.ProductStyleID, Specification.Default)))).Click().Build().Perform();
            Thread.Sleep(2000);
            driver.FindElement(By.XPath(PropertyReader.GetProperty(Property.GLcode, Locators.Specification.Default))).Clear();
            Thread.Sleep(2000);
            driver.FindElement(By.XPath(PropertyReader.GetProperty(Property.GLcode, Locators.Specification.Default))).SendKeys("flexo");
            Thread.Sleep(2000);
            String value = "//*[text()='Flexo-Corporate-Corporate-Income']";
            driver.FindElement(By.XPath(value)).Click();
            driver.FindElement(By.XPath(PropertyReader.GetProperty(Property.GLcode, Locators.Specification.Default))).SendKeys("flexo");
            //*[text()='Product Length ']/parent::div/input
            //driver.FindElement(By.XPath("//*[text()="+value)).Click();
            //WebElement countrydropdown = driver.findElement(By.ClassName("ui-autocomplete-pagination-results"));
            // SelectElement country = new SelectElement(driver.FindElement(By.ClassName("ui-autocomplete-pagination-results")));
            //country.SelectByText("Flexo-Corporate-Corporate-Income");
            //builder.MoveToElement(driver.FindElement(By.XPath(PropertyReader.GetProperty(Property.GLCodeSelection, Specification.Default)))).DoubleClick().Build().Perform();
            // driver.FindElement(By.XPath(PropertyReader.GetProperty(Property.Customer, Specification.Default))).Click();
            // driver.FindElement(By.XPath(PropertyReader.GetProperty(Property.customerSearch, Specification.Default))).Click();
            var locato = new localt();
            //   PageFactory.InitElements(driver, locato);
            // driver.Url = locato.googleurl;
            driver.FindElement(By.XPath(PropertyReader.GetProperty(Property.ProductLength, Locators.Specification.Default))).SendKeys("10");
            driver.FindElement(By.XPath(PropertyReader.GetProperty(Property.ProductWidth, Locators.Specification.Default))).SendKeys("11");
            driver.FindElement(By.XPath(PropertyReader.GetProperty(Property.ProductDepth, Locators.Specification.Default))).SendKeys("12");
            driver.FindElement(By.XPath(PropertyReader.GetProperty(Property.MaterialGrade, Locators.Specification.Default))).SendKeys("32ECT K2C");
            Thread.Sleep(2000);
            String value3 = "//*[text()='32ECT K2C']";
            driver.FindElement(By.XPath(value3)).Click();
            test.Info("Now i am in the second test");
            Thread.Sleep(3000);
            driver.Quit();
        }
    }
}
