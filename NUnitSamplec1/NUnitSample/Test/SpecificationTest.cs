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
using Microsoft.Office.Interop.Excel;
using Microsoft.CSharp;
using NUnitSample.Configurations;
using _Excel= Microsoft.Office.Interop.Excel;

namespace NUnitSample.Test
{
    [TestFixture, SingleThreaded]

    class SpecificationTest
    {
        ExtentTest test;
        IWebDriver driver;
        private static ExtentReports extent;
        [OneTimeSetUp]



        public void reportGeneration()
        {
            string reportname = "RSC product Spec Report";
            LoadHtmlFile.SetupReporting(reportname);
        }


        [OneTimeSetUp]
        public void BrowserLogin()
        {
        //  driver = new ChromeDriver();
        }
       
        private static ExcelFile ExcelFileInstance = null;
        private ExcelFile excelFile =getExcelFile();  
        LoadHtmlFile loadfile;
        public static ExcelFile getExcelFile()
           
        {
            if (ExcelFileInstance == null)
            {
                ExcelFileInstance = ExcelUtility.LoadExcel();
               
            }
            return ExcelFileInstance;
        }
        [Test]
        public void NavigateToSpecification()
        {
            test = LoadHtmlFile.CreateTestwrap("create a product for RSC product style ");
            test.Log(Status.Warning, "This is the final Log");
            LoadHtmlFile.SetupReporting("D:\\Nunit_Workspace\\NUnitSamplec\\NUnitSample\\Reports\\SampleFile1.html");
            test.Log(Status.Pass, MarkupHelper.CreateLabel("Manoj" + " Test Case PASSED", ExtentColor.Green));
            driver.Url = PropertyReader.GetProperty(Property.ProductMenu_ID, Locators.Specification.Default);
            driver.Manage().Window.Maximize();
            LoadHtmlFile.CreateTestwrap("New file for check");
            PropertyReader.GetProperty(Property.BASEURL, Locators.Specification.Default);   
            driver.FindElement(By.XPath(PropertyReader.GetProperty(Property.Product_Management, Locators.Specification.Default))).Click();
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            Thread.Sleep(1000);
            driver.FindElement(By.Id(PropertyReader.GetProperty(Property.Specifications, Locators.Specification.Default))).Click();
            test.Info("Now i am in the test");              
        }
        [Test]
        public void ExcelOperations()
        {
            //Excel excelfile = new Excel();
            Excel.KillExcelProcesses();
            Excel excelfile = new Excel();
            String path = "D:\\Nunit_Workspace\\NUnitSamplec\\NUnitSample\\ExcelIO\\CBSInputTestDataPostgress2.xls";

            excelfile.OpenExcelFile(path);
            // Excel excelfile = new Excel(path, "Specification", "Scenario", "Create_Specification", "Create_RSC_makeBoard");

            // Console.WriteLine(excelfile.getinputdata("Scenario", "Create_Specification", "Create_RSC_makeBoard", "SPECCUSTNAME"));
           // Console.WriteLine(excelfile.getinputdata("Scenario", "Create_Specification", "Create_Diecut_makeBoard", "SPECCUSTNAME"));
          //  Console.WriteLine(excelfile.getinputdata("Scenario", "Create_Specification", "Create_AssembledSet", "SPECPRODSTYLE"));
            // excelfile.writeexcelfile("Scenario", "Create_Specification", "Create_AssembledSet", "SPECPRODSTYLE","i am not updating");
            // excelfile.writeexcelfile("Scenario", "Create_Specification", "Create_AssembledSet", "SPECPRODSTYLE", "i am not updating");
           // Console.WriteLine(excelfile.getinputdata("Scenario", "Create_Specification", "Create_RSC_makeBoard", "SPECCUSTNAME"));
            //Console.WriteLine(excelfile.getinputdata("Scenario", "Manojtestsolution1", "test1", "SPECCUSTNAME"));
            //Console.WriteLine(excelfile.getinputdata("Scenario", "Manojtestsolution2", "test2", "SPECCUSTNAME"));
            Console.WriteLine(excelfile.GetInputData("Specification", "Manojtestsolution15", "test20", "SPECCUSTNAME"));
            Console.WriteLine(excelfile.GetInputData("Specification", "Create_Specification", "Create_HSC_makeBoard", "SPECPRODSTYLE"));
            Console.WriteLine(excelfile.GetInputData("Specification", "Create_Specification", "Create_Pad_makeBoard", "SPECPRODSTYLE"));
            excelfile.WriteExcelFile("Specification", "Create_Specification", "Create_AssembledSet", "SPECPRODSTYLE", "before update123");
            Console.WriteLine(excelfile.GetInputData("Specification", "Create_Specification", "Create_AssembledSet", "SPECPRODSTYLE"));
            excelfile.OpenSheet("ORDERS");
            excelfile.WriteExcelFile("ORDERS", "CREATEORDER", "create_MultipleShipto", "CUSTOMER_NAME", "This is going to be updatedeee");
            Console.WriteLine(excelfile.GetInputData("ORDERS", "CREATEORDER", "create_MultipleShipto", "CUSTOMER_NAME"));
            Console.WriteLine(excelfile.GetInputData("ORDERS", "CREATEORDER", "create_makeship_order", "CUSTOMER_NAME"));
            //Console.WriteLine(excelfile.GetInputData("Scenario", "CREATEORDER", "create_makeship_order", "CUSTOMER_NAME"));
            //  excelfile.writeexcelfile("Scenario", "Create_Specification", "Create_AssembledSet", "SPECPRODSTYLE", "after update");
            excelfile.WriteExcelFile("ORDERS", "CREATEORDER", "create_MultipleShipto", "CUSTOMER_NAME", "hello 123 not an");
            //excelfile.saveexcel();
            //Console.WriteLine(excelfile.getinputdata("Scenario", "Create_Specification", "Create_AssembledSet", "SPECPRODSTYLE"));

           // excelfile.SaveAndCloseExcel();
            //  excelfile.ReadCell(1,1);
             //Excel.KillExcelProcesses();
        }

      

        [OneTimeTearDown]
        public void closebrowser()
        {
            LoadHtmlFile.GenerateReport();
         //   driver.Quit();
        }


    }
}
