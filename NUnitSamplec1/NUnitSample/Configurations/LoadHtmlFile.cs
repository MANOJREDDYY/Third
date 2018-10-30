using AventStack.ExtentReports;
using AventStack.ExtentReports.MarkupUtils;
using AventStack.ExtentReports.Reporter;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NUnitSample.Configurations
{
    
    class LoadHtmlFile
    {
        public static ExtentReports extent;
        public static ExtentTest test;
        public static ExtentHtmlReporter htmlreporter;
        public static void SetupReporting(string reportname)
        {
            //htmlreporter = new ExtentHtmlReporter("D:\\Nunit_Workspace\\NUnitSamplec\\NUnitSample\\Reports\\SampleFile.html");
            htmlreporter = new ExtentHtmlReporter(reportname);
            htmlreporter.AppendExisting = true;
            extent = new ExtentReports();
            extent.AttachReporter(htmlreporter);
            Console.WriteLine("Started the test test 1 ");
            extent.AddSystemInfo("Host Name", "Manoj");
            extent.AddSystemInfo("Environment", "QA");
            extent.AddSystemInfo("User Name", "Manoj");
        }
        public static ExtentTest CreateTestwrap(string description)
        {
            test = extent.CreateTest(description, "");
            IMarkup m = MarkupHelper.CreateLabel(description, ExtentColor.Red);
            return test;
        }
        public static void GenerateReport()
        {
            extent.Flush();
        }

        public LoadHtmlFile()
        {
           
        }
    }
}
