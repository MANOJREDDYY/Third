using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium.Firefox;
using NUnitSample;
using NUnitSample.Properties;

namespace NUnitSample.BaseClass
{
    public class BaseTest
    {
        public IWebDriver driver;
        

        public void Setup(String browsername) {
            if (browsername.Equals("chrome"))
                driver = new ChromeDriver();
            else
                driver = new FirefoxDriver();
        }
            [OneTimeSetUp]
        public void open()
        {
// driver = new ChromeDriver();
            driver.Url = "https://www.facebook.com/";
            
        }
        [OneTimeTearDown]
        public void close()
        {
            driver.Quit();
        }
        public static IEnumerable<String> BrowserToRunWith()
        {
            String[] browsers = Resources.browser.Split(',');
            foreach (String b in browsers)
            {
                yield return b;
            }
        }
    }
}
