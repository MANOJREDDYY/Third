using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnitSample.Locators;
using OpenQA.Selenium;

namespace NUnitSample
{
    class Commonfunctions
    {
        public String getelement(IWebDriver driver,String elementname)
        {
            var element = "Property." + elementname;   
            return PropertyReader.GetProperty(element, Specification.Default);
        }
        public IWebElement findelement(IWebDriver driver,String locator, String elementname)
        {
            var element = "Property." + elementname;
            // String elementvalue = "By." + locator + "("+ element + ")";
            var ele= PropertyReader.GetProperty(element, Specification.Default);
            IWebElement Elemen;
            switch (locator.ToLower())
            {
                case "id":
                    Elemen = driver.FindElement(By.Id(ele));
                    break;
                case "xpath":
                    Elemen = driver.FindElement(By.XPath(ele));
                    break;
                case "linktext":
                    Elemen = driver.FindElement(By.LinkText(ele));
                    break;
                default:
                    Elemen = driver.FindElement(By.Id(""));
                    break;
            }
            return Elemen;
        }
    }
}
