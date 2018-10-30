using NUnit.Framework;
using OpenQA.Selenium;
using NUnitSample.BaseClass;
using System.Threading;
using System;

namespace NUnitSample 
{
    [TestFixture]
    [Parallelizable]
    public class TestClass : BaseTest
    {
        [Test]
        public void TestMethod()
        {
            Thread.Sleep(2000);
            Console.Write("Going to Start Chrome browser");
            driver.FindElement(By.XPath("//*[@id='email']"));
            driver.FindElement(By.XPath("//*[@id='email']")).SendKeys("Nunit framework");
            Thread.Sleep(5000);
        }
        [Test]
        public void TestMethod2()
        {
            Thread.Sleep(5000);
            Console.Write("Method 2");
            driver.FindElement(By.XPath("//*[@id='email']"));
            driver.FindElement(By.XPath("//*[@id='email']")).SendKeys("Nunit method 2");
            Thread.Sleep(5000);
        }
    }
}

