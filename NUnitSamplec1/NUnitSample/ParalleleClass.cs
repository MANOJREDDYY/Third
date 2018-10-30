using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnitSample.BaseClass;
using System.Threading;
using OpenQA.Selenium;

namespace NUnitSample 
{
    [TestFixture]
    [Parallelizable]
    class ParalleleClass : BaseTest
    {
        
        [Test]
        [TestCaseSource(typeof(BaseTest),"BrowserToRunWith")]
        public void ParallelMethod(String browse)
       
        {
            Setup(browse);
            Thread.Sleep(2000);
            Console.Write("Going to Start Chrome browser2");
            driver.FindElement(By.XPath("//*[@id='email']"));
            driver.FindElement(By.XPath("//*[@id='email']")).SendKeys("Parallel class");
            Thread.Sleep(5000);
        }

    }
}
