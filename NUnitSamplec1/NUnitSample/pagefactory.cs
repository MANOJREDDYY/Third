using OpenQA.Selenium;
using OpenQA.Selenium.Support.PageObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NUnitSample
{
    class localt
    {
        [FindsBy(How = How.Id, Using = "username")]
        private IWebElement UserName { get; set; }
        public string googleurl="https://www.google.com";
    }
}
