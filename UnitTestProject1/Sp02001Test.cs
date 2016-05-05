using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using OpenQA.Selenium.IE;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Support.Extensions;

using SpoRecIS;

namespace UnitTestProject1
{
    [TestClass]
    public class Sp02001Test
    {
        private IWebDriver _driver;

        public Sp02001Test()        {
            this._driver = new ChromeDriver();
        }

        [TestMethod]
        public void TestOpen()
        {
            using (BasePage target = Sp02001.Open(this._driver))
            {
                Assert.AreNotEqual(this._driver.PageSource.IndexOf("利用者情報入力"),-1);
            }
        }

        ~Sp02001Test()
        {
            this._driver.Quit();
        }

    }
}
