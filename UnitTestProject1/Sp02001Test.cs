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
        [TestMethod]
        public void TestOpen()
        {
            using(IWebDriver driver = new ChromeDriver())
            {
                Sp02001 target = Sp02001.Open(driver);

                Assert.IsInstanceOfType(target, typeof(Sp02001));
                Assert.AreNotEqual(driver.PageSource.IndexOf("利用者情報入力"), -1);
            }
        }

        [TestMethod]
        public void TestLogin()
        {
            using (IWebDriver driver = new ChromeDriver())
            {
                Sp02001 target = Sp02001.Open(driver);
                target.Login("1230158", "56rthe20");

                Assert.AreNotEqual(driver.PageSource.IndexOf("抽選申込一覧"), -1);
            }
        }
    }
}
