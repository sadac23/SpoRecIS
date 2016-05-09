using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OpenQA.Selenium.IE;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Support.Extensions;

namespace SpoRecIS
{
    /// <summary>
    /// 利用者情報入力（/sp02001）
    /// </summary>
    public class Sp02001 : BasePage
    {
        public Sp02001(IWebDriver driver) : base(driver)
        {
        }

        public static Sp02001 Open(IWebDriver driver)
        {
            driver.Navigate().GoToUrl("https://www.net.city.nagoya.jp/cgi-bin/sp02001");
            return new Sp02001(driver);
        }

        /// <summary>
        /// ログインを行う。
        /// </summary>
        /// <param name="id">ユーザID</param>
        /// <param name="pass">パスワード</param>
        /// <returns>抽選申込一覧</returns>
        public BasePage Login(string entryId, string entryPass)
        {
            // 利用者番号
            IWebElement id = this._driver.FindElement(By.Name("id"));
            id.SendKeys(entryId);

            // 暗証番号
            IWebElement pass = this._driver.FindElement(By.Name("pass"));
            pass.SendKeys(entryPass);

            // ＯＫ
            IWebElement B1 = this._driver.FindElement(By.Name("B1"));
            B1.Click();

            return null;
        }
    }
}
