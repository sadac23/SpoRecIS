using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;
using System.IO;
using System.Drawing.Imaging;
using System.Diagnostics;
using System.Reflection;

using OpenQA.Selenium.IE;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Support.Extensions;

using log4net;
using log4net.Appender;
using log4net.Repository.Hierarchy;

namespace SpoRecIS
{
    public class LotteryApplicationConfirmation
    {
        ChromeDriver _driver = null;
        string _url = string.Empty;

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="log"></param>
        /// <param name="url"></param>
        public LotteryApplicationConfirmation(ChromeDriver driver, string url)
        {
            this._driver = driver;
            this._url = url;
        }

        /// <summary>
        /// 抽選結果文字列を取得する
        /// </summary>
        /// <returns></returns>
        public string GetQueryResults(string entryId, string entryPass){
            string resultString = string.Empty;

            try
            {
                //画面遷移
                this._driver.Navigate().GoToUrl(this._url);

                //利用者番号
                IWebElement id = this._driver.FindElement(By.Name("id"));
                id.SendKeys(entryId);

                //暗証番号
                IWebElement pass = this._driver.FindElement(By.Name("pass"));
                pass.SendKeys(entryPass);

                //ＯＫ
                IWebElement B1 = this._driver.FindElement(By.Name("B1"));
                B1.Click();

                //結果
                IWebElement divMain = this._driver.FindElement(By.Id("main"));

                resultString = "判定不能";

                //申込が存在しないかを判定
                if (this._driver.PageSource.IndexOf("残念ながらすべて落選しました。") > 0)
                {
                    resultString = "当選なし";
                }
                else if (this._driver.PageSource.IndexOf("すべての当選情報を確認されています。") > 0)
                {
                    resultString = "当選あり";
                } 
                else if (this._driver.PageSource.IndexOf("上記の内容で抽選申込を受け付けました。") > 0)
                {
                    resultString = "当選あり";
                }
            }
            catch (NoSuchElementException ex1)
            {
                //失敗文字列を返す
                resultString = "照会失敗";
            }

            return resultString;
        }
    }
}
