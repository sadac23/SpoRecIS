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

using ReadWriteCsv;

namespace WindowsFormsApplication1
{
    /// <summary>
    /// スポレク入力支援
    /// </summary>
    public partial class Form1 : Form
    {
        private string _url = string.Empty;
        private string _inputFile = string.Empty;
        private ChromeDriver _driver = null;
        private ILog _log = null;
        private string _errorMessage = "システムエラーが発生しました。";

        public Form1()
        {
            try
            {
                InitializeComponent();

                FileVersionInfo ver = FileVersionInfo.GetVersionInfo(Assembly.GetExecutingAssembly().Location);
                this.Text = ver.Comments + "  ver." + ver.ProductVersion;

                this._driver = new ChromeDriver();
                this._driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(Double.Parse(ConfigurationManager.AppSettings["wait_seconds"])));
                this._inputFile = ConfigurationManager.AppSettings["input_file"];

                this._log = Log.getLogger();

                //ログファイル名取得
                Logger rootLogger = ((Hierarchy)this._log.Logger.Repository).Root;
                FileAppender appender = (FileAppender)rootLogger.GetAppender("MylogAppender") as FileAppender;
                string logFilename = appender.File;

                this._errorMessage = this._errorMessage
                    + "\r\n\r\n管理者に調査を依頼する場合はメールでログファイルを送付してください。"
                    + "\r\n[ログファイル]" + logFilename
                +"\r\n[送付先]satosatosato11112222@hotmail.com";
            }
            catch(Exception ex)
            {
                MessageBox.Show(this._errorMessage, "システムエラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 抽選申込結果ボタン押下処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            string queryResultString = string.Empty;

            try
            {
                this._log.Info("** 抽選申込結果ボタン押下処理 - start **");

                // Read sample data from CSV file
                using (CsvFileReader reader = new CsvFileReader(this._inputFile))
                {
                    using (CsvFileWriter writer = new CsvFileWriter(ConfigurationManager.AppSettings["output_file"]))
                    {
                        //セパレータの設定
                        reader.Separator = Convert.ToChar(ConfigurationManager.AppSettings["csv_separator"].Replace(@"\t", "\t"));
                        writer.Separator = Convert.ToChar(ConfigurationManager.AppSettings["csv_separator"].Replace(@"\t", "\t"));

                        CsvRow readRow = new CsvRow();

                        while (reader.ReadRow(readRow))
                        {
                            //抽選申込結果文字列を取得
                            queryResultString = this.GetQueryResults(readRow[0], readRow[1]);

                            //結果書き出し
                            CsvRow writeRow = new CsvRow();
                            writeRow.Add(readRow[0]);
                            writeRow.Add(readRow[1]);
                            writeRow.Add(queryResultString);
                            writer.WriteRow(writeRow);

                            if (this.GetGroupBox1Value() == this.radioButton1.Text)
                            {
                                SaveScreenshot(readRow[0] + ConfigurationManager.AppSettings["screenshot_file_extension"]);
                            }
                        }
                    }
                }

                this._log.Info("** 抽選申込結果ボタン押下処理 - end **");
            }
            catch (Exception ex)
            {
                this._log.Error("** 抽選申込結果ボタン押下処理 - error **", ex);
                MessageBox.Show(this._errorMessage, "システムエラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// 未使用
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            _url = "https://www.net.city.nagoya.jp/cgi-bin/sp01001";
            OpenQA.Selenium.Chrome.ChromeDriver driver = new OpenQA.Selenium.Chrome.ChromeDriver();
            driver.Url = _url;

            //利用者番号
            IWebElement id = driver.FindElement(By.Name("id"));
            id.SendKeys("1278192");

            //暗証番号
            IWebElement pass = driver.FindElement(By.Name("pass"));
            pass.SendKeys("1234");

            //地域
            SelectElement area = new SelectElement(driver.FindElement(By.Name("area")));
            area.SelectByValue("04");

            //利用施設
            SelectElement sisetu = new SelectElement(driver.FindElement(By.Name("sisetu")));
            sisetu.SelectByValue("5102");

            //種目
            SelectElement syumoku = new SelectElement(driver.FindElement(By.Name("syumoku")));
            syumoku.SelectByValue("001");

            //利用月日
            SelectElement month = new SelectElement(driver.FindElement(By.Name("month")));
            month.SelectByValue("09");

            SelectElement day = new SelectElement(driver.FindElement(By.Name("day")));
            day.SelectByValue("06");

            //供用区分
            SelectElement time = new SelectElement(driver.FindElement(By.Name("time")));
            time.SelectByValue("03");

            //申込
            IWebElement B1 = driver.FindElement(By.Name("B1"));
            B1.Click();

            //ダイアログ
            OpenQA.Selenium.IAlert alert = driver.SwitchTo().Alert();
            alert.Accept();

        }

        /// <summary>
        /// 抽選申し込み情報の入力を行い申し込みボタンを押下する
        /// </summary>
        /// <param name="entryId">入力利用者番号</param>
        /// <param name="entryPass">入力暗証番号</param>
        /// <param name="entryMonth">入力月</param>
        /// <param name="entryDay">入力日</param>
        /// <param name="entryArea">入力地域コード</param>
        /// <param name="entrySisetu">入力施設コード</param>
        /// <param name="entryTime">入力時間帯コード</param>
        private void Register(string entryId, string entryPass, string entryMonth
                            , string entryDay, string entryArea, string entrySisetu, string entryTime)
        {
            try
            {
                //画面遷移
                this._driver.Navigate().GoToUrl(ConfigurationManager.AppSettings["url_sp01001"]);

                //利用者番号
                IWebElement id = this._driver.FindElement(By.Name("id"));
                id.SendKeys(entryId);

                //暗証番号
                IWebElement pass = this._driver.FindElement(By.Name("pass"));
                pass.SendKeys(entryPass);

                //地域
                SelectElement area = new SelectElement(this._driver.FindElement(By.Name("area")));
                area.SelectByValue(entryArea);

                //利用施設
                SelectElement sisetu = new SelectElement(this._driver.FindElement(By.Name("sisetu")));
                sisetu.SelectByValue(entrySisetu);

                //種目
                SelectElement syumoku = new SelectElement(this._driver.FindElement(By.Name("syumoku")));
                syumoku.SelectByValue(ConfigurationManager.AppSettings["syumoku"]);

                //利用月日
                SelectElement month = new SelectElement(this._driver.FindElement(By.Name("month")));
                month.SelectByValue(entryMonth);

                SelectElement day = new SelectElement(this._driver.FindElement(By.Name("day")));
                day.SelectByValue(entryDay);

                //供用区分
                SelectElement time = new SelectElement(this._driver.FindElement(By.Name("time")));
                time.SelectByValue(entryTime);

                //申込
                IWebElement B1 = this._driver.FindElement(By.Name("B1"));
                B1.Click();

                //ダイアログ
                WebDriverWait wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(Double.Parse(ConfigurationManager.AppSettings["wait_seconds"])));
                IAlert alert = wait.Until(drv => drv.SwitchTo().Alert());
                string alertText = alert.Text;
                alert.Accept();
                
                //                B1 = this._driver.FindElement(By.Name("B1"));   //ダイアログ取得前に待機

                //WebDriverWait wait = new WebDriverWait(this._driver, TimeSpan.FromSeconds(Double.Parse(ConfigurationManager.AppSettings["wait_seconds"])));
                //wait.until(ExpectedConditions.alertIsPresent());
                
//                WebDriverWait wait = new WebDriverWait(this._driver, TimeSpan.FromSeconds(3));
//                IWebElement element = wait.Until(driver => driver.FindElement(By.Name("q")));
//                IWebElement element = wait.Until(ExpectedConditions.AlertIsPresent<IWebDriver>());
//                IWebElement element = wait.Until<>(ExpectedConditions.AlertIsPresent<IWebDriver>());


//                OpenQA.Selenium.IAlert alert = this._driver.SwitchTo().Alert();
//                alert.Accept();

                //結果取得
                IWebElement return01001 = this._driver.FindElement(By.Name("return01001"));
            }
            catch (NoSuchElementException ex1)
            {
                this._log.Error("Register - error", ex1);
                
                //要素取得エラーは無視
            }
            catch (UnhandledAlertException ex2)
            {
                this._log.Error("Register - error", ex2);

                //ダイアログを取得し閉じる
                OpenQA.Selenium.IAlert alert = this._driver.SwitchTo().Alert();
                alert.Accept();
            }
            catch (NoAlertPresentException ex3)
            {
                this._log.Error("Register - error", ex3);

                //要素取得エラーは無視
            }
        }

        /// <summary>
        /// 抽選申込照会結果文字列を取得する
        /// </summary>
        /// <param name="entryId">入力利用者番号</param>
        /// <param name="entryPass">入力暗証番号</param>
        /// <returns>抽選申込結果文字列</returns>
        private string GetQueryResults(string entryId, string entryPass)
        {
            string resultString = string.Empty;

            try
            {
                //画面遷移
                this._driver.Navigate().GoToUrl(ConfigurationManager.AppSettings["url_sp02001"]);

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

                resultString = "申込あり";

                //申込が存在しないかを判定
                if (this._driver.PageSource.IndexOf(ConfigurationManager.AppSettings["not_apply_string"]) > 0)
                {
                    resultString = "申込なし";
                }
            }
            catch(NoSuchElementException ex1)
            {
                this._log.Error("GetQueryResults - error", ex1);

                //失敗文字列を返す
                resultString = "照会失敗";
            }

            return resultString;
        }

        /// <summary>
        /// 抽選申込ボタン押下処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            string resultString = string.Empty;
            string ssFilename = string.Empty;

            try
            {
                this._log.Info("** 抽選申込ボタン押下処理 - start **");

                // Read sample data from CSV file
                using (CsvFileReader reader = new CsvFileReader(this._inputFile))
                {
                    reader.Separator = Convert.ToChar(ConfigurationManager.AppSettings["csv_separator"].Replace(@"\t","\t"));
                    CsvRow row = new CsvRow();
                    while (reader.ReadRow(row))
                    {
                        Register(row[0], row[1], row[2], row[3], row[4], row[5], row[6]);
                        resultString = GetResultString();
                        WriteProcessLog(row[0], row[1], resultString);
                        if (this.GetGroupBox1Value() == this.radioButton1.Text)
                        {
                            ssFilename = row[0] + "_" + row[1] + "_" + row[2] + "_" + row[3] 
                                + "_" + row[4] + "_" + row[5] + "_" + row[6] 
                                + ConfigurationManager.AppSettings["screenshot_file_extension"];
                            SaveScreenshot(ssFilename);
                        }
                    }
                }

                this._log.Info("** 抽選申込ボタン押下処理 - end **");
            }
            catch(Exception ex)
            {
                this._log.Error("** 抽選申込ボタン押下処理 - error **", ex);
                MessageBox.Show(this._errorMessage, "システムエラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// スクリーンショットを保存する
        /// </summary>
        /// <param name="filename"></param>
        private void SaveScreenshot(string filename) 
        {
            Screenshot ss = ((ITakesScreenshot)this._driver).GetScreenshot();
            ss.SaveAsFile(Path.Combine(ConfigurationManager.AppSettings["screenshot_save_directry"], filename), ImageFormat.Png);
        }

        /// <summary>
        /// 処理ログに書き込む（未使用）
        /// </summary>
        /// <param name="entryId"></param>
        /// <param name="entryPass"></param>
        /// <param name="entryResultString"></param>
        private void WriteProcessLog(string entryId, string entryPass, string entryResultString) 
        {
            Console.Write(entryId);
            Console.Write(entryPass);
            Console.Write(entryResultString);
            Console.WriteLine();
        }

        /// <summary>
        /// 申し込み結果を取得する
        /// </summary>
        /// <returns></returns>
        private string GetResultString()
        {
            return "○";
        }

        /// <summary>
        /// GroupBox1の選択値を取得する
        /// </summary>
        /// <returns></returns>
        private string GetGroupBox1Value()
        {
            string value = string.Empty;

            //GroupBox1の選択項目を取得 
            foreach (System.Windows.Forms.RadioButton rbSS in groupBox1.Controls)
            {
                if (rbSS.Checked)
                {
                    value = rbSS.Text;
                }
            }
            return value;
        }

    }
}
