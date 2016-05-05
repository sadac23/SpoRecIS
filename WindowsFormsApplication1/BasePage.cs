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
    public abstract class BasePage : IDisposable
    {
        private IWebDriver _driver;

        protected BasePage(IWebDriver driver)
        {
            this._driver = driver;
        }

        void IDisposable.Dispose()
        {
            this._driver = null;
        }
    }
}
