using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Service;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Threading;

namespace ConsoleApp1
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            WindowsDriver driver;

            var appiumLocalService = new AppiumServiceBuilder()
                .UsingPort(4723)
                .WithStartUpTimeOut(TimeSpan.FromSeconds(30))
                .Build();
            appiumLocalService.Start();

            AppiumOptions desiredCapabilities = new AppiumOptions();
            desiredCapabilities.App = @"C:\AxStream\Signed_Release\AxStream_x64 v3.10.6.8321\AxSTREAM.exe";
            desiredCapabilities.AutomationName = "Windows";
            desiredCapabilities.AddAdditionalAppiumOption("ms:WaitForAppLaunch", "20");
            driver = new WindowsDriver(appiumLocalService, desiredCapabilities);

            //driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
            Thread.Sleep(2000);
            driver.FindElement(By.XPath("//*[@Name= ' File    ']")).Click();
            driver.FindElement(By.XPath("//*[@Name= 'Open...\tCtrl+O']")).Click();
            driver.FindElement(By.XPath("//*[@AutomationId='1001']")).Click();
            driver.FindElement(By.XPath("//*[@Name= 'Address']")).SendKeys(@"C:\Users\IAB\Desktop\TS_Export");
            driver.FindElement(By.XPath("//*[@AutomationId= '41477']")).SendKeys(Keys.Return);
            driver.FindElement(By.XPath("//*[@AutomationId= '1148']")).SendKeys($@"Mark15_inducer.axx");
            driver.FindElement(By.XPath("//*[@AutomationId= '1' and @Name= 'Open']")).Click();
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
            driver.FindElement(By.XPath("//*[@Name= 'Viewer']")).Click();

            var element = driver.FindElement(By.XPath("//*[@AutomationId='59664']"));
            element.Click();
            element.SendKeys(Keys.Shift + Keys.F10); // замена ПКМ

            var DesktopSession = new AppiumServiceBuilder()
                .UsingPort(4723)
                .Build();
            DesktopSession.Start();
            AppiumOptions DesktopCapabilities = new AppiumOptions();
            DesktopCapabilities.App = "Root";
            DesktopCapabilities.AutomationName = "Windows";
            WindowsDriver driver2 = new WindowsDriver(DesktopSession, DesktopCapabilities);

            var AddWatcherElement = driver2.FindElement(By.XPath("//*[@Name= 'Export ...']"));
            //WebDriverWait Desktopwait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            //Desktopwait.Until(pred => AddWatcherElement.Displayed);                           // задержка пока что-то не откроется
            AddWatcherElement.Click();
            driver.FindElement(By.XPath("//*[@Name= '3D Components']")).Click();

            //driver.Quit();
        }
    }
}