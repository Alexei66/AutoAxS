using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Service;
using OpenQA.Selenium.Appium.Windows;
using System;

namespace ConsoleApp1
{
    internal class Program
    {
        [STAThread]
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
            //Thread.Sleep(2000);

            driver.Manage().Window.Maximize();

            driver.FindElement(By.XPath("//*[@Name= ' File    ']")).Click();
            driver.FindElement(By.XPath("//*[@Name= 'Open...\tCtrl+O']")).Click();
            driver.FindElement(By.XPath("//*[@AutomationId='1001']")).Click();

            // Копируем текст в буфер обмена

            SetClipboardAndPaste(driver, "//*[@Name= 'Address']", @"C:\Users\IAB\Desktop\TS_Export");

            driver.FindElement(By.XPath("//*[@AutomationId= '41477']")).SendKeys(Keys.Return);

            SetClipboardAndPaste(driver, "//*[@AutomationId= '1148']", @"Mark15_inducer.axx");

            driver.FindElement(By.XPath("//*[@AutomationId= '1' and @Name= 'Open']")).Click();
            //driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10); // задержка 10 сек
            driver.FindElement(By.XPath("//*[@Name= 'Viewer']")).Click();

            // кнопка с 3D видом

            var element = driver.FindElement(By.XPath("//*[@AutomationId='59664']"));
            element.Click();
            element.SendKeys(Keys.Shift + Keys.F10); // замена ПКМ

            /* //var DesktopSession = new AppiumServiceBuilder()
             //    .UsingPort(4723)
             //    .Build();
             //DesktopSession.Start();
             //AppiumOptions DesktopCapabilities = new AppiumOptions();
             //DesktopCapabilities.App = "Root";
             //DesktopCapabilities.AutomationName = "Windows";*/
            WindowsDriver driver2 = ContextWindow();

            var addWatcherElement = driver2.FindElement(By.XPath("//*[@Name= 'Export ...']"));
            //WebDriverWait Desktopwait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            //Desktopwait.Until(pred => AddWatcherElement.Displayed);                           // задержка пока что-то не откроется
            addWatcherElement.Click();

            /////////TurboGrid (CFX, FLUENT)/////////

            driver.FindElement(By.XPath("//*[@Name= 'TurboGrid (CFX, FLUENT)']")).Click();
            driver.FindElement(By.XPath("//*[@Name= 'Path...']")).Click();
            driver.FindElement(MobileBy.XPath("/Window/Window/Window/Pane[2]/Pane[3]/ProgressBar/Pane/ToolBar")).Click();

            SetClipboardAndPaste(driver, "//*[@Name= 'Address']", @"D:\EXPORT_IMPORT\New folder");
            //driver.FindElement(By.XPath("//*[@Name= 'Address']")).SendKeys(@"D:\EXPORT_IMPORT\New folder");
            driver.FindElement(By.XPath("//*[@AutomationId= '41477']")).SendKeys(Keys.Return);

            driver.FindElement(By.XPath("//*[@Name= 'Save']")).Click();
            driver.FindElement(By.XPath("//*[@AutomationId= '1' and @Name= 'OK']")).Click();

            /////////Gambit (FLUENT)/////////

            driver.FindElement(By.XPath("//*[@Name= 'Gambit (FLUENT)']")).Click();

            /////////AutoGrid (NUMECA)/////////

            driver.FindElement(By.XPath("//*[@Name= 'AutoGrid (NUMECA)']")).Click();

            /////////3D Components/////////

            driver.FindElement(By.XPath("//*[@Name= '3D Components']")).Click();

            /////////Geometry for External CFD/////////

            driver.FindElement(By.XPath("//*[@Name= 'Geometry for External CFD']")).Click();

            /////////STL/////////

            driver.FindElement(By.XPath("//*[@Name= 'STL']")).Click();

            /////////ProE IBL/////////

            driver.FindElement(By.XPath("//*[@Name= 'ProE IBL']")).Click();

            /////////AxCFD/////////

            driver.FindElement(By.XPath("//*[@Name= 'AxCFD']")).Click();

            /////////esTurbo (Star-CCM+)/////////

            driver.FindElement(By.XPath("//*[@Name= 'esTurbo (Star-CCM+)']")).Click();

            /////////LSD IMP/////////

            driver.FindElement(By.XPath("//*[@Name= 'LSD IMP']")).Click();

            /////////CAM EXPORT/////////

            driver.FindElement(By.XPath("//*[@Name= 'CAM EXPORT']")).Click();

            /////////Z Plane/////////

            driver.FindElement(By.XPath("//*[@Name= 'Z Plane']")).Click();

            /////////ANSYS WorkBench/////////

            driver.FindElement(By.XPath("//*[@Name= 'ANSYS WorkBench']")).Click();

            //driver.Quit();
        }

        private static WindowsDriver ContextWindow()
        {
            var DesktopSession = new AppiumServiceBuilder()
                .UsingPort(4723)
                .Build();
            DesktopSession.Start();

            AppiumOptions DesktopCapabilities = new AppiumOptions();
            DesktopCapabilities.App = "Root";
            DesktopCapabilities.AutomationName = "Windows";

            return new WindowsDriver(DesktopSession, DesktopCapabilities);
        }

        private static void SetClipboardAndPaste(WindowsDriver driver, string xpath, string text)
        {
            System.Windows.Forms.Clipboard.SetText(text);
            var element = driver.FindElement(By.XPath(xpath));
            element.Click();
            element.SendKeys(Keys.Control + "v");
        }
    }
}