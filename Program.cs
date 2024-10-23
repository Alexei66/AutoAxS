using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
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

            var appiumLocalService = new OpenQA.Selenium.Appium.Service.AppiumServiceBuilder()
                .UsingPort(4723)
                .WithStartUpTimeOut(TimeSpan.FromSeconds(30))
                .Build();
            appiumLocalService.Start();

            AppiumOptions desiredCapabilities = new AppiumOptions();
            desiredCapabilities.App = @"C:\AxStream\Signed_Release\AxStream_x64 v3.10.6.8321\AxSTREAM.exe";
            desiredCapabilities.AutomationName = "Windows";
            desiredCapabilities.AddAdditionalAppiumOption("ms:WaitForAppLaunch", "20");
            driver = new WindowsDriver(appiumLocalService, desiredCapabilities);

            var fileName = "Mark15_inducer";

            driver.Manage().Window.Maximize();

            driver.FindElement(By.XPath("//*[@Name= ' File    ']")).Click();
            driver.FindElement(By.XPath("//*[@Name= 'Open...\tCtrl+O']")).Click();

            ExportActions.SetClipboardAndPaste(driver, "//*[@AutomationId='1001']", @"C:\Users\IAB\Desktop\TS_Export");

            ExportActions.SetClipboardAndPaste(driver, "//*[@AutomationId= '1148']", $@"{fileName}.axx");

            driver.FindElement(By.XPath("//*[@Name= 'Viewer']")).Click();

            ExportActions.TurboGridCFX_FLUENT(driver);
            ExportActions.GambitFLUENT(driver);
            ExportActions.AutoGridNUMECA(driver);
            ExportActions.Components3D(driver);
            ExportActions.GeometryForExternalCFD(driver);
            ExportActions.STL(driver);
            ExportActions.ProE_IBL(driver);
            ExportActions.AxCFD(driver, fileName);
            ExportActions.esTurboStarCCM(driver);
            ExportActions.SurfaceData(driver);
            ExportActions.LSD_IMP(driver);
            ExportActions.CAM_Export(driver);
            ExportActions.Z_Plane(driver);
            ExportActions.ANSYS_WorkBench(driver);

            driver.Quit();
            appiumLocalService.Dispose();

            ////driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10); // задержка 10 сек

            ////WebDriverWait Desktopwait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            ////Desktopwait.Until(pred => AddWatcherElement.Displayed);                           // задержка пока что-то не откроется

            ////Thread.Sleep(2000);

            //// TODO добавить кнопку с включением 3D вида
        }
    }
}