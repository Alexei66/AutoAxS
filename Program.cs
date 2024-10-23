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

            var fileName = "Mark15_inducer";

            driver.Manage().Window.Maximize();

            driver.FindElement(By.XPath("//*[@Name= ' File    ']")).Click();
            driver.FindElement(By.XPath("//*[@Name= 'Open...\tCtrl+O']")).Click();

            SetClipboardAndPaste(driver, "//*[@AutomationId='1001']", @"C:\Users\IAB\Desktop\TS_Export");

            SetClipboardAndPaste(driver, "//*[@AutomationId= '1148']", $@"{fileName}.axx");

            driver.FindElement(By.XPath("//*[@Name= 'Viewer']")).Click();

            //driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10); // задержка 10 сек

            //WebDriverWait Desktopwait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            //Desktopwait.Until(pred => AddWatcherElement.Displayed);                           // задержка пока что-то не откроется

            //Thread.Sleep(2000);

            // TODO добавить кнопку с включением 3D вида

            /////////TurboGrid (CFX, FLUENT)/////////

            FindExportButton(driver);

            driver.FindElement(By.XPath("//*[@Name= 'TurboGrid (CFX, FLUENT)']")).Click();
            driver.FindElement(By.XPath("//*[@Name= 'Path...']")).Click();

            SetClipboardAndPaste(driver, "/Window/Window/Window/Pane[2]/Pane[3]/ProgressBar/Pane/ToolBar", @"D:\EXPORT_IMPORT\New folder\New folder");

            driver.FindElement(By.XPath("//*[@Name= 'Save']")).Click();
            driver.FindElement(By.XPath("//*[@AutomationId= '1' and @Name= 'OK']")).Click();

            /////////Gambit (FLUENT)/////////
            FindExportButton(driver);

            driver.FindElement(By.XPath("//*[@Name= 'Gambit (FLUENT)']")).Click();
            driver.FindElement(By.XPath("//*[@Name= 'Path...']")).Click();

            SetClipboardAndPaste(driver, "/Window/Window/Window/Pane[2]/Pane[3]/ProgressBar/Pane/ToolBar", @"D:\EXPORT_IMPORT\New folder\New folder (2)");

            driver.FindElement(By.XPath("//*[@Name= 'Save']")).Click();
            driver.FindElement(By.XPath("//*[@AutomationId= '1' and @Name= 'OK']")).Click();

            /////////AutoGrid (NUMECA)/////////

            FindExportButton(driver);
            driver.FindElement(By.XPath("//*[@Name= 'AutoGrid (NUMECA)']")).Click();
            driver.FindElement(By.XPath("//*[@Name= 'Path...']")).Click();

            SetClipboardAndPaste(driver, "/Window/Window/Window/Pane[2]/Pane[3]/ProgressBar/Pane/ToolBar", @"D:\EXPORT_IMPORT\New folder\New folder (3)");

            driver.FindElement(By.XPath("//*[@Name= 'Save']")).Click();
            driver.FindElement(By.XPath("//*[@AutomationId= '1' and @Name= 'OK']")).Click();

            /////////3D Components/////////

            Exp3DComponents(driver, @"D:\EXPORT_IMPORT\New folder\New folder (4)", "*.igs (Trimmed surface)");

            Exp3DComponents(driver, @"D:\EXPORT_IMPORT\New folder\New folder (5)", "*.igs (BREP representation)");

            Exp3DComponents(driver, @"D:\EXPORT_IMPORT\New folder\New folder (6)", "*.step");

            Exp3DComponents(driver, @"D:\EXPORT_IMPORT\New folder\New folder (7)", "*.stl");

            /////////Geometry for External CFD/////////

            ExpExternalCFD(driver, @"D:\EXPORT_IMPORT\New folder\New folder (8)", "*.igs (Trimmed surface)");
            ExpExternalCFD(driver, @"D:\EXPORT_IMPORT\New folder\New folder (9)", "*.igs (BREP representation)");
            ExpExternalCFD(driver, @"D:\EXPORT_IMPORT\New folder\New folder (10)", "*.step");
            ExpExternalCFD(driver, @"D:\EXPORT_IMPORT\New folder\New folder (11)", "*.stL");
            ExpExternalCFD(driver, @"D:\EXPORT_IMPORT\New folder\New folder (12)", "*.xml(NDF file)");

            /////////STL/////////
            FindExportButton(driver);
            driver.FindElement(By.XPath("//*[@Name= 'STL']")).Click();
            driver.FindElement(By.XPath("//*[@Name= 'Path...']")).Click();

            SetClipboardAndPaste(driver, "/Window/Window/Window/Pane[2]/Pane[3]/ProgressBar/Pane/ToolBar", @"D:\EXPORT_IMPORT\New folder\New folder (13)");

            driver.FindElement(By.XPath("//*[@Name= 'Save']")).Click();
            driver.FindElement(By.XPath("//*[@AutomationId= '1' and @Name= 'OK']")).Click();

            /////////ProE IBL/////////
            FindExportButton(driver);
            driver.FindElement(By.XPath("//*[@Name= 'ProE IBL']")).Click();
            driver.FindElement(By.XPath("//*[@Name= 'Path...']")).Click();

            SetClipboardAndPaste(driver, "/Window/Window/Window/Pane[2]/Pane[3]/ProgressBar/Pane/ToolBar", @"D:\EXPORT_IMPORT\New folder\New folder (14)");

            driver.FindElement(By.XPath("//*[@Name= 'Save']")).Click();
            driver.FindElement(By.XPath("//*[@AutomationId= '1' and @Name= 'OK']")).Click();

            /////////AxCFD/////////
            FindExportButton(driver);
            driver.FindElement(By.XPath("//*[@Name= 'AxCFD']")).Click();

            SetClipboardAndPaste(driver, "/Window/Window/Pane[2]/Pane[3]/ProgressBar/Pane/ToolBar", @"D:\EXPORT_IMPORT\New folder\New folder (15)");

            SetClipboardAndPaste(driver, "//*[@Name= 'File name:' and @AutomationId= '1001']", fileName);

            /////////esTurbo (Star-CCM+)/////////
            FindExportButton(driver);
            driver.FindElement(By.XPath("//*[@Name= 'esTurbo (Star-CCM+)']")).Click();

            SetClipboardAndPaste(driver, "/Window/Window/Pane[2]/Pane[3]/ProgressBar/Pane/ToolBar", @"D:\EXPORT_IMPORT\New folder\New folder (16)");

            driver.FindElement(By.XPath("//*[@Name= 'Select Folder']")).Click();

            /////////Surface data/////////
            FindExportButton(driver);
            driver.FindElement(By.XPath("//*[@Name= 'Surface data ']")).Click();

            SetClipboardAndPaste(driver, "/Window/Window/Pane[2]/Pane[3]/ProgressBar/Pane/ToolBar", @"D:\EXPORT_IMPORT\New folder\New folder (17)");

            driver.FindElement(By.XPath("//*[@Name= 'Select Folder']")).Click();

            /////////LSD IMP/////////
            FindExportButton(driver);
            driver.FindElement(By.XPath("//*[@Name= 'LSD IMP']")).Click();

            SetClipboardAndPaste(driver, "/Window/Window/Pane[2]/Pane[3]/ProgressBar/Pane/ToolBar", @"D:\EXPORT_IMPORT\New folder\New folder (18)");

            driver.FindElement(By.XPath("//*[@Name= 'Select Folder']")).Click();

            /////////CAM EXPORT/////////
            FindExportButton(driver);
            driver.FindElement(By.XPath("//*[@Name= 'CAM EXPORT']")).Click();
            driver.FindElement(By.XPath("//*[@Name= 'Path...']")).Click();

            SetClipboardAndPaste(driver, "/Window/Window/Window/Pane[2]/Pane[3]/ProgressBar/Pane/ToolBar", @"D:\EXPORT_IMPORT\New folder\New folder (19)");

            driver.FindElement(By.XPath("//*[@Name= 'Save']")).Click();
            driver.FindElement(By.XPath("//*[@AutomationId= '1' and @Name= 'OK']")).Click();

            /////////Z Plane/////////
            FindExportButton(driver);
            driver.FindElement(By.XPath("//*[@Name= 'Z Plane']")).Click();
            driver.FindElement(By.XPath("//*[@Name= 'Path...']")).Click();

            SetClipboardAndPaste(driver, "/Window/Window/Window/Pane[2]/Pane[3]/ProgressBar/Pane/ToolBar", @"D:\EXPORT_IMPORT\New folder\New folder (20)");

            driver.FindElement(By.XPath("//*[@Name= 'Save']")).Click();
            driver.FindElement(By.XPath("//*[@AutomationId= '1' and @Name= 'OK']")).Click();

            /////////ANSYS WorkBench/////////
            FindExportButton(driver);
            driver.FindElement(By.XPath("//*[@Name= 'ANSYS WorkBench']")).Click();

            SetClipboardAndPaste(driver, "/Window/Window/Pane[2]/Pane[3]/ProgressBar/Pane/ToolBar", @"D:\EXPORT_IMPORT\New folder\New folder (21)");

            driver.FindElement(By.XPath("//*[@Name= 'Save']")).Click();

            driver.Quit();
        }

        private static void Exp3DComponents(WindowsDriver driver, string savePath, string fileType)
        {
            FindExportButton(driver);

            driver.FindElement(By.XPath("//*[@Name= '3D Components']")).Click();
            driver.FindElement(By.XPath("//*[@Name= 'Path...']")).Click();

            SetClipboardAndPaste(driver, "/Window/Window/Window/Pane[2]/Pane[3]/ProgressBar/Pane/ToolBar", savePath);

            driver.FindElement(By.XPath("//*[@Name= 'Save']")).Click();
            driver.FindElement(MobileBy.XPath("/Window/Window/ComboBox[2]/Button")).Click();

            var DesktopSession = new AppiumServiceBuilder()
                .UsingPort(4723)
                .Build();
            DesktopSession.Start();
            AppiumOptions DesktopCapabilities = new AppiumOptions();
            DesktopCapabilities.App = "Root";
            DesktopCapabilities.AutomationName = "Windows";
            WindowsDriver driver2 = new WindowsDriver(DesktopSession, DesktopCapabilities);

            var addWatcherElement = driver2.FindElement(By.XPath($"//*[@Name= '{fileType}']"));
            addWatcherElement.Click();

            driver.FindElement(By.XPath("//*[@AutomationId= '1' and @Name= 'OK']")).Click();
        }

        private static void ExpExternalCFD(WindowsDriver driver, string savePath, string fileType)
        {
            FindExportButton(driver);

            driver.FindElement(By.XPath("//*[@Name= 'Geometry for External CFD']")).Click();
            driver.FindElement(By.XPath("//*[@Name= 'Path...']")).Click();

            SetClipboardAndPaste(driver, "/Window/Window/Window/Pane[2]/Pane[3]/ProgressBar/Pane/ToolBar", savePath);

            driver.FindElement(By.XPath("//*[@Name= 'Save']")).Click();
            driver.FindElement(MobileBy.XPath("/Window/Window/ComboBox[1]/Button")).Click();

            var DesktopSession = new AppiumServiceBuilder()
                .UsingPort(4723)
                .Build();
            DesktopSession.Start();
            AppiumOptions DesktopCapabilities = new AppiumOptions();
            DesktopCapabilities.App = "Root";
            DesktopCapabilities.AutomationName = "Windows";
            WindowsDriver driver2 = new WindowsDriver(DesktopSession, DesktopCapabilities);

            var addWatcherElement = driver2.FindElement(By.XPath($"//*[@Name= '{fileType}']"));
            addWatcherElement.Click();

            driver.FindElement(By.XPath("//*[@AutomationId= '1' and @Name= 'OK']")).Click();
        }

        private static void FindExportButton(WindowsDriver driver)
        {
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

            var addWatcherElement = driver2.FindElement(By.XPath("//*[@Name= 'Export ...']"));
            addWatcherElement.Click();
        }

        private static void SetClipboardAndPaste(WindowsDriver driver, string xpath, string text)
        {
            System.Windows.Forms.Clipboard.SetText(text);
            var element = driver.FindElement(By.XPath(xpath));
            element.Click();
            element.SendKeys(Keys.Enter + Keys.Control + "v" + Keys.Enter);
        }
    }
}