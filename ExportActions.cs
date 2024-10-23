using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Service;
using OpenQA.Selenium.Appium.Windows;

namespace ConsoleApp1
{
    public class ExportActions
    {
        public static void TurboGridCFX_FLUENT(WindowsDriver driver)
        {
            FindExportButton(driver);

            driver.FindElement(By.XPath("//*[@Name= 'TurboGrid (CFX, FLUENT)']")).Click();
            driver.FindElement(By.XPath("//*[@Name= 'Path...']")).Click();

            SetClipboardAndPaste(driver, "/Window/Window/Window/Pane[2]/Pane[3]/ProgressBar/Pane/ToolBar", @"D:\EXPORT_IMPORT\New folder\New folder");
            SaveExport(driver);
        }

        public static void GambitFLUENT(WindowsDriver driver)
        {
            FindExportButton(driver);

            driver.FindElement(By.XPath("//*[@Name= 'Gambit (FLUENT)']")).Click();
            driver.FindElement(By.XPath("//*[@Name= 'Path...']")).Click();

            SetClipboardAndPaste(driver, "/Window/Window/Window/Pane[2]/Pane[3]/ProgressBar/Pane/ToolBar", @"D:\EXPORT_IMPORT\New folder\New folder (2)");
            SaveExport(driver);
        }

        public static void AutoGridNUMECA(WindowsDriver driver)
        {
            FindExportButton(driver);

            driver.FindElement(By.XPath("//*[@Name= 'AutoGrid (NUMECA)']")).Click();
            driver.FindElement(By.XPath("//*[@Name= 'Path...']")).Click();

            SetClipboardAndPaste(driver, "/Window/Window/Window/Pane[2]/Pane[3]/ProgressBar/Pane/ToolBar", @"D:\EXPORT_IMPORT\New folder\New folder (3)");
            SaveExport(driver);
        }

        public static void Components3D(WindowsDriver driver)
        {
            Components3D(driver, @"D:\EXPORT_IMPORT\New folder\New folder (4)", "*.igs (Trimmed surface)");
            Components3D(driver, @"D:\EXPORT_IMPORT\New folder\New folder (5)", "*.igs (BREP representation)");
            Components3D(driver, @"D:\EXPORT_IMPORT\New folder\New folder (6)", "*.step");
            Components3D(driver, @"D:\EXPORT_IMPORT\New folder\New folder (7)", "*.stl");
        }

        public static void GeometryForExternalCFD(WindowsDriver driver)
        {
            GeometryForExternalCFD(driver, @"D:\EXPORT_IMPORT\New folder\New folder (9)", "*.igs (BREP representation)");
            GeometryForExternalCFD(driver, @"D:\EXPORT_IMPORT\New folder\New folder (8)", "*.igs (Trimmed surface)");
            GeometryForExternalCFD(driver, @"D:\EXPORT_IMPORT\New folder\New folder (10)", "*.step");
            GeometryForExternalCFD(driver, @"D:\EXPORT_IMPORT\New folder\New folder (11)", "*.stL");
            GeometryForExternalCFD(driver, @"D:\EXPORT_IMPORT\New folder\New folder (12)", "*.xml(NDF file)");
        }

        public static void STL(WindowsDriver driver)
        {
            FindExportButton(driver);

            driver.FindElement(By.XPath("//*[@Name= 'STL']")).Click();
            driver.FindElement(By.XPath("//*[@Name= 'Path...']")).Click();

            SetClipboardAndPaste(driver, "/Window/Window/Window/Pane[2]/Pane[3]/ProgressBar/Pane/ToolBar", @"D:\EXPORT_IMPORT\New folder\New folder (13)");
            SaveExport(driver);
        }

        public static void ProE_IBL(WindowsDriver driver)
        {
            FindExportButton(driver);

            driver.FindElement(By.XPath("//*[@Name= 'ProE IBL']")).Click();
            driver.FindElement(By.XPath("//*[@Name= 'Path...']")).Click();

            SetClipboardAndPaste(driver, "/Window/Window/Window/Pane[2]/Pane[3]/ProgressBar/Pane/ToolBar", @"D:\EXPORT_IMPORT\New folder\New folder (14)");
            SaveExport(driver);
        }

        public static void AxCFD(WindowsDriver driver, string fileName)
        {
            FindExportButton(driver);

            driver.FindElement(By.XPath("//*[@Name= 'AxCFD']")).Click();

            SetClipboardAndPaste(driver, "/Window/Window/Pane[2]/Pane[3]/ProgressBar/Pane/ToolBar", @"D:\EXPORT_IMPORT\New folder\New folder (15)");

            SetClipboardAndPaste(driver, "//*[@Name= 'File name:' and @AutomationId= '1001']", fileName);
        }

        public static void esTurboStarCCM(WindowsDriver driver)
        {
            FindExportButton(driver);

            driver.FindElement(By.XPath("//*[@Name= 'esTurbo (Star-CCM+)']")).Click();
            SetClipboardAndPaste(driver, "/Window/Window/Pane[2]/Pane[3]/ProgressBar/Pane/ToolBar", @"D:\EXPORT_IMPORT\New folder\New folder (16)");
            driver.FindElement(By.XPath("//*[@Name= 'Select Folder']")).Click();
        }

        public static void SurfaceData(WindowsDriver driver)
        {
            FindExportButton(driver);

            driver.FindElement(By.XPath("//*[@Name= 'Surface data ']")).Click();
            SetClipboardAndPaste(driver, "/Window/Window/Pane[2]/Pane[3]/ProgressBar/Pane/ToolBar", @"D:\EXPORT_IMPORT\New folder\New folder (17)");
            driver.FindElement(By.XPath("//*[@Name= 'Select Folder']")).Click();
        }

        public static void LSD_IMP(WindowsDriver driver)
        {
            FindExportButton(driver);

            driver.FindElement(By.XPath("//*[@Name= 'LSD IMP']")).Click();
            SetClipboardAndPaste(driver, "/Window/Window/Pane[2]/Pane[3]/ProgressBar/Pane/ToolBar", @"D:\EXPORT_IMPORT\New folder\New folder (18)");
            driver.FindElement(By.XPath("//*[@Name= 'Select Folder']")).Click();
        }

        public static void CAM_Export(WindowsDriver driver)
        {
            FindExportButton(driver);

            driver.FindElement(By.XPath("//*[@Name= 'CAM EXPORT']")).Click();
            driver.FindElement(By.XPath("//*[@Name= 'Path...']")).Click();

            SetClipboardAndPaste(driver, "/Window/Window/Window/Pane[2]/Pane[3]/ProgressBar/Pane/ToolBar", @"D:\EXPORT_IMPORT\New folder\New folder (19)");
            SaveExport(driver);
        }

        public static void Z_Plane(WindowsDriver driver)
        {
            FindExportButton(driver);

            driver.FindElement(By.XPath("//*[@Name= 'Z Plane']")).Click();
            driver.FindElement(By.XPath("//*[@Name= 'Path...']")).Click();

            SetClipboardAndPaste(driver, "/Window/Window/Window/Pane[2]/Pane[3]/ProgressBar/Pane/ToolBar", @"D:\EXPORT_IMPORT\New folder\New folder (20)");
            SaveExport(driver);
        }

        public static void ANSYS_WorkBench(WindowsDriver driver)
        {
            FindExportButton(driver);

            driver.FindElement(By.XPath("//*[@Name= 'ANSYS WorkBench']")).Click();
            SetClipboardAndPaste(driver, "/Window/Window/Pane[2]/Pane[3]/ProgressBar/Pane/ToolBar", @"D:\EXPORT_IMPORT\New folder\New folder (21)");
            driver.FindElement(By.XPath("//*[@Name= 'Save']")).Click();
        }

        // Метод для нахождения кнопки экспорта
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

        private static void Components3D(WindowsDriver driver, string savePath, string fileType)
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

        private static void GeometryForExternalCFD(WindowsDriver driver, string savePath, string fileType)
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

        public static void SetClipboardAndPaste(WindowsDriver driver, string xpath, string text)
        {
            System.Windows.Forms.Clipboard.SetText(text);
            var element = driver.FindElement(By.XPath(xpath));
            element.Click();
            element.SendKeys(Keys.Enter + Keys.Control + "v" + Keys.Enter);
        }

        // Метод для сохранения
        private static void SaveExport(WindowsDriver driver)
        {
            driver.FindElement(By.XPath("//*[@Name= 'Save']")).Click();
            driver.FindElement(By.XPath("//*[@AutomationId= '1' and @Name= 'OK']")).Click();
        }
    }
}