using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium.Appium.Windows;

namespace PP5AutoUITests.Session
{
    public class PP5IDEDriver : DriverBase, ISessionDriver
    {
        public PP5IDEDriver()
        {
            SessionName = PowerPro5Config.IDEProcessName;
            WindowName = PowerPro5Config.IDE_TIEditorWindowName;
            sessionType = SessionType.PP5IDE;
            //targetAppPath = @"C:\Users\adam.chen\Desktop\Debug\bin\Chroma.PP5IDE.exe";
            //targetAppWorkingDir = @"C:\Users\adam.chen\Desktop\Debug\bin\";
            //base.targetAppWorkingDir = @"C:\Program Files (x86)\Chroma\PowerPro5\Bin\";
            targetAppPath = string.Format(PowerPro5Config.SubPathPattern, PowerPro5Config.ReleaseBinFolder, PowerPro5Config.IDEProcessName + ".exe");
            //targetAppWorkingDir = PowerPro5Config.ReleaseBinFolder;
        }

        public PP5Driver CreateNewDriver()
        {
            OpenQA.Selenium.Appium.AppiumOptions appCapabilities = new OpenQA.Selenium.Appium.AppiumOptions();
            appCapabilities.AddAdditionalCapability("app", "Root");
            appCapabilities.AddAdditionalCapability("deviceName", DeviceName);
            currentDriver = new PP5Driver(new Uri(WindowsApplicationDriverUrl), appCapabilities, TimeSpan.FromSeconds(180), SessionType.Desktop);

            currentDriver.ShouldNotBeNull($"PP5Driver: {sessionType.ToString()} created failed!");

            //PowerPro5Config.PP5WindowSessionType = sessionType.ToString();

            // Set implicit timeout to 1.5 seconds to make element search to retry every 500 ms for at most three times
            //currentDriver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1.5);
            return currentDriver;
        }

        public override PP5Driver CreateDriver()
        {
            processes = Process.GetProcessesByName(Path.GetFileNameWithoutExtension(targetAppPath));
            if (processes.Length == 0)
            {
                CreateNewDriver();
            }
            else
            {
                AttachExistingDriver();
            }
            return currentDriver;
        }

        //public new void AttachExistingDriver()
        //{
        //    base.AttachExistingDriver();
        //}
    }
}
