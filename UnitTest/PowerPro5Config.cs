using Microsoft.Win32;
using System;
using System.IO;

namespace PP5AutoUITests
{
    internal class PowerPro5Config
    {
        private static bool isInitialized = false;
        internal const string SubPathPattern = "{0}/{1}";
        internal const string SubsubPathPattern = "{0}/{1}/{2}";
        internal const string DefaultReleaseFolder = "C:/Program Files (x86)/Chroma/PowerPro5";
        internal static string ReleaseFolder;
        internal static string ReleaseTIUserPreTestFolder;
        internal static string ReleaseDataFolder;
        internal static string ReleaseBinFolder;
        internal const string SystemCommandFileName = "SystemCommand.csx";
        internal const string WindowsApplicationDriverUrl = "http://127.0.0.1:4723/";
        internal const string MainPanelProcessName = "Chroma.MainPanel";
        internal const string IDEProcessName = "Chroma.PP5IDE";
        internal const string MainPanelWindowName = "Chroma ATE - MainPanel";
        internal const string IDEWindowName = "Chroma ATS IDE";
        internal const string IDE_TIEditorWindowName = "Chroma ATS IDE - [TI Editor]";
        internal const string IDE_TPEditorWindowName = "Chroma ATS IDE - [TP Editor]";
        internal const string IDE_GUIEditorWindowName = "Chroma ATS IDE - [GUI Editor - Demo]";
        internal const string IDE_ManagementWindowName = "Chroma ATS IDE - [Management]";
        internal const string LoginWindowName = "Login";
        internal const string LoginWindowAutomationID = "LoginWindow";
        internal const string MainPanelAutomationID = "winMainPanel";
        internal static string PP5WindowHandleHex;
        internal static string PP5WindowSessionType;
        internal static string SystemCommand = "SystemCommand";

        internal static void Initialize()
        {
            if (isInitialized)
                return;

            // 嘗試從登錄檔讀取實際的PP5安裝路徑
            ReleaseFolder = GetPP5InstallPathFromRegistry();
            if (string.IsNullOrEmpty(ReleaseFolder))
            {
                ReleaseFolder = DefaultReleaseFolder;
            }

            // 根據安裝路徑設定其他相關路徑
            ReleaseTIUserPreTestFolder = Path.Combine(ReleaseFolder, "TestItem", "UserDefined", "TI", "PreTest");
            ReleaseDataFolder = Path.Combine(ReleaseFolder, "Data");
            ReleaseBinFolder = Path.Combine(ReleaseFolder, "Bin");

            isInitialized = true;
        }

        private static string GetPP5InstallPathFromRegistry()
        {
            try
            {
                // 假設登錄檔中的路徑位於 HKEY_LOCAL_MACHINE\SOFTWARE\Chroma\PowerPro5\Directory 的 ROOT值
                using (RegistryKey key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Chroma\PowerPro5\Directory"))
                {
                    if (key != null)
                    {
                        object installPath = key.GetValue("ROOT");
                        if (installPath != null)
                        {
                            return installPath.ToString();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // 可在此記錄例外資訊，若需要的話
                Console.WriteLine("讀取登錄檔失敗: " + ex.Message);
            }
            return null;
        }
    }
}