//#define WRITE_LOG
//using System.Timers;
using System.Collections;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using Castle.Core.Internal;
using Chroma.UnitTest.Common;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static PP5AutoUITests.AutoUIActionHelper;
using static PP5AutoUITests.AutoUIExtension;
using static PP5AutoUITests.ControlElementExtension;
using static PP5AutoUITests.ThreadHelper;
using static PP5AutoUITests.DllHelper;
using static PP5AutoUITests.FileProcessingExtension;
using static PP5AutoUITests.TIEditorHelper;
using Keys = OpenQA.Selenium.Keys;
using static PP5AutoUITests.TestCases_TIEditor_EnumCreatorDialog;
using System.Collections.Generic;
using System.Timers;
using Microsoft.Win32;
using OpenQA.Selenium.Interactions;
using System;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Appium.Interfaces;
using OpenQA.Selenium.Opera;
using Chroma.UnitTest.Common.AutoUI;
using System.Data;
using Castle.Components.DictionaryAdapter.Xml;
using System.Linq;
using System.Configuration;
using System.Runtime.CompilerServices;
using System.Windows.Automation;
using static System.Net.Mime.MediaTypeNames;
using Castle.DynamicProxy.Generators.Emitters.SimpleAST;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting.Logging;
using System.Collections.Concurrent;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TaskbarClock;
using OpenQA.Selenium.Support.PageObjects;
using System.IO;
using System.Security.AccessControl;
using System.Drawing;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Status;
using Newtonsoft.Json.Linq;
using System.Text;
using System.Data.Common;
using System.Drawing.Imaging;

namespace PP5AutoUITests
{
    [TestClass]
    public class TestBase : UTCommon.TestBase
    {
        //private const string WinAppDriverPath = @"C:\Program Files (x86)\Windows Application Driver\WinAppDriver.exe";
        //public string TestResultOutputDirectory = Directory.GetCurrentDirectory();
        public static readonly string TestResultOutputDirectory = CurrWorkingDirectory;

        public static class TestResultCollection
        {
            public static Dictionary<ITestMethod, TestResult[]> Results { get; set; } = new Dictionary<ITestMethod, TestResult[]>();
        }

        public class MyTestMethodAttribute : TestMethodAttribute
        {

            public MyTestMethodAttribute(string displayName) : base(displayName) { }

            public override TestResult[] Execute(ITestMethod testMethod)
            {
                TestResult[] results = base.Execute(testMethod);

                TestResultCollection.Results.Add(testMethod, results);

                return results;
            }
        }

        /// <summary>The RepeatAttribute is used on a test method to specify that it should be executed multiple times. If any repetition fails, the remaining ones are not run and a failure is reported.</summary>
        /// <param name="repeatCount">int representing the repitition count.</param>"
        [AttributeUsage(AttributeTargets.Method, AllowMultiple = false)]
        public class RepeatAttribute : Attribute
        {
            private const int MIN_REPEAT_COUNT = 1;
            private const int MAX_REPEAT_COUNT = 50;

            public RepeatAttribute(
                int repeatCount)
            {
                if (repeatCount < MIN_REPEAT_COUNT || MAX_REPEAT_COUNT < repeatCount)
                {
                    repeatCount = MIN_REPEAT_COUNT;
                }

                Value = repeatCount;
            }

            public int Value { get; }
        }


        public static Executor AutoUIExecutor;
        //public static WindowsDriver<WindowsElement> MainPanel_Session, PP5IDE_Session, Desktop_Session;
        //static WindowsDriver<WindowsElement> currentDriver;
        public static PP5Driver CurrentDriver => AutoUIExecutor.GetCurrentDriver();
        public static PP5Driver CurrentDriverNotGeneric => AutoUIExecutor.GetCurrentDriverNotGeneric();
        //{
        //    get
        //    {
        //        currentDriver = AutoUIExecutor.GetCurrentDriver();
        //        return currentDriver;
        //    }
        //}

        //public AutoUIAction uIActionMainPanel, uIActionPP5IDE;
        //public static AutoUIAction uIAction;
        public static Process winAppDriverProcess;

        protected static bool IsPP5IDEWindowStale => ExpectedConditions.StalenessOf(_PP5IDEWindow)(CurrentDriver);

        //private bool isIDEWindowPresent = false;
        public static bool IsIDEWindowPresent
        {
            get
            {
                return _PP5IDEWindow != null && !IsPP5IDEWindowStale;
            }
        }

        static IElement _PP5IDEWindow;
        public static IElement PP5IDEWindow
        {
            get
            {
                if (!IsIDEWindowPresent)
                    _PP5IDEWindow = GetPP5Window();
                return _PP5IDEWindow;
            }
        }

        public static DriverLogger driverLogger;


        //public MemoryCache<(int?, string), IWebElement> dataTableConditionCache;
        //public MemoryCache<(int?, string), IWebElement> dataTableResultCache;
        //public MemoryCache<(int?, string), IWebElement> dataTableTempCache;
        //public MemoryCache<(int?, string), IWebElement> dataTableGlobalCache;
        //public MemoryCache<(int?, string), IWebElement> dataTableDefectCodeCache;
        //public MemoryCache<(int?, string), IWebElement> dataTableTestItemsCache;
        //public MemoryCache<(int?, string), IWebElement> dataTableTestCmdParamCache;
        //public MemoryCache<(int?, string), IWebElement> dataTableAllTestItemsCache;

        public MemoryCache<string, IElement> dataTablesCache;

        //public static WindowsElement menuCache;
        public MemoryCache<string, IElement> CommandsMapCache;
        public string[] moduleNames = { "TI Editor", "TP Editor", "Report", "On-Line Control",
                                        "Management", "Hardware Configuration", "GUI Editor", "Execution" };

        internal static CaptureAppScreenshotTimer _scrnShotTimer;
        //public static Dictionary<string, bool> cmdGroupSourceTypeDict;
        internal static OrderedDictionary<string, CommandGroupData> cmdGroupDataDict;
        internal static TaskManager taskManager = new TaskManager();
        internal static tifiTestData tiTestData = new tifiTestData();
        internal static FileSystemWatcherWrapper fileWatcher;

        static object fileWatcherLock = new object();
        public delegate void LogMessageDelegate(string message);
        public delegate string GetTIFilePathDelegate(string TIName, TestItemSourceType sourceType, TestItemType itemType, TestItemRunType runType);
        public static void LogMessage(string message) => Logger.LogMessage(message);
        public string TIFilePath = null;
        public static HashSet<string> TIFilePaths = new HashSet<string>();
        public static string TIName;
        //public static bool isSubOrThreadTI = false;

        [AssemblyInitialize]
        public static void BeforeClass(TestContext tc)
        {
            TestEnvSetup();     // Test environment setup before testing started

            AutoUIExecutor = Executor.GetInstance();
            driverLogger = new DriverLogger(AutoUIExecutor);
            PP5LogIn();
            _PP5IDEWindow = null;

            //winAppDriverProcess = AutoUIExecutor.WinAppDriverProcessElement;  // Get winAppDriver session
            //PP5IDE_Session = AutoUIExecutor.DesktopSessionElement;            // get PP5 IDE session
            //MainPanel_Session = AutoUIExecutor.PP5SessionElement;                    // get PP5 main panel session    
            //Desktop_Session = AutoUIExecutor.MyDesktopSessionElement;                    // get desktop session   
        }

        [AssemblyCleanup]
        public static void AfterClass()
        {
            //EnableMouse(true);   // Enable mouse usage after all testing are finished
            //ActivateDevice(true, DeviceType.Mouse);

            // close winAppDriver
            //AutoUIExecutor.WinAppDriverProcess.Dispose();
            //AutoUIExecutor.WinAppDriverProcess.Close();
            AutoUIExecutor.WinAppDriverProcess.Kill();

            //Console.WriteLine("After all tests");
            GC.Collect();
            CurrentDriver.ReleaseComObject();
            AutoUIExecutor.ReleaseComObject();
        }

        //[ClassInitialize(InheritanceBehavior.BeforeEachDerivedClass)]
        //public static void Setup(TestContext context)
        //{

        //}

        //[ClassCleanup(InheritanceBehavior.BeforeEachDerivedClass)]
        //public static void Cleanup(TestContext context)
        //{

        //}

        [TestInitialize]
        public void TestMethodSetup()
        {
            // Element cache initialize
            CommandsMapCache = new MemoryCache<string, IElement>();
            dataTablesCache = new MemoryCache<string, IElement>();

            //dataTableConditionCache = new MemoryCache<(int?, string), IWebElement>();
            //dataTableResultCache = new MemoryCache<(int?, string), IWebElement>();
            //dataTableTempCache = new MemoryCache<(int?, string), IWebElement>();
            //dataTableGlobalCache = new MemoryCache<(int?, string), IWebElement>();
            //dataTableDefectCodeCache = new MemoryCache<(int?, string), IWebElement>();
            //dataTableTestItemsCache = new MemoryCache<(int?, string), IWebElement>();
            //dataTableTestCmdParamCache = new MemoryCache<(int?, string), IWebElement>();
            //dataTableAllTestItemsCache = new MemoryCache<(int?, string), IWebElement>();
        }

        [TestCleanup]
        public void AfterTest()
        {
            //Console.WriteLine("After each test");
        }

        [TestMethod]
        public void TryTestBase()
        {
            //Console.WriteLine("TestBase try method executed!");
            OpenNewTIEditorWindow();
        }

        public static void TestEnvSetup()
        {
#if !DEBUG
            //EnableMouse(false); // Disable mouse usage before testing
            //ActivateDevice(false, DeviceType.Mouse);
#endif
            DisableMemoryMonitorWindow(); // Disable Memory Monitor that shows memory warning window before starting the driver
            JsonUpdateProperty(filePath: $"{PowerPro5Config.ReleaseDataFolder}//SystemSetup.ssx",   // Turn off TI&TP autosave
                               nodePath: "Datas",
                               propertyName: "AutoBackupInterval",
                               newValue: 0);

            //LoadCommandGroupInfos();
            //await LoadCommandGroupTask();
            int taskId = taskManager.StartNewTask("LoadCommandGroup", LoadCommandGroup);
            SharedSetting.forceRefreshPP5Window = false;

            fileWatcher = new FileSystemWatcherWrapper();                                                                                   // 创建 FileSystemWatcherWrapper 实例
            //foreach (string path in GetAllTIFileFolders())
            //    fileWatcher.CreateFileWatcher(path, NotifyFilters.FileName, "tix");                                                         // 创建文件监视器: .tix (watch all TIs in all TI related folders)
            fileWatcher.CreateFileWatcher(PowerPro5Config.ReleaseDataFolder, NotifyFilters.LastWrite | NotifyFilters.CreationTime, "til");  // 创建文件监视器: .til (watch TestItemList.til in System Data folder)
            fileWatcher.FileChanged += FileWatcher_UpdateTiInfos;                                                                           // 订阅 FileChanged 事件
        }

        // 事件处理程序
        private static void FileWatcher_UpdateTiInfos(object sender, FileSystemEventArgs e)
        {
            //Console.WriteLine($"File Changed: {e.FullPath}, ChangeType: {e.ChangeType}");
            //WaitUntil(() => !IsFileLocked(e.FullPath));

            //Thread.Sleep(6000);
            //string tiName = Path.GetFileNameWithoutExtension(e.FullPath);
            WaitUntilTIFinishedSaving();
            TIFilePaths.Add(tiTestData.GetFullFileName(TIName));
            WaitUntil(() => !IsFileLocked(e.FullPath));
            //Logger.LogMessage("Loading cmd group names: \"Sub TI\"");
            //Logger.LogMessage("Loading cmd group names: \"Thread TI\"");
            if (tiTestData.IsSpecialTIFileName(TIName))
            {
                LoadCommandGroupCommandNames("Thread TI");
                LoadCommandGroupCommandNames("Sub TI");
            }

            //taskManager.StartNewTask(() => LogMessage("Loading cmd group names: \"Sub TI\""));
            //taskManager.StartNewTask(() => LogMessage("Loading cmd group names: \"Thread TI\""));
            //taskManager.StartNewTask(() => LoadCommandGroupCommandNames("Sub TI"));
            //taskManager.StartNewTask(() => LoadCommandGroupCommandNames("Thread TI"));
        }

        //const int ERROR_SHARING_VIOLATION = 32;
        //const int ERROR_LOCK_VIOLATION = 33;
        //private static bool IsFileLocked(string file)
        //{
        //    //check that problem is not in destination file
        //    if (File.Exists(file) == true)
        //    {
        //        FileStream stream = null;
        //        try
        //        {
        //            stream = File.Open(file, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        //        }
        //        catch (Exception ex2)
        //        {
        //            //_log.WriteLog(ex2, "Error in checking whether file is locked " + file);
        //            int errorCode = Marshal.GetHRForException(ex2) & ((1 << 16) - 1);
        //            if ((ex2 is IOException) && (errorCode == ERROR_SHARING_VIOLATION || errorCode == ERROR_LOCK_VIOLATION))
        //            {
        //                return true;
        //            }
        //        }
        //        finally
        //        {
        //            if (stream != null)
        //                stream.Close();
        //        }
        //    }
        //    return false;
        //}

        internal TaskInfo GetTaskInfo(int taskId)
        {
            return taskManager.GetTaskInfo(taskId);
        }

        internal int GetFirstTaskId()
        {
            return taskManager.GetAllTaskIds().OrderBy(n => n).First();
        }

        internal int GetLastTaskId()
        {
            return taskManager.GetAllTaskIds().OrderBy(n => n).Last();
        }

        internal bool CheckAllTasksCompleted()
        {
            return taskManager.GetAllTaskIds().TrueForAll(id => GetTaskInfo(id).Status is TaskStatus.Completed);
        }

        public static void LoadCommandGroup()
        {
            LoadCommandGroupInfos();

            foreach (string groupName in GetGroupNames())
                LoadCommandGroupCommandNames(groupName);

            //Logger.LogMessage($"LoadCommandGroup() executed time (ms): {taskManager.GetTaskInfo(taskId).TotalTimeMilliseconds}");
            Logger.LogMessage("LoadCommandGroup() done.");
        }

        //public static Task LoadCommandGroupTask()
        //{
        //    return Task.Run(() =>
        //    {
        //        try
        //        {
        //            LoadCommandGroup();
        //            Logger.LogMessage("Task has completed.");
        //        }
        //        catch (Exception ex)
        //        {
        //            Logger.LogMessage($"Error in LoadCommandGroup: {ex.Message}");
        //            throw; // Re-throw to propagate the error
        //        }
        //    });
        //}

        public static List<string> GetGroupNames()
        {
            if (cmdGroupDataDict.Count == 0)
                LoadCommandGroupInfos();

            return cmdGroupDataDict.Keys.ToList();
        }

        public static int GetGroupID(string groupName)
        {
            if (cmdGroupDataDict.Count == 0)
                LoadCommandGroupInfos();

            return cmdGroupDataDict[groupName].GroupID;
        }

        public static void LoadCommandGroupInfos()
        {
            // Read SystemCommand.csx and find all GroupName and IsDevice values
            JsonGetProperty(filePath: $"{PowerPro5Config.ReleaseDataFolder}//SystemCommand.csx", nodePath: "CommandGroupInfos@IsDevice", out List<bool> IsDeviceBooleanValues);
            JsonGetProperty(filePath: $"{PowerPro5Config.ReleaseDataFolder}//SystemCommand.csx", nodePath: "CommandGroupInfos@GroupID", out List<int> GroupIds);
            JsonGetProperty(filePath: $"{PowerPro5Config.ReleaseDataFolder}//SystemCommand.csx", nodePath: "CommandGroupInfos@GroupName", out List<string> GroupNameKeys);

            var commandGroupDataList = new List<CommandGroupData>();
            for (int i = 0; i < GroupNameKeys.Count; i++)
            {
                commandGroupDataList.Add(new CommandGroupData(GroupIds[i], IsDeviceBooleanValues[i]));
            }

            // Add Thread TI and Sub TI cmd group data manually
            string[] threadAndSubTILabels = new string[] { "Thread TI", "Sub TI" };
            for (int i = 0; i < threadAndSubTILabels.Length; i++)
            {
                commandGroupDataList.Add(new CommandGroupData(30000 + i, false));
                GroupNameKeys.Add(threadAndSubTILabels[i]);
            }

            // Create dictionary of groupname - isdevice (true, false) pairs
            cmdGroupDataDict = CreateOrderedDictionary(GroupNameKeys, commandGroupDataList);
        }

        public static void LoadCommandGroupCommandNames(string groupName)
        {
            List<string> commandNames;
            List<bool> visibilityList = null;

            if (groupName == "Thread TI" || groupName == "Sub TI")
                commandNames = GetThreadTiOrSubTiNames(groupName);
            else
            {
                string commandFilePath = GetCommandFileFullPath();
                //JsonGetProperty(filePath: commandFilePath, nodePath: "CommandGroupInfos@GroupName", out List<string> GroupNameKeys);

                int groupIndex = GetGroupNames().IndexOf(groupName);

                // Read SystemCommand.csx and find all GroupName and IsDevice values
                JsonGetProperty(filePath: commandFilePath, nodePath: $"CommandGroupInfos[{groupIndex}]/Commands@CommandName", out commandNames);
                JsonGetProperty(filePath: commandFilePath, nodePath: $"CommandGroupInfos[{groupIndex}]/Commands@Visible", out visibilityList);
            }

            //JsonGetProperty(filePath: $"{PowerPro5Config.ReleaseDataFolder}//SystemCommand.csx", nodePath: "CommandGroupInfos@GroupName", out List<string> GroupNameKeys);

            //List<CommandGroupData> commandGroupDataList = new List<CommandGroupData>();
            //for (int i = 0; i < GroupNameKeys.Count; i++)
            //{
            //    commandGroupDataList.Add(new CommandGroupData(commandNames[i]));
            //}

            // Create dictionary of groupname - isdevice (true, false) pairs
            //cmdGroupDataDict = CreateDictionary(GroupNameKeys, commandGroupDataList);
            if (groupName == "System, Flow Control")
                commandNames.Remove("Call_TI");

            cmdGroupDataDict[groupName].CommandNames = visibilityList == null
                                                       ? commandNames.OrderBy(n => n).ToList()
                                                       : commandNames.Where((cmd, idx) => visibilityList[idx]).OrderBy(n => n).ToList();
        }


        private List<string> QueryCommandNames(string groupName)
        {
            if (cmdGroupDataDict.AsDictionary().TryGetValue(groupName, out CommandGroupData groupData))
            {
                if (groupData.CommandNames.Count == 0)
                {
                    //taskManager.StartNewTask("LoadCommandGroupCommandNames", () => LoadCommandGroupCommandNames(groupName));
                    LoadCommandGroupCommandNames(groupName);
                    return cmdGroupDataDict[groupName].CommandNames;
                }
                else
                {
                    //if (groupName == "Thread TI" || groupName == "Sub TI")
                    //    LoadCommandGroupCommandNames(groupName);
#if DEBUG
                    //foreach (var commandName in groupData.CommandNames)
                    //    Logger.LogMessage($"In QueryCommandNames method, commandName:{commandName}");
#endif
                    return groupData.CommandNames;
                }
            }
            else
                throw new GroupNameNotExistedException(groupName);
        }

        public string QueryGroupName(string commandName)
        {
            if (!HasCommand(commandName))
                throw new CommandNameNotExistedException(commandName);

            return cmdGroupDataDict.AsDictionary().FirstOrDefault(kvp => kvp.Value.CommandNames.Contains(commandName)).Key;
        }

        public bool HasGroup(string groupName)
        {
            return cmdGroupDataDict.ContainsKey(groupName);
        }

        public bool HasCommand(string groupName, string commandName)
        {
            if (!HasGroup(groupName))
                return false;

            return cmdGroupDataDict[groupName].CommandNames.Contains(commandName);
        }

        public bool HasCommand(string commandName)
        {
            return cmdGroupDataDict.AsDictionary().Count(kvp => kvp.Value.CommandNames.Contains(commandName)) != 0;
        }

        public static bool AddCommandInCGIList(string groupName, string cmdName)
        {
            //if (cmdGroupDataDict.Count == 0)
            //    LoadCommandGroupInfos();

            //var groupList = cmdGroupDataDict.Keys.ToList();

            int CGIListIdx = GetGroupNames().IndexOf(groupName);

            string nodepathToUpdate = $"CommandGroupInfos[{CGIListIdx}]/Commands[0>-1]@CommandName={cmdName}";

            // Create a new node in the cgi list and saved in the SystemCommand.csx file
            bool isJsonCreateNewNodeSuccess = JsonCreateNewNodeInList(filePath: GetCommandFileFullPath(),
                                                                      nodePath: nodepathToUpdate);

            taskManager.StartNewTask("LoadCommandGroupCommandNames", () => LoadCommandGroupCommandNames(groupName));
            //LoadCommandGroupCommandNames(groupName);
            return isJsonCreateNewNodeSuccess;
        }

        public static bool AddEmptyCommandInCGIList(string groupName, string cmdName, bool cmdIsVisible = true, bool isUserDefined = true)
        {
            //if (cmdGroupDataDict.Count == 0)
            //    LoadCommandGroupInfos();

            //var groupList = cmdGroupDataDict.Keys.ToList();

            int CGIListIdx = GetGroupNames().IndexOf(groupName);
            int groupId = GetGroupID(groupName);
            GetCommandSourceType(groupName, out bool IsDevice);

            string nodePart = $"CommandGroupInfos[{CGIListIdx}]/Commands[>-1]";
            string propertyPart = $"CommandName={cmdName},IsDevice={IsDevice},GroupID={groupId},Visible={cmdIsVisible},IsUserDefined={isUserDefined}";
            string nodepathToUpdate = $"{nodePart}@{propertyPart}";

            // Create a new node in the cgi list and saved in the SystemCommand.csx file
            bool isJsonCreateNewNodeSuccess = JsonCreateNewNodeInList(filePath: GetCommandFileFullPath(),
                                                                      nodePath: nodepathToUpdate);

            taskManager.StartNewTask("LoadCommandGroupCommandNames", () => LoadCommandGroupCommandNames(groupName));
            //LoadCommandGroupCommandNames(groupName);
            return isJsonCreateNewNodeSuccess;
        }

        public static List<string> GetThreadTiOrSubTiNames(string ThreadTiOrSubTiLabel)
        {
            //lock (fileWatcherLock)
            //{
            List<string> ThreadTiOrSubTiNames = new List<string>();
            if (ThreadTiOrSubTiLabel == "Thread TI" || ThreadTiOrSubTiLabel == "Sub TI")
            {
                ThreadTiOrSubTiNames.AddRange(tiTestData.GetSpecialTIFileNames(ThreadTiOrSubTiLabel));
            }
            else
                throw new ArgumentException("Input label is not \"Thread TI\"  \"Sub TI\"");

            return ThreadTiOrSubTiNames;
            //}
        }

        public static List<string> GetAllTIFileFolders()
        {
            List<string> TIFilePaths = new List<string>();
            foreach (TestItemSourceType sourceType in Enum.GetValues(typeof(TestItemSourceType)))
            {
                foreach (TestItemRunType runType in Enum.GetValues(typeof(TestItemRunType)))
                {
                    foreach (TestItemType itemType in Enum.GetValues(typeof(TestItemType)))
                    {
                        TIFilePaths.Add(GetTIFileFolder(sourceType, itemType, runType));
                    }
                }
            }
            return TIFilePaths;
        }

        public static string GetTIFileFolder(TestItemSourceType sourceType = TestItemSourceType.UserDefined, TestItemType itemType = TestItemType.TI, TestItemRunType runType = TestItemRunType.UUT)
        {
            string pathItemType = Path.Combine(PowerPro5Config.ReleaseFolder, "TestItem", sourceType.ToString(), itemType.ToString());
            if (itemType == TestItemType.TI)
                return Path.Combine(pathItemType, runType.GetDescription().Replace(" ", ""));
            else
                return Path.Combine(pathItemType);
        }

        public static string GetTIFilePath(string TIName, TestItemSourceType sourceType = TestItemSourceType.UserDefined, TestItemType itemType = TestItemType.TI, TestItemRunType runType = TestItemRunType.UUT)
        {
            return Path.Combine(GetTIFileFolder(sourceType, itemType, runType), TIName + ".tix");
        }

        public static List<string> GetTIFileNamesByPath(TestItemSourceType sourceType = TestItemSourceType.UserDefined, TestItemType itemType = TestItemType.TI, TestItemRunType runType = TestItemRunType.UUT)
        {
            string folder;
            List<string> TIFileNames = null;

            folder = GetTIFileFolder(sourceType, itemType, runType);

            if (Directory.GetFiles(folder, "*.tix").Length == 0)
            {
                //NO matching *.tix files
            }
            else
            {
                //has matching *.tix files
                TIFileNames = Directory.GetFiles(folder, "*.tix").Select(Path.GetFileNameWithoutExtension).ToList();
            }
            return TIFileNames;
        }

        public static bool IsThreadTIOrSubTI(string folderName) => SplitPathIntoDirectories(folderName).Any(dir => dir == "ThreadTI" || dir == "SubTI");

        public static IEnumerable<string> SplitPathIntoDirectories(string path)
        {
            if (Path.HasExtension(path))    // folder + filename
            {
                return Path.GetDirectoryName(path).Split('\\');
            }
            else // Pure folder
            {
                return path.Split('\\');
            }
        }

        public static void PP5LogIn()
        {
            try
            {
                CurrentDriver.SwitchToWindow(out bool switchToMainPanel);
                if (!switchToMainPanel)
                    return;

                WaitUntil(() => CheckWindowOpened(PowerPro5Config.LoginWindowName));

                if (!CheckLoginPageIsOpened())
                    return;

                // KeyboardInput VirtualKeys=""root"Keys.Tab + Keys.TabKeys.Tab + Keys.TabKeys.Tab + Keys.Tab" CapsLock=False NumLock=True ScrollLock=False
                Logger.LogMessage("KeyboardInput VirtualKeys=\"\"root\" CapsLock=False NumLock=True ScrollLock=False");
                System.Threading.Thread.Sleep(100);

                //PP5Session.FindElementByAccessibilityId("Username")?.SendKeys("root");
                //PP5Session.FindElementByAccessibilityId("OK_button")?.Click();

                CurrentDriver.GetWebElementFromWebDriver(MobileBy.AccessibilityId("Username")).SendText("root");
                CurrentDriver.GetWebElementFromWebDriver(MobileBy.AccessibilityId("OK_button")).LeftClick();

                WaitUntil(() => CheckPP5WindowIsOpened(WindowType.MainPanelMenu));
                //WaitUntil(() => CheckMainPanelIsOpened());

                //// Tag current window type as MainPanelMenu in text file
                //WindowTypeHelper.UpdateCurrentWindowType(WindowType.MainPanelMenu);
            }
            catch (Exception ex)
            {
                Logger.LogMessage(ex.ToString());
                return;
            }
        }

        // At MainPanel, Open New TI
        public void OpenNewTIEditorWindow()
        {
            // Delete TI Backup file if existed
            //if (File.Exists(Path.Combine(PowerPro5Config.ReleaseDataFolder, "TestItemBackup.tix")))
            //    File.Delete(Path.Combine(PowerPro5Config.ReleaseDataFolder, "TestItemBackup.tix"));
            PowerPro5Config.ReleaseDataFolder.DeleteFilesWithDifferentExtension("TestItemBackup");

            if (!CheckMainPanelIsOpened())
            {
                // Switch to PP5 IDE session
                AutoUIExecutor.SwitchTo(SessionType.PP5IDE);

                // Close all existing windows
                while (IsPP5IDEWindow(PP5IDEWindow.Text))
                {
                    // Click close button
                    //PP5IDEWindow.GetElement<IElement, IElement>(MobileBy.AccessibilityId("CloseButton")).LeftClick();
                    PP5IDEWindow.PerformClick("/ByIdOrName[CloseButton]", ClickType.LeftClick);

                    // If save window message box popup, click No
                    if (CurrentDriver.CheckElementExisted(MobileBy.AccessibilityId("MessageBoxExDialog")))
                        //PP5IDEWindow.GetElement(timeOut: SharedSetting.NORMAL_TIMEOUT, MobileBy.AccessibilityId("MessageBoxExDialog"),
                        //                                        By.Name("No")).LeftClick();
                        PP5IDEWindow.PerformClick("/ByIdOrName[MessageBoxExDialog,No]", ClickType.LeftClick);
                }


                // Open New TI Editor
                MenuSelect("Functions", "TI Editor");
                //WaitUntil(() => .Displayed);
                // Adam, 2024/07/08, only get PP5 IDE window when module is launched
                WaitUntil(() => PP5IDEWindow.Displayed);
                //isIDEWindowPresent = true;
            }
            else
            {
                // Click "Test Item" button in MainPanel
                //CurrentDriver.GetWebElementFromWebDriver(By.Name("Test Item")).LeftClick();
                CurrentDriver.PerformClick("/ByName#Retry[Test Item]", ClickType.LeftClick);

                System.Threading.SpinWait.SpinUntil(() => Process.GetProcessesByName(PowerPro5Config.IDEProcessName).Count() >= 2, 10000);

                while (Process.GetProcessesByName(PowerPro5Config.IDEProcessName).All(p => p.MainWindowHandle == new IntPtr()))
                {
                    System.Threading.Thread.Sleep(100);
                }

                //AutoUIExecutor.StartIDESession();
                AutoUIExecutor.SwitchTo(SessionType.PP5IDE);
            }

            // Open New TI
            //PerformOpenNewTI();
            OpenNewTI();

            // Wait until TIEditor window is opened
            WaitUntil(() => CheckPP5WindowIsOpened(WindowType.TIEditor));
        }

        // At MainPanel, Open New TPEditor Window
        public void OpenNewTPEditorWindow()
        {
            // Delete TP Backup file if existed
            PowerPro5Config.ReleaseDataFolder.DeleteFilesWithDifferentExtension("TestProgramBackup");
            //if (File.Exists(Path.Combine(PowerPro5Config.ReleaseDataFolder, "TestProgramBackup.tpx")))
            //    File.Delete(Path.Combine(PowerPro5Config.ReleaseDataFolder, "TestProgramBackup.tpx"));

            if (!CheckMainPanelIsOpened())
            {
                // Switch to PP5 IDE session
                AutoUIExecutor.SwitchTo(SessionType.PP5IDE);

                // Check how many windows need to close
                IEnumerable<string> openedWindows = GetIntersectionWithOrder(GetSubMenuListItemNames("Windows"), moduleNames);


                // If TP Editor window existed, close all windows
                //if (openedWindows.Contains(WindowType.TPEditor.GetDescription()))
                //{
                //    int windowNumbersToClose = openedWindows.Count();

                //    for (int i = 0; i < windowNumbersToClose; i++)
                //    {
                //        CurrentDriver.GetElementFromWebElement(MobileBy.AccessibilityId("CloseButton")).LeftClick();

                //        // If save window message box popup, click No
                //        if (CurrentDriver.CheckElementExisted(MobileBy.AccessibilityId("MessageBoxExDialog")))
                //            CurrentDriver.GetElementFromWebElement(timeOut: 5000, MobileBy.AccessibilityId("MessageBoxExDialog"),
                //                                                    By.Name("No")).LeftClick();
                //    }
                //}
                while (IsPP5IDEWindow(PP5IDEWindow.Text))
                {
                    // Click close button
                    //PP5IDEWindow.GetWebElementFromWebElement(MobileBy.AccessibilityId("CloseButton")).LeftClick();
                    PP5IDEWindow.PerformClick("/ByIdOrName[CloseButton]", ClickType.LeftClick);

                    // If save window message box popup, click No
                    if (PP5IDEWindow.GetWebElementFromWebElement(MobileBy.AccessibilityId("MessageBoxExDialog")) != null)
                        //PP5IDEWindow.GetElement(timeOut: 5000, MobileBy.AccessibilityId("MessageBoxExDialog"),
                        //                                        By.Name("No")).LeftClick();
                        PP5IDEWindow.PerformClick("/ByIdOrName[MessageBoxExDialog, No]", ClickType.LeftClick);
                }

                // Open New TP Editor
                MenuSelect("Functions", "TP Editor");
                WaitUntil(() => PP5IDEWindow.Displayed);
                //isIDEWindowPresent = true;
            }
            else
            {
                // Click "Test Program" button in MainPanel
                //CurrentDriver.GetWebElementFromWebDriver(By.Name("Test Program")).LeftClick();
                CurrentDriver.PerformClick("/ByName[Test Program]", ClickType.LeftClick);

                System.Threading.SpinWait.SpinUntil(() => Process.GetProcessesByName(PowerPro5Config.IDEProcessName).Count() >= 2, 10000);
                while (Process.GetProcessesByName(PowerPro5Config.IDEProcessName).All(p => p.MainWindowHandle == new IntPtr()))
                {
                    System.Threading.Thread.Sleep(100);
                }

                //                foreach (var pp5IDE in Process.GetProcessesByName(PowerPro5Config.IDEProcessName))
                //                {
                //                    // Wait for pp5IDE's MainWindowHandle is created
                //                    while (pp5IDE.MainWindowHandle == new IntPtr())
                //                    {
                //                        System.Threading.Thread.Sleep(100);
                //#if WRITE_LOG
                //                    Console.WriteLine($"pp5IDE.MainWindowHandle:{pp5IDE.MainWindowHandle}");
                //                    Console.WriteLine("IDE not ready yet, sleep for 100ms.");
                //#endif
                //                    }
                //                }

                AutoUIExecutor.SwitchTo(SessionType.PP5IDE);
            }

            PerformOpenNewTP(false);

            // Wait until TPEditor window is opened
            WaitUntil(() => CheckPP5WindowIsOpened(WindowType.TPEditor));
        }

        // At MainPanel, Open New GUI Editor
        public void OpenDefaultGUIEditorWindow()
        {
            if (!CheckMainPanelIsOpened())
            {
                // Switch to PP5 IDE session
                AutoUIExecutor.SwitchTo(SessionType.PP5IDE);

                // Check how many windows need to close
                IEnumerable<string> openedWindows = GetIntersectionWithOrder(GetSubMenuListItemNames("Windows"), moduleNames);

                // If GUI Editor window existed, close all windows
                if (openedWindows.Contains(WindowType.GUIEditor.GetDescription()))
                {
                    int windowNumbersToClose = openedWindows.Count();

                    for (int i = 0; i < windowNumbersToClose; i++)
                    {
                        CurrentDriver.GetWebElementFromWebDriver(MobileBy.AccessibilityId("CloseButton")).LeftClick();

                        // If save window message box popup, click No
                        if (CurrentDriver.CheckElementExisted(MobileBy.AccessibilityId("MessageBoxExDialog")))
                            CurrentDriver.GetWebElementFromWebDriver(timeOut: 5000, MobileBy.AccessibilityId("MessageBoxExDialog"),
                                                                    By.Name("No")).LeftClick();
                    }
                }

                MenuSelect("Functions", "GUI Editor");

                //if (CheckPP5WindowIsOpened(WindowType.GUIEditor))
                //{
                //    // If GUI window existed, open a new GUI template
                //    MenuSelect("File", "New");
                //}
            }
            else
            {
                // Click "GUI Editor" button in MainPanel
                CurrentDriver.GetWebElementFromWebDriver(By.Name("GUI Editor")).LeftClick();

                System.Threading.SpinWait.SpinUntil(() => Process.GetProcessesByName(PowerPro5Config.IDEProcessName).Count() >= 2, 10000);
                while (Process.GetProcessesByName(PowerPro5Config.IDEProcessName).All(p => p.MainWindowHandle == new IntPtr()))
                {
                    System.Threading.Thread.Sleep(100);
                }

                AutoUIExecutor.SwitchTo(SessionType.PP5IDE);

                //                foreach (var pp5IDE in Process.GetProcessesByName(PowerPro5Config.IDEProcessName))
                //                {
                //                    // Wait for pp5IDE's MainWindowHandle is created
                //                    while (pp5IDE.MainWindowHandle == new IntPtr())
                //                    {
                //                        System.Threading.Thread.Sleep(100);
                //#if WRITE_LOG
                //                    Console.WriteLine($"pp5IDE.MainWindowHandle:{pp5IDE.MainWindowHandle}");
                //                    Console.WriteLine("IDE not ready yet, sleep for 100ms.");
                //#endif
                //                    }
                //                }

                //while(!CheckPP5WindowIsOpened(WindowType.GUIEditor)){}

                //System.Threading.SpinWait.SpinUntil(() => CheckPP5WindowIsOpened(WindowType.GUIEditor));
                //WaitUntil(() => CheckPP5WindowIsOpened(WindowType.GUIEditor));

                //System.Threading.SpinWait.SpinUntil(() => Process.GetProcessesByName(PowerPro5Config.IDEProcessName).Count() >= 2, 10000);
                //foreach (var pp5IDE in Process.GetProcessesByName(PowerPro5Config.IDEProcessName))
                //    pp5IDE.WaitForInputIdle();

            }
            // Wait until GUIEditor window is opened
            WaitUntil(() => CheckPP5WindowIsOpened(WindowType.GUIEditor));

            MenuSelect("File", "New");

            // Wait until GUIEditor window is opened
            WaitUntil(() => CheckPP5WindowIsOpened(WindowType.GUIEditor));
        }

        // At MainPanel, Open New Report window
        public void OpenDefaultReportWindow()
        {
            if (!CheckMainPanelIsOpened())
            {
                // Switch to PP5 IDE session
                AutoUIExecutor.SwitchTo(SessionType.PP5IDE);

                // Check how many windows need to close
                IEnumerable<string> openedWindows = GetIntersectionWithOrder(GetSubMenuListItemNames("Windows"), moduleNames);

                // If Report window existed, close all windows
                if (openedWindows.Contains(WindowType.Report.GetDescription()))
                {
                    int windowNumbersToClose = openedWindows.Count();

                    for (int i = 0; i < windowNumbersToClose; i++)
                    {
                        CurrentDriver.GetWebElementFromWebDriver(MobileBy.AccessibilityId("CloseButton")).LeftClick();

                        // If save window message box popup, click No
                        if (CurrentDriver.CheckElementExisted(MobileBy.AccessibilityId("MessageBoxExDialog")))
                            CurrentDriver.GetWebElementFromWebDriver(timeOut: 5000, MobileBy.AccessibilityId("MessageBoxExDialog"),
                                By.Name("No")).LeftClick();
                    }
                }

                MenuSelect("Functions", "Report");

                //if (CheckPP5WindowIsOpened(WindowType.Report))
                //{
                //    // If Report window existed, open a new Report template
                //    MenuSelect("File", "New");

                //    // If save Report window message box popup, click No
                //    if (CurrentDriver.CheckElementExisted(By.Name("Report")))
                //        CurrentDriver.GetElementFromWebElement(timeOut: 5000, By.Name("Report"),
                //                                                By.Name("No")).LeftClick();
                //}
                //else
                //    MenuSelect("Functions", "Report");

                //System.Threading.SpinWait.SpinUntil(() => CheckPP5WindowIsOpened(WindowType.Report), 15000);
            }
            else
            {
                // Click "Report" button in MainPanel
                CurrentDriver.GetWebElementFromWebDriver(By.Name("Report")).LeftClick();

                //while(!CheckPP5WindowIsOpened(WindowType.GUIEditor)){}
                //System.Threading.SpinWait.SpinUntil(() => CheckPP5WindowIsOpened(WindowType.Report), 15000);

                //System.Threading.SpinWait.SpinUntil(() => Process.GetProcessesByName(PowerPro5Config.IDEProcessName).Count() >= 2, 10000);
                System.Threading.SpinWait.SpinUntil(() => Process.GetProcessesByName(PowerPro5Config.IDEProcessName).Count() >= 2, 10000);
                while (Process.GetProcessesByName(PowerPro5Config.IDEProcessName).All(p => p.MainWindowHandle == new IntPtr()))
                {
                    System.Threading.Thread.Sleep(100);
                }

                AutoUIExecutor.SwitchTo(SessionType.PP5IDE);

                //                foreach (var pp5IDE in Process.GetProcessesByName(PowerPro5Config.IDEProcessName))
                //                {
                //                    // Wait for pp5IDE's MainWindowHandle is created
                //                    while (pp5IDE.MainWindowHandle == new IntPtr())
                //                    {
                //                        System.Threading.Thread.Sleep(100);
                //#if WRITE_LOG
                //                    Console.WriteLine($"pp5IDE.MainWindowHandle:{pp5IDE.MainWindowHandle}");
                //                    Console.WriteLine("IDE not ready yet, sleep for 100ms.");
                //#endif
                //                    }
                //                }


                //System.Threading.SpinWait.SpinUntil(() => CheckPP5WindowIsOpened(WindowType.Report));
                //WaitUntil(() => CheckPP5WindowIsOpened(WindowType.Report));
            }

            // Wait until report window is opened
            WaitUntil(() => CheckPP5WindowIsOpened(WindowType.Report));
        }

        // At MainPanel, Open New Management window
        public void OpenDefaultManagementWindow()
        {
            if (!CheckMainPanelIsOpened())
            {
                // Switch to PP5 IDE session
                AutoUIExecutor.SwitchTo(SessionType.PP5IDE);

                //// Check how many windows need to close
                //IEnumerable<string> openedWindows = GetIntersectionWithOrder(GetSubMenuListItemNames("Windows"), moduleNames);

                //// If Management window existed, close all windows
                //if (openedWindows.Contains(WindowType.Management.GetDescription()))
                //{
                //    int windowNumbersToClose = openedWindows.Count();

                //    for (int i = 0; i < windowNumbersToClose; i++)
                //    {
                //        CurrentDriver.GetElementFromWebElement(MobileBy.AccessibilityId("CloseButton")).LeftClick();

                //        // If save window message box popup, click No
                //        if (CurrentDriver.CheckElementExisted(MobileBy.AccessibilityId("MessageBoxExDialog")))
                //            CurrentDriver.GetElementFromWebElement(timeOut: 5000, MobileBy.AccessibilityId("MessageBoxExDialog"),
                //                                                    By.Name("No")).LeftClick();
                //    }
                //}

                // Close all existing windows
                while (IsPP5IDEWindow(PP5IDEWindow.Text))
                {
                    // Click close button
                    //PP5IDEWindow.GetWebElementFromWebElement(MobileBy.AccessibilityId("CloseButton")).LeftClick();
                    PP5IDEWindow.PerformClick("/ByIdOrName[CloseButton]", ClickType.LeftClick);

                    // If save window message box popup, click No
                    if (CurrentDriver.CheckElementExisted(MobileBy.AccessibilityId("MessageBoxExDialog")))
                        //PP5IDEWindow.GetElement(timeOut: 5000, MobileBy.AccessibilityId("MessageBoxExDialog"),
                        //                                        By.Name("No")).LeftClick();
                        PP5IDEWindow.PerformClick("/ByIdOrName[MessageBoxExDialog,No]", ClickType.LeftClick);
                }

                MenuSelect("Functions", "Management");

                //if (CheckPP5WindowIsOpened(WindowType.Management))
                //{
                //    // If Management window existed, do nothing
                //}
                //else
                //    MenuSelect("Functions", "Management");

                //System.Threading.SpinWait.SpinUntil(() => CheckPP5WindowIsOpened(WindowType.Management), 3000);
                //WaitUntil(() => GetPP5Window().Displayed);
            }
            else
            {
                // Click "Management" button in MainPanel
                //CurrentDriver.GetWebElementFromWebDriver(By.Name("Management")).LeftClick();
                CurrentDriver.PerformClick("/ByName[Management]", ClickType.LeftClick);

                //AutoUIExecutor.SwitchTo(SessionType.PP5IDE);

                //while (!CheckPP5WindowIsOpened(WindowType.Management))
                //while (!System.Threading.SpinWait.SpinUntil(() => CheckPP5WindowIsOpened(WindowType.Management), 15000))
                //{
                //    var sessDetaileds = currentDriver.SessionDetails;
                //    AutoUIExecutor.SwitchTo(SessionType.PP5IDE);
                //}
                System.Threading.SpinWait.SpinUntil(() => Process.GetProcessesByName(PowerPro5Config.IDEProcessName).Count() >= 2, 10000);
                while (Process.GetProcessesByName(PowerPro5Config.IDEProcessName).All(p => p.MainWindowHandle == new IntPtr()))
                {
                    System.Threading.Thread.Sleep(100);
                }
                AutoUIExecutor.SwitchTo(SessionType.PP5IDE);

                //                foreach (var pp5IDE in Process.GetProcessesByName(PowerPro5Config.IDEProcessName))
                //                {
                //                    // Wait for pp5IDE's MainWindowHandle is created
                //                    while (pp5IDE.MainWindowHandle == new IntPtr())
                //                    {
                //                        System.Threading.Thread.Sleep(100);
                //#if WRITE_LOG
                //                    Console.WriteLine($"pp5IDE.MainWindowHandle:{pp5IDE.MainWindowHandle}");
                //                    Console.WriteLine("IDE not ready yet, sleep for 100ms.");
                //#endif
                //                    }
                //                }

                //System.Threading.SpinWait.SpinUntil(() => CheckPP5WindowIsOpened(WindowType.Management));
                //WaitUntil(() => CheckPP5WindowIsOpened(WindowType.Management), 5000);
            }

            // Wait until Management window is opened
            WaitUntil(() => CheckPP5WindowIsOpened(WindowType.Management));
        }

        // At MainPanel, Open New Execution Window
        public void OpenDefaultExecutionWindow(string tpName)
        {
            if (!CheckMainPanelIsOpened())
            {
                // Switch to PP5 IDE session
                AutoUIExecutor.SwitchTo(SessionType.PP5IDE);

                // Check how many windows need to close
                IEnumerable<string> openedWindows = GetIntersectionWithOrder(GetSubMenuListItemNames("Windows"), moduleNames);

                // If Execution window existed, close all windows
                if (openedWindows.Contains(WindowType.Execution.GetDescription()))
                {
                    int windowNumbersToClose = openedWindows.Count();

                    for (int i = 0; i < windowNumbersToClose; i++)
                    {
                        CurrentDriver.GetWebElementFromWebDriver(MobileBy.AccessibilityId("CloseButton")).LeftClick();

                        // If save window message box popup, click No
                        if (CurrentDriver.CheckElementExisted(MobileBy.AccessibilityId("MessageBoxExDialog")))
                            CurrentDriver.GetWebElementFromWebDriver(timeOut: 5000, MobileBy.AccessibilityId("MessageBoxExDialog"),
                                                                    By.Name("No")).LeftClick();
                    }
                }

                MenuSelect("Functions", "Execution");
                //WaitUntil(() => GetPP5Window().Displayed);
                // Adam, 2024/07/08, only get PP5 IDE window when module is launched
                WaitUntil(() => PP5IDEWindow.Displayed);
                //isIDEWindowPresent = true;

                //if (CheckPP5WindowIsOpened(WindowType.Execution))
                //{
                //    // If Execution window existed, open an existing TP
                //    MenuSelect("File", "OpenTestProgram...");
                //}
                //else
                //{
                //    MenuSelect("Functions", "Execution");
                //}
                //IWebElement TPtoOpen = CurrentDriver.GetElementFromWebElement(By.Name("Open"))
                //                                    .GetCellBy(1, 2);

                ////string TPName = TPtoOpen.Text;
                //TPtoOpen.LeftClick();
                //CurrentDriver.GetElementFromWebElement(5000, By.Name("Open"), By.Name("Ok"))
                //             .LeftClick();

                //while(GetPP5Window().GetEditContent(1) != TPtoOpen.Text){}
                //System.Threading.SpinWait.SpinUntil(() => GetPP5Window().GetElementFromWebElement(By.ClassName("FullScreenExecutionStateAeraView"))
                //                                                        .GetFirstPaneElement()
                //                                                        .GetEditElement(1).GetCellValue() == tpName, 3000);
            }
            else
            {
                // Click "Execution" button in MainPanel
                CurrentDriver.GetWebElementFromWebDriver(By.Name("Execution")).LeftClick();

                System.Threading.SpinWait.SpinUntil(() => Process.GetProcessesByName(PowerPro5Config.IDEProcessName).Count() >= 2, 10000);
                while (Process.GetProcessesByName(PowerPro5Config.IDEProcessName).All(p => p.MainWindowHandle == new IntPtr()))
                {
                    System.Threading.Thread.Sleep(100);
                }
                AutoUIExecutor.SwitchTo(SessionType.PP5IDE);

                //                foreach (var pp5IDE in Process.GetProcessesByName(PowerPro5Config.IDEProcessName))
                //                {
                //                    // Wait for pp5IDE's MainWindowHandle is created
                //                    while (pp5IDE.MainWindowHandle == new IntPtr())
                //                    {
                //                        System.Threading.Thread.Sleep(100);
                //#if WRITE_LOG
                //                    Console.WriteLine($"pp5IDE.MainWindowHandle:{pp5IDE.MainWindowHandle}");
                //                    Console.WriteLine("IDE not ready yet, sleep for 100ms.");
                //#endif
                //                    }
                //                }

                //System.Threading.Thread.Sleep(5000);

                //System.Threading.SpinWait.SpinUntil(() => GetPP5Window().GetElementFromWebElement(By.ClassName("FullScreenExecutionStateAeraView"))
                //    .GetFirstPaneElement()
                //    .GetEditElement(1).GetCellValue() == tpName);
                //WaitUntil(() => GetPP5Window().GetElementFromWebElement(By.ClassName("FullScreenExecutionStateAeraView"))
                //                              .GetFirstPaneElement()
                //                              .GetEditElement(1).GetCellValue() == tpName, 8000);
            }

            PerformOpenTPFile(tpName);

            // Wait until execution window is loaded
            // Adam, 2024/07/08, only get PP5 IDE window when module is launched
            WaitUntil(() => PP5IDEWindow.GetExtendedElementBySingleWithRetry(PP5By.ClassName("FullScreenExecutionStateAeraView"))
                                        .GetFirstPaneElement()
                                        .GetEditElement(1).GetCellValue() == tpName, 8000);
            //isIDEWindowPresent = true;
        }

        //public void PerformOpenNewGUIFile()
        //{
        //    MenuSelect("File", "New");
        //}

        public void PerformOpenTPFile(string TPName)
        {
            CurrentDriver.GetExtendedElementBySingleWithRetry(PP5By.Name("Open"))
                         .GetFirstDataGridElement()
                         .GetCellByName(2, TPName)
                         .LeftClick();

            CurrentDriver.GetWebElementFromWebDriver(5000, By.Name("Open"), By.Name("Ok"))
                         .LeftClick();
        }

        public void PerformOpenAndSaveTI(string TIName, TestItemType tiType, TestItemRunType tiRunType)
        {
            PerformOpenNewTI(tiType, tiRunType);
            SaveAsNewTI(TIName);
            //TIFilePath = GetTIFilePath(TIName, TestItemSourceType.UserDefined, tiType, tiRunType);
            //UpdateThreadOrSubTiCmdGroup("Thread TI");
            //UpdateThreadOrSubTiCmdGroup("Sub TI");
        }

        public void PerformOpenAndSaveTI(string TIName, TestItemType tiType, TestItemRunType tiRunType, bool isInputDescription = false, string desc = "")
        {
            PerformOpenNewTI(tiType, tiRunType);
            SaveAsNewTI(TIName, isInputDescription, desc);
            //TIFilePath = GetTIFilePath(TIName, TestItemSourceType.UserDefined, tiType, tiRunType);
            //UpdateThreadOrSubTiCmdGroup("Thread TI");
            //UpdateThreadOrSubTiCmdGroup("Sub TI");
        }

        public void PerformOpenNewTP(bool tpNotSaved)
        {
            // If save TP window message box popup, click No
            if (tpNotSaved)
            {
                if (CurrentDriver.CheckElementExisted(By.Name("Exit")))
                    CurrentDriver.GetWebElementFromWebDriver(timeOut: 5000, By.Name("Exit"),
                                                            By.Name("No")).LeftClick();
            }

            //if (fromMainPanel)
            //{
            // LeftClick on RadioButton "New Test Program"
            Logger.LogMessage("LeftClick on RadioButton \"New Test Program\"");
            CurrentDriver.GetWebElementFromWebDriver(timeOut: 5000, By.Name("Enter Point"),
                                                    MobileBy.AccessibilityId("NewRadioBtn")).LeftClick();

            // Enter Point window, LeftClick on Text "Ok"
            Logger.LogMessage("LeftClick on Text \"Ok\"");
            CurrentDriver.GetWebElementFromWebDriver(timeOut: 5000, By.Name("Enter Point"),
                                                    By.Name("Ok")).LeftClick();
            //}

            // New Test Program window, LeftClick on Button "Ok"
            Logger.LogMessage("LeftClick on Button \"Ok\"");
            CurrentDriver.GetWebElementFromWebDriver(timeOut: 5000, By.Name("New Test Program"),
                                                    By.Name("Ok")).LeftClick();

            // If save TP window message box popup, click No
            CurrentDriver.SwitchToWindow(out _);
            if (CurrentDriver.CheckElementExisted(MobileBy.AccessibilityId("MessageBoxExDialog")))
                CurrentDriver.GetWebElementFromWebDriver(timeOut: 5000, MobileBy.AccessibilityId("MessageBoxExDialog"),
                                                        By.Name("OK")).LeftClick();
        }

        public void PerformOpenNewTI(bool closeCurrTI = true)
        {
            // If save TI window message box popup, click No
            if (closeCurrTI)
                PerformCloseTI();

            MenuSelect("Functions", "TI Editor");
            WaitUntil(() => GetPP5Window() != null);

            OpenNewTI(tiType: TestItemType.TI, tiRunType: TestItemRunType.UUT);

            //// LeftClick on RadioButton "New Test Item"
            ////if (fromMainPanel)
            ////{
            ////Console.WriteLine("LeftClick on RadioButton \"New Test Item\"");
            //PP5IDEWindow.GetElementWithRetry(timeOut: SharedSetting.NORMAL_TIMEOUT, nTryCount: 2,
            //                                          By.Name("Enter Point"),
            //                                          MobileBy.AccessibilityId("NewRadioBtn")).LeftClick();
            ////Thread.Sleep(1500);

            //// Enter Point window, LeftClick on Text "Ok"
            ////Console.WriteLine("LeftClick on Text \"Ok\"");
            //PP5IDEWindow.GetElementFromWebElement(timeOut: SharedSetting.NORMAL_TIMEOUT, 
            //                                 By.Name("Enter Point"),
            //                                 By.Name("Ok")).LeftClick();
            ////}

            //// New Test Item window, LeftClick on Button "Ok"
            ////Console.WriteLine("LeftClick on Button \"Ok\"");
            //PP5IDEWindow.GetElementFromWebElement(timeOut: SharedSetting.NORMAL_TIMEOUT, 
            //                                 MobileBy.AccessibilityId("LoginDialog"),
            //                                 MobileBy.AccessibilityId("OkBtn")).LeftClick();

            //// Close time consuming message box (Debug Mode)
            //// If save TI window message box popup, click No
            //CurrentDriver.SwitchToWindow(out _);
            //if (CurrentDriver.CheckElementExisted(MobileBy.AccessibilityId("MessageBoxExDialog")))
            //    PP5IDEWindow.GetElementFromWebElement(timeOut: 5000, MobileBy.AccessibilityId("MessageBoxExDialog"),
            //                                            By.Name("OK")).LeftClick();

            //AutoUIExecutor.SwitchTo(SessionType.PP5IDE);
        }

        public void PerformOpenNewTI(TestItemType tiType, TestItemRunType tiRunType)
        {
            // If save TI window message box popup, click No
            PerformCloseTI();

            MenuSelect("Functions", "TI Editor");
            WaitUntil(() => GetPP5Window() != null);

            OpenNewTI(tiType, tiRunType);

            //// Close time consuming message box (Debug Mode)
            //// If save TI window message box popup, click No
            //CurrentDriver.SwitchToWindow(out _);
            //if (CurrentDriver.CheckElementExisted(MobileBy.AccessibilityId("MessageBoxExDialog")))
            //    CurrentDriver.GetWebElementFromWebDriver(timeOut: 5000, MobileBy.AccessibilityId("MessageBoxExDialog"),
            //                                            By.Name("OK")).LeftClick();
        }

        public void LoadOldTI(string tiName)
        {
            // LeftClick on RadioButton "Load Test Item"
            //Console.WriteLine("LeftClick on RadioButton \"Load Test Item\"");

            IElement TIEnterWindow = PP5IDEWindow.GetExtendedElementBySingleWithRetry(PP5By.Name("Enter Point"));

            TIEnterWindow.GetRdoBtnElement("LoadRadioBtn").LeftClick();

            // Enter Point window, LeftClick on Text "Ok"
            //Console.WriteLine("LeftClick on Text \"Ok\"");
            TIEnterWindow.GetBtnElement("Ok").LeftClick();

            //RefreshDataTable(DataTableAutoIDType.LoginGrid);
            //List<string> existingTINames = GetSingleColumnValues(DataTableAutoIDType.LoginGrid, "Test Item Name", excludeLastRow: false);

            //if (!existingTINames.Contains(tiName))
            //    Assert.Fail($"No TI existed with given TI Name: {tiName}");


            // Search the TI
            IElement LoadTIWindow = PP5IDEWindow.GetExtendedElementBySingleWithRetry(PP5By.Name("Load Test Item"));

            LoadTIWindow.GetEditElement("searchText").SendText(tiName);
            Press(Keys.Enter);

            while (LoadTIWindow.GetSelectedRow("LoginGrid").GetCellValue(1 /*"Test Item Name"*/) != tiName)
            {
                LoadTIWindow.GetBtnElement("NextBtn").LeftClick();
            }

            // Load Test Item window, LeftClick on Button "Ok"
            //Console.WriteLine("LeftClick on Button \"Ok\"");
            //PP5IDEWindow.GetElementFromWebElement(timeOut: 5000, MobileBy.AccessibilityId("LoginDialog"),
            //                                        MobileBy.AccessibilityId("OkBtn")).LeftClick();

            // If TI is found, in Load Test Item window, LeftClick on Button "Ok"
            LoadTIWindow.GetBtnElement("Ok").LeftClick();

            // Close time consuming message box (Debug Mode)
            // If save TI window message box popup, click No
            //AutoUIExecutor.SwitchTo(SessionType.PP5IDE);
            //if (CurrentDriver.CheckElementExisted(MobileBy.AccessibilityId("MessageBoxExDialog")))
            //CurrentDriver.GetElementFromWebElement(timeOut: 5000, MobileBy.AccessibilityId("MessageBoxExDialog"),
            //                                        By.Name("OK")).LeftClick();
        }

        public void LoadOldTI(string tiName, TestItemType type = TestItemType.TI, TestItemRunType runType = TestItemRunType.UUT)
        {
            // LeftClick on RadioButton "Load Test Item"
            //Console.WriteLine("LeftClick on RadioButton \"Load Test Item\"");
            //PP5IDEWindow.GetElementFromWebElement(timeOut: 5000, By.Name("Enter Point"),
            //                                        MobileBy.AccessibilityId("LoadRadioBtn")).LeftClick();
            PP5IDEWindow/*.GetWindowElement("Enter Point")*/
                        .GetWebElementFromWebElement(MobileBy.AccessibilityId("LoadRadioBtn"))
                        .LeftClick();

            // Enter Point window, LeftClick on Text "Ok"
            //Console.WriteLine("LeftClick on Text \"Ok\"");
            //PP5IDEWindow.GetElementFromWebElement(timeOut: 5000, By.Name("Enter Point"),
            //                                        By.Name("Ok")).LeftClick();
            PP5IDEWindow.GetExtendedElementBySingleWithRetry(PP5By.Name("Enter Point"))
                        .GetBtnElement("Ok")
                        .LeftClick();

            // Choose the type and run type
            //PP5IDEWindow.GetElementFromWebElement(MobileBy.AccessibilityId("LoginDialog"))
            //             .GetElementFromWebElement(By.Name(type.GetDescription())).LeftClick();
            PP5IDEWindow.GetExtendedElementBySingleWithRetry(PP5By.Name("Load Test Item"))
                        .GetRdoBtnElement(type.GetDescription())
                        .LeftClick();

            //PP5IDEWindow.GetElementFromWebElement<ElementControlType.Window, ElementControlType.RadioButton>
            //                        ("Load Test Item", By.Name(type.GetDescription()))

            //GetRdoBtnElement(type.GetDescription()).LeftClick();
            //PP5IDEWindow.GetElementFromWebElement(MobileBy.AccessibilityId("Load Test Item"))
            //             .GetElementFromWebElement(By.Name(runType.GetDescription())).LeftClick();
            PP5IDEWindow.GetExtendedElementBySingleWithRetry(PP5By.Name("Load Test Item"))
                        .GetRdoBtnElement(runType.GetDescription())
                        .LeftClick();

            //.GetRdoBtnElement(runType.GetDescription()).LeftClick();

            //RefreshDataTable(DataTableAutoIDType.LoginGrid);
            //List<string> existingTINames = GetSingleColumnValues(DataTableAutoIDType.LoginGrid, "Test Item Name", excludeLastRow: false);
            //if (!existingTINames.Contains(tiName))
            //    Assert.Fail($"No TI existed with given TI Name: {tiName}");

            //GetCellBy("LoginGrid", existingTINames.IndexOf(tiName), "Test Item Name").LeftClick();
            PP5IDEWindow.GetFirstDataGridElement()
                        .GetCellByName(1, tiName)
                        .LeftClick();

            // Load Test Item window, LeftClick on Button "Ok"
            //Console.WriteLine("LeftClick on Button \"Ok\"");
            //PP5IDEWindow.GetElementFromWebElement(timeOut: 5000, MobileBy.AccessibilityId("LoginDialog"),
            //                                        MobileBy.AccessibilityId("OkBtn")).LeftClick();
            PP5IDEWindow.GetExtendedElementBySingleWithRetry(PP5By.Name("Load Test Item"))
                        .GetBtnElement("Ok")
                        .LeftClick();

            // Close time consuming message box (Debug Mode)
            // If save TI window message box popup, click No
            //AutoUIExecutor.SwitchTo(SessionType.PP5IDE);
            //if (CurrentDriver.CheckElementExisted(MobileBy.AccessibilityId("MessageBoxExDialog")))
            //CurrentDriver.GetElementFromWebElement(timeOut: 5000, MobileBy.AccessibilityId("MessageBoxExDialog"),
            //                                        By.Name("OK")).LeftClick();
        }

        public void LoadOldTI(string tiName, out string desc)
        {
            // LeftClick on RadioButton "Load Test Item"
            IElement TIEnterWindow = PP5IDEWindow.GetExtendedElementBySingleWithRetry(PP5By.Name("Enter Point"));

            TIEnterWindow.GetRdoBtnElement("LoadRadioBtn").LeftClick();

            // Enter Point window, LeftClick on Text "Ok"
            TIEnterWindow.GetBtnElement("Ok").LeftClick();

            IElement LoadTIWindow = PP5IDEWindow.GetExtendedElementBySingleWithRetry(PP5By.Name("Load Test Item"));

            // Search the TI
            LoadTIWindow.GetEditElement("searchText").SendText(tiName);
            Press(Keys.Enter);

            while (LoadTIWindow.GetSelectedRow("LoginGrid").GetCellValue(0 /*"Test Item Name"*/) != tiName)
            {
                LoadTIWindow.GetBtnElement("NextBtn").LeftClick();
            }

            // Get Description result
            desc = LoadTIWindow.GetEditElement("DesTxtBox").Text;

            // If TI is found, in Load Test Item window, LeftClick on Button "Ok"
            LoadTIWindow.GetBtnElement("Ok").LeftClick();

            // Close time consuming message box (Debug Mode)
            // If save TI window message box popup, click No
            //AutoUIExecutor.SwitchTo(SessionType.PP5IDE);
            //CurrentDriver.GetElementFromWebElement(timeOut: 5000, MobileBy.AccessibilityId("MessageBoxExDialog"), By.Name("OK")).LeftClick();
        }

        public void PerformLoadOldTI(string tiName)
        {
            MenuSelect("File", "Open...");

            // Search the TI
            IElement LoadTIWindow = PP5IDEWindow.GetExtendedElementBySingleWithRetry(PP5By.Name("Load Test Item"));

            LoadTIWindow.GetEditElement("searchText").SendText(tiName);
            Press(Keys.Enter);

            while (LoadTIWindow.GetSelectedRow("LoginGrid").GetCellValue(1 /*"Test Item Name"*/) != tiName)
            {
                LoadTIWindow.GetBtnElement("NextBtn").LeftClick();
            }

            // If TI is found, in Load Test Item window, LeftClick on Button "Ok"
            LoadTIWindow.GetBtnElement("Ok").LeftClick();
        }

        public void PerformLoadTIBySearchingTIName(string tiName)
        {
            // LeftClick on RadioButton "Load Test Item"
            IElement TIEnterWindow = PP5IDEWindow.GetExtendedElementBySingleWithRetry(PP5By.Name("Enter Point"));

            TIEnterWindow.GetRdoBtnElement("LoadRadioBtn").LeftClick();

            // Enter Point window, LeftClick on Text "Ok"
            TIEnterWindow.GetBtnElement("Ok").LeftClick();

            // Search TI by TIName
            //IWebElement TISearchBox = CurrentDriver.GetElementFromWebElement(5000, MobileBy.AccessibilityId("LoginDialog"),
            //                                                         MobileBy.AccessibilityId("searchBox"));
            //TISearchBox.ClearContent();
            //TISearchBox.SendComboKeys(tiName, Keys.Enter);

            IElement LoadTIWindow = PP5IDEWindow.GetExtendedElementBySingleWithRetry(PP5By.Name("Load Test Item"));

            // Search the TI
            LoadTIWindow.GetEditElement("searchText").SendText(tiName);
            Press(Keys.Enter);

            while (LoadTIWindow.GetSelectedRow("LoginGrid").GetCellValue(0 /*"Test Item Name"*/) != tiName)
            {
                LoadTIWindow.GetBtnElement("NextBtn").LeftClick();
            }

            // If TI is found, in Load Test Item window, LeftClick on Button "Ok"
            LoadTIWindow.GetBtnElement("Ok").LeftClick();

            // Close time consuming message box (Debug Mode)
            // If save TI window message box popup, click No
            //AutoUIExecutor.SwitchTo(SessionType.PP5IDE);
            //CurrentDriver.GetElementFromWebElement(timeOut: 5000, MobileBy.AccessibilityId("MessageBoxExDialog"), By.Name("OK")).LeftClick();
        }

        public void PerformCloseTI()
        {
            if (PP5IDEWindow.ModuleName != WindowType.TIEditor.GetDescription())
                return;
            PP5IDEWindow.PerformClick("/ByIdOrName[CloseButton]", ClickType.LeftClick); // Close the TI window
            if (PP5IDEWindow.PerformGetElement("/Window[Exit]") != null)
                PP5IDEWindow.PerformClick("/Window[Exit]/Button[No]", ClickType.LeftClick);
        }

        public void OpenNewTI()
        {
            OpenNewTI(TestItemType.TI, TestItemRunType.UUT);
        }

        public void OpenNewTI(TestItemType tiType, TestItemRunType tiRunType)
        {
            // Get the radio button text of given item type and run type
            string tiTypeText = tiType.GetDescription();
            string tiRunTypeText = tiRunType.GetDescription();

            // LeftClick on RadioButton "New Test Item"
            //Logger.LogMessage("LeftClick on RadioButton \"New Test Item\"");
            //CurrentDriver.GetElementFromWebElement(timeOut: 5000, By.Name("Enter Point"),
            //                                        MobileBy.AccessibilityId("NewRadioBtn")).LeftClick();
            //PP5IDEWindow.GetElementWithRetry(timeOut: SharedSetting.NORMAL_TIMEOUT, nTryCount: 2,
            //                                          By.Name("Enter Point"),
            //                                          MobileBy.AccessibilityId("NewRadioBtn")).LeftClick();
            //PP5IDEWindow.PerformClick("/ByName#Retry[Enter Point]/ByNameOrId[NewRadioBtn]", ClickType.LeftClick);
            //PP5IDEWindow.GetElement2(By.Name("Enter Point"), nRetry: 2)
            //            .GetElement2(MobileBy.AccessibilityId("NewRadioBtn")).LeftClick();
            //PP5IDEWindow.GetExtendedElementByChainedWithRetry(new By[] { PP5By.Name("Enter Point"), PP5By.Id("NewRadioBtn") })
            //            .LeftClick();
            //PP5By.Chained(PP5By.Name("Enter Point"), PP5By.Id("NewRadioBtn")).FindElement(PP5IDEWindow).LeftClick();
            //bi.FindElement(PP5IDEWindow).LeftClick();
            //PP5IDEWindow.GetExtendedElement(PP5By.Chained(PP5By.Name("Enter Point"), PP5By.Id("NewRadioBtn")))
            //            .LeftClick();
            //PP5IDEWindow.GetElementWithTimeoutAndRetry(SharedSetting.NORMAL_TIMEOUT, PP5By.Name("Enter Point"), PP5By.Id("NewRadioBtn")).LeftClick();
            PP5IDEWindow.PerformClick("/ByName#Retry[Enter Point]/ByNameOrId[NewRadioBtn]", ClickType.LeftClick);

            // Enter Point window, LeftClick on Text "Ok"
            //Logger.LogMessage("LeftClick on Text \"Ok\"");
            //CurrentDriver.GetElementFromWebElement(timeOut: 5000, By.Name("Enter Point"),
            //                                        By.Name("Ok")).LeftClick();
            PP5IDEWindow.PerformClick("/ByName[Enter Point, Ok]", ClickType.LeftClick);

            // New Test Item window, select TI Type & Run Type
            //CurrentDriver.GetElementFromWebElement(timeOut: 5000, MobileBy.AccessibilityId("LoginDialog"),
            //                                        By.Name(tiTypeText)).LeftClick();
            //CurrentDriver.GetElementFromWebElement(timeOut: 5000, MobileBy.AccessibilityId("LoginDialog"),
            //                                        By.Name(tiRunTypeText)).LeftClick();
            PP5IDEWindow.PerformClick($"/ByNameOrId[LoginDialog, {tiTypeText}]", ClickType.LeftClick);
            PP5IDEWindow.PerformClick($"/ByNameOrId[LoginDialog, {tiRunTypeText}]", ClickType.LeftClick);

            // New Test Item window, LeftClick on Button "Ok"
            //Logger.LogMessage("LeftClick on Button \"Ok\"");
            //CurrentDriver.GetElementFromWebElement(timeOut: 5000, MobileBy.AccessibilityId("LoginDialog"),
            //                                        MobileBy.AccessibilityId("OkBtn")).LeftClick();
            PP5IDEWindow.PerformClick("/ById[LoginDialog,OkBtn]", ClickType.LeftClick);
        }

        public void SaveAsNewTI(string tiName)
        {
            // LeftClick on Text "File" > "Save As..."
            MenuSelect("File", "Save As...");
            //WaitUntil(() => PP5IDEWindow.PerformGetElement("/Window[Save Test Item]") != null);
            PP5IDEWindow.PerformInput("/ByName#Retry[Save Test Item]/ById[TINameTxtBox]", InputType.SendContent, tiName);
            PP5IDEWindow.PerformClick("/Window[Save Test Item]/ByName[Ok]", ClickType.LeftClick);
            WaitUntilTIFinishedSaving();
            //WaitUntil(() => PP5IDEWindow.PerformGetElement("/ByCondition#ToolBar[Visible]/RadioButton[4]").Enabled);

            //// Rename the TI Name
            //if (GetSingleColumnValues("LoginGrid", "Test Item Name").Contains(tiName))
            //    Assert.Fail($"Same TI existed with TI Name: {tiName}");

            //PP5IDEWindow.GetElementFromWebElement(MobileBy.AccessibilityId("TINameTxtBox"))
            //            .SendText(tiName);
            // LeftClick on Button "Ok"
            //Logger.LogMessage("LeftClick on Button \"Ok\"");
            //PP5IDEWindow.GetElementFromWebElement(By.Name("Save Test Item"))
            //            .GetBtnElement("Ok")
            //            .LeftClick();

            // Wait for Save As radiobutton enabled (TI is saved)
            //WaitUntil(() => PP5IDEWindow.GetToolbarElement((e) => e.isElementVisible())
            //                            .GetRdoBtnElement(4)
            //                            .Enabled);
        }

        public void SaveAsNewTI(string tiName, bool isInputDescription = false, string desc = "")
        {
            // LeftClick on Text "File" > "Save As..."
            MenuSelect("File", "Save As...");

            //// Rename the TI Name
            //if (GetSingleColumnValues("LoginGrid", "Test Item Name").Contains(tiName))
            //    Assert.Fail($"Same TI existed with TI Name: {tiName}");

            PP5IDEWindow.GetWebElementFromWebElement(MobileBy.AccessibilityId("TINameTxtBox"))
                        .SendText(tiName);

            if (isInputDescription)
                PP5IDEWindow.GetWebElementFromWebElement(MobileBy.AccessibilityId("DesTxtBox"))
                            .SendText(desc);

            // LeftClick on Button "Ok"
            PP5IDEWindow.GetExtendedElementBySingleWithRetry(PP5By.Name("Save Test Item"))
                        .GetBtnElement("Ok")
                        .LeftClick();

            // Wait for Save As radiobutton enabled (TI is saved)
            WaitUntil(() => PP5IDEWindow.GetToolbarElement((e) => e.isElementVisible())
                                        .GetRdoBtnElement(4)
                                        .Enabled);
        }

        public void ChangeGroupAndSaveAsNewTI(string tiName, string group = "")
        {
            // LeftClick on Text "File" > "Save As..."
            MenuSelect("File", "Save As...");

            //// Rename the TI Name
            //if (GetSingleColumnValues("LoginGrid", "Test Item Name").Contains(tiName))
            //    Assert.Fail($"Same TI existed with TI Name: {tiName}");

            IElement SaveTIWindow = PP5IDEWindow.GetExtendedElementBySingleWithRetry(PP5By.Name("Save Test Item"));

            SaveTIWindow.GetEditElement("TINameTxtBox").SendText(tiName);

            // Change the group
            //WindowsElement comboBox = GetComboBoxElementByID("GroupCmb");
            //if (CheckComboBoxHasItemByName(comboBox, group))
            //    ComboBoxSelectByName(comboBox, group);
            ComboBoxSelectByName("GroupCmb", group);

            // LeftClick on Button "Ok"
            SaveTIWindow.GetBtnElement("Ok").LeftClick();
        }

        public void SetTIActive(string TIName, bool IsActivate)
        {
            SwitchToModule(WindowType.Management);
            SearchTI(TIName);

            if (IsActivate)
                PP5IDEWindow.PerformClick("/ByCell@TestItem_DataGrid[3]/CheckBox[0]", ClickType.TickCheckBox);
            else
                PP5IDEWindow.PerformClick("/ByCell@TestItem_DataGrid[3]/CheckBox[0]", ClickType.UnTickCheckBox);
        }

        public void TIModifyRemark(string TIName, string Remark)
        {
            SwitchToModule(WindowType.Management);
            SearchTI(TIName);

            PP5IDEWindow.PerformInput("/ByCell@TestItem_DataGrid[5]", InputType.SendContent, Remark);
        }

        public void TIRename(string TIName, string newTIName)
        {
            SwitchToModule(WindowType.Management);
            SearchTI(TIName);

            PP5IDEWindow.PerformClick("/ByName[Rename]", ClickType.LeftClick);
            AutoUIExecutor.SwitchTo(SessionType.Desktop);
            CurrentDriver.PerformInput("/Window[Rename Test Item]/Edit[NewName]", InputType.SendSingleKeys, newTIName);
            CurrentDriver.PerformClick("/Window[Rename Test Item]/ByName[OK]", ClickType.LeftClick);
            AutoUIExecutor.SwitchTo(SessionType.PP5IDE);
            WaitUntil(() => PP5IDEWindow.PerformGetElement("/ByCell@TestItem_DataGrid[1]").Text == newTIName);
            //if (CheckMessageBoxOpened(PP5By.Id("MessageBoxExDialog")))
            //{
            //CurrentDriver.PerformClick("/ById[MessageBoxExDialog]/ByName[Yes]", ClickType.LeftClick);
            //}
            //TIFilePath = Path.Combine(Directory.GetParent(TIFilePath).FullName, newTIName + Path.GetExtension(TIFilePath));
        }

        public void SearchTI(string TIName)
        {
            PP5IDEWindow.PerformClick("/ByCondition#ToolBar[Visible]/RadioButton[1]", ClickType.LeftClick); // Click on TP/TI button

            IElement mainTab = PP5IDEWindow.PerformGetElement("/Tab[mainTab]");                             // Click on Test Item tab
            IElement TestItemTabItem = mainTab.TabSelect(1, "Test Item");

            mainTab.TabSelect(1, "Test Program");                                                           // Switch to TP then back to TI (Since search TI will failed)
            mainTab.TabSelect(1, "Test Item");

            TestItemTabItem.PerformInput("/ByIdOrName[searchText]", InputType.SendContent, TIName);         // Search the TI with name
            Press(Keys.Enter);
        }

        public void TIExecuteAction(TIAction tiAction, string TIName)
        {
            switch (tiAction)
            {
                case TIAction.SetTIActive:
                    SetTIActive(TIName, IsActivate: true);
                    break;
                case TIAction.SetTIInactive:
                    SetTIActive(TIName, IsActivate: false);
                    break;
            }
        }

        public void TPExecuteAction(TPAction tpAction)
        {
            switch (tpAction)
            {
                case TPAction.SwitchToPreTestPage:
                    TestProgramTestTypeNavi(TestItemRunType.Pre);
                    break;
                case TPAction.SwitchToUUTTestPage:
                    TestProgramTestTypeNavi(TestItemRunType.UUT);
                    break;
                case TPAction.SwitchToPostTestPage:
                    TestProgramTestTypeNavi(TestItemRunType.Post);
                    break;
                case TPAction.SwitchToSystemTIPage:
                    TestItemListNavi(TestItemSourceType.System);
                    break;
                case TPAction.SwitchToUserDefinedTIPage:
                    TestItemListNavi(TestItemSourceType.UserDefined);
                    break;
                case TPAction.SwitchToTestConditionVariablePage:
                    TestProgramAllSettingPageNavi(TestProgramSettingTabType.TestItem);
                    TestProgramParamInfoNavi(TestProgramParameterTabType.Condition);
                    break;
                case TPAction.SwitchToVectorVariablePage:
                    TestProgramAllSettingPageNavi(TestProgramSettingTabType.TestItem);
                    TestProgramParamInfoNavi(TestProgramParameterTabType.Vector);
                    break;
                case TPAction.SwitchToGlobalVariablePage:
                    TestProgramAllSettingPageNavi(TestProgramSettingTabType.TestItem);
                    TestProgramParamInfoNavi(TestProgramParameterTabType.Global);
                    break;
                case TPAction.SwitchToResultVariablePage:
                    TestProgramAllSettingPageNavi(TestProgramSettingTabType.TestItem);
                    TestProgramParamInfoNavi(TestProgramParameterTabType.Result);
                    break;
                case TPAction.SwitchToTPInfoPage:
                    TestProgramAllSettingPageNavi(TestProgramSettingTabType.TPInfo);
                    break;
                case TPAction.SwitchToReportFormatByTIPage:
                    TestProgramAllSettingPageNavi(TestProgramSettingTabType.RptFormat);
                    ReportFormatNavi(ReportFormatTabType.ByTI);
                    break;
                case TPAction.SwitchToReportFormatByTPPage:
                    TestProgramAllSettingPageNavi(TestProgramSettingTabType.RptFormat);
                    ReportFormatNavi(ReportFormatTabType.ByTP);
                    break;
                default:
                    TestProgramTestTypeNavi(TestItemRunType.UUT);
                    break;
            }
        }

        public void VariableTabNavi(VariableTabType tabType)
        {
            IWebElement ele;
            switch (tabType)
            {
                case VariableTabType.Condition:
                    ele = PP5IDEWindow.GetWebElementFromWebElement(MobileBy.AccessibilityId("CndRdoBtn"));
                    break;
                case VariableTabType.Result:
                    ele = PP5IDEWindow.GetWebElementFromWebElement(MobileBy.AccessibilityId("RstRdoBtn"));
                    break;
                case VariableTabType.Temp:
                    ele = PP5IDEWindow.GetWebElementFromWebElement(MobileBy.AccessibilityId("TmpRdoBtn"));
                    break;
                case VariableTabType.Global:
                    ele = PP5IDEWindow.GetWebElementFromWebElement(MobileBy.AccessibilityId("GlbRdoBtn"));
                    break;
                case VariableTabType.DefectCode:
                    ele = PP5IDEWindow.GetWebElementFromWebElement(MobileBy.AccessibilityId("DftRdoBtn"));
                    break;
                default:
                    ele = PP5IDEWindow.GetWebElementFromWebElement(MobileBy.AccessibilityId("CndRdoBtn"));
                    break;
            }

            if (!ele.Selected)
                ele.LeftClick();
            Assert.IsTrue(ele.Selected);
        }

        public void TestItemTabNavi(TestItemTabType tabType)
        {
            IWebElement ele;
            switch (tabType)
            {
                case TestItemTabType.TIContext:
                    ele = PP5IDEWindow.GetWebElementFromWebElement(MobileBy.AccessibilityId("TIRdoBtn"));
                    break;
                case TestItemTabType.TIDescription:
                    ele = PP5IDEWindow.GetWebElementFromWebElement(MobileBy.AccessibilityId("TIDesRdoBtn"));
                    break;
                default:
                    ele = PP5IDEWindow.GetWebElementFromWebElement(MobileBy.AccessibilityId("TIRdoBtn"));
                    break;
            }

            if (!ele.Selected)
                ele.LeftClick();
            Assert.IsTrue(ele.Selected);
        }

        public void TestItemListNavi(TestItemSourceType tiSType)
        {
            IWebElement ele;
            switch (tiSType)
            {
                case TestItemSourceType.System:
                    ele = PP5IDEWindow.GetWebElementFromWebElement(MobileBy.AccessibilityId("SysRdoBtn"));
                    break;
                case TestItemSourceType.UserDefined:
                    ele = PP5IDEWindow.GetWebElementFromWebElement(MobileBy.AccessibilityId("UserDefRdoBtn"));
                    break;
                default:
                    ele = PP5IDEWindow.GetWebElementFromWebElement(MobileBy.AccessibilityId("SysRdoBtn"));
                    break;
            }

            if (!ele.Selected)
                ele.LeftClick();
            Assert.IsTrue(ele.Selected);
        }

        public void TestProgramTestTypeNavi(TestItemRunType tiRunType)
        {
            IWebElement ele;
            switch (tiRunType)
            {
                case TestItemRunType.Pre:
                    ele = PP5IDEWindow.GetWebElementFromWebElement(MobileBy.AccessibilityId("PreTestRdoBtn"));
                    break;
                case TestItemRunType.UUT:
                    ele = PP5IDEWindow.GetWebElementFromWebElement(MobileBy.AccessibilityId("UUTTestRdoBtn"));
                    break;
                case TestItemRunType.Post:
                    ele = PP5IDEWindow.GetWebElementFromWebElement(MobileBy.AccessibilityId("PostTestRdoBtn"));
                    break;
                default:
                    ele = PP5IDEWindow.GetWebElementFromWebElement(MobileBy.AccessibilityId("UUTTestRdoBtn"));
                    break;
            }

            if (!ele.Selected)
                ele.LeftClick();
            Assert.IsTrue(ele.Selected);
        }

        public void TestProgramAllSettingPageNavi(TestProgramSettingTabType tpSettingType)
        {
            IWebElement ele;
            switch (tpSettingType)
            {
                case TestProgramSettingTabType.TestItem:
                    ele = PP5IDEWindow.GetWebElementFromWebElement(MobileBy.AccessibilityId("TIRdoBtn"));
                    break;
                case TestProgramSettingTabType.TPInfo:
                    ele = PP5IDEWindow.GetWebElementFromWebElement(MobileBy.AccessibilityId("TPInfoRdoBtn"));
                    break;
                case TestProgramSettingTabType.RptFormat:
                    ele = PP5IDEWindow.GetWebElementFromWebElement(MobileBy.AccessibilityId("ReportFormatRdoBtn"));
                    break;
                default:
                    ele = PP5IDEWindow.GetWebElementFromWebElement(MobileBy.AccessibilityId("TIRdoBtn"));
                    break;
            }

            if (!ele.Selected)
                ele.LeftClick();
            Assert.IsTrue(ele.Selected);
        }

        public void TestProgramParamInfoNavi(TestProgramParameterTabType tiParamType)
        {
            IWebElement ele;
            switch (tiParamType)
            {
                case TestProgramParameterTabType.Condition:
                    ele = PP5IDEWindow.GetWebElementFromWebElement(MobileBy.AccessibilityId("ParameterRdoBtn"));
                    break;
                case TestProgramParameterTabType.Vector:
                    ele = PP5IDEWindow.GetWebElementFromWebElement(MobileBy.AccessibilityId("VectorRdoBtn"));
                    break;
                case TestProgramParameterTabType.Global:
                    ele = PP5IDEWindow.GetWebElementFromWebElement(MobileBy.AccessibilityId("GlobalRdoBtn"));
                    break;
                case TestProgramParameterTabType.Result:
                    ele = PP5IDEWindow.GetWebElementFromWebElement(MobileBy.AccessibilityId("ResultRdoBtn"));
                    break;
                default:
                    ele = PP5IDEWindow.GetWebElementFromWebElement(MobileBy.AccessibilityId("ParameterRdoBtn"));
                    break;
            }

            if (!ele.Selected)
                ele.LeftClick();
            Assert.IsTrue(ele.Selected);
        }

        public void ReportFormatNavi(ReportFormatTabType rptFrmtType)
        {
            IWebElement ele;
            switch (rptFrmtType)
            {
                case ReportFormatTabType.ByTI:
                    ele = PP5IDEWindow.GetWebElementFromWebElement(MobileBy.AccessibilityId("ByTIName"));
                    break;
                case ReportFormatTabType.ByTP:
                    ele = PP5IDEWindow.GetWebElementFromWebElement(MobileBy.AccessibilityId("ByTPName"));
                    break;
                default:
                    ele = PP5IDEWindow.GetWebElementFromWebElement(MobileBy.AccessibilityId("ByTIName"));
                    break;
            }

            if (!ele.Selected)
                ele.LeftClick();
            Assert.IsTrue(ele.Selected);
        }

        /// <summary>
        /// For variable: Condition, Global
        /// </summary>
        /// <param name="tabType"></param>
        /// <param name="ShowName"></param>
        /// <param name="CallName"></param>
        /// <param name="DataType">Ex: Float, Integer, etc...</param>
        /// <param name="EditType">Edit, ComboBox, External_Signal</param>
        public void CreateNewVariable1(VariableTabType tabType, string ShowName, string CallName,
                                       VariableDataType DataType, VariableEditType EditType,
                                       OrderedDictionary enumItems, int enumItemSelectionIndex = 0)
        {
            enumItems ??= new OrderedDictionary();

            PP5DataGrid varDataGrid = CreateNewVariable1(tabType, ShowName, CallName, DataType, EditType);

            // If combobox, need to edit enum item, and select on Default
            //bool isVectorType = DataType != VariableDataType.LineInVector
            //                 || DataType != VariableDataType.SpecVector
            //                 || DataType != VariableDataType.ACSpecVector
            //                 || DataType != VariableDataType.ExtMeasVector
            //                 || DataType != VariableDataType.LoadVector
            //                 || DataType != VariableDataType.ACLoadVector
            //                 || DataType != VariableDataType.ConstantVector;
            if (EditType == VariableEditType.ComboBox && !IsVariableVectorDataType(DataType))
            {
                IWebElement enumItemEditorWindow = null;

                // Click in Enum Item cell
                varDataGrid.GetCellBy(varDataGrid.LastRowNo, "Enum Item").LeftClick();

                foreach (DictionaryEntry enumItem in enumItems)
                {
                    enumItemEditorWindow = AddNewEnumItemByNameOrId(enumItem.Key.ToString(), enumItem.Value.ToString());
                }

                // Press OK to finish editing enum items
                enumItemEditorWindow.GetWebElementFromWebElement(PP5By.Name("Ok")).LeftClick();

                // Select enum item in Default cell
                string SelectedEnumValue = enumItems.Count != 0 ? enumItems[enumItemSelectionIndex].ToString() : "";
                if (!SelectedEnumValue.IsEmpty())
                    varDataGrid.GetCellBy(varDataGrid.LastRowNo, "Default").ComboBoxSelectByName(SelectedEnumValue);
            }
        }

        public void CreateNewVariable1(VariableTabType tabType, string ShowName, string CallName,
                                       VariableDataType DataType, VariableEditType EditType, string arrSize1,
                                       OrderedDictionary enumItems, int enumItemSelectionIndex = 0)
        {
            enumItems ??= new OrderedDictionary();

            PP5DataGrid varDataGrid = CreateNewVariable1(tabType, ShowName, CallName, DataType, EditType, arrSize1);

            // If combobox, need to edit enum item, and select on Default
            //bool isVectorType = DataType != VariableDataType.LineInVector
            //                 || DataType != VariableDataType.SpecVector
            //                 || DataType != VariableDataType.ACSpecVector
            //                 || DataType != VariableDataType.ExtMeasVector
            //                 || DataType != VariableDataType.LoadVector
            //                 || DataType != VariableDataType.ACLoadVector
            //                 || DataType != VariableDataType.ConstantVector;
            if (EditType == VariableEditType.ComboBox && !IsVariableVectorDataType(DataType))
            {
                IWebElement enumItemEditorWindow = null;

                // Click in Enum Item cell
                varDataGrid.GetCellBy(varDataGrid.LastRowNo, "Enum Item").LeftClick();

                foreach (DictionaryEntry enumItem in enumItems)
                {
                    enumItemEditorWindow = AddNewEnumItemByNameOrId(enumItem.Key.ToString(), enumItem.Value.ToString());
                }

                // Press OK to finish editing enum items
                enumItemEditorWindow.GetWebElementFromWebElement(PP5By.Name("Ok")).LeftClick();

                // Select enum item in Default cell
                string SelectedEnumValue = enumItems.Count != 0 ? enumItems[enumItemSelectionIndex].ToString() : "";
                if (!SelectedEnumValue.IsEmpty())
                    varDataGrid.GetCellBy(varDataGrid.LastRowNo, "Default").ComboBoxSelectByName(SelectedEnumValue);
            }
        }

        public void CreateNewVariable1(VariableTabType tabType, string ShowName, string CallName,
                               VariableDataType DataType, VariableEditType EditType, string arrSize1, string arrSize2,
                               OrderedDictionary enumItems, int enumItemSelectionIndex = 0)
        {
            enumItems ??= new OrderedDictionary();

            PP5DataGrid varDataGrid = CreateNewVariable1(tabType, ShowName, CallName, DataType, EditType, arrSize1, arrSize2);

            // If combobox, need to edit enum item, and select on Default
            if (EditType == VariableEditType.ComboBox && !IsVariableVectorDataType(DataType))
            {
                IWebElement enumItemEditorWindow = null;

                // Click in Enum Item cell
                varDataGrid.GetCellBy(varDataGrid.LastRowNo, "Enum Item").LeftClick();

                foreach (DictionaryEntry enumItem in enumItems)
                {
                    enumItemEditorWindow = AddNewEnumItemByNameOrId(enumItem.Key.ToString(), enumItem.Value.ToString());
                }

                // Press OK to finish editing enum items
                enumItemEditorWindow.GetWebElementFromWebElement(By.Name("Ok")).LeftClick();

                // Select enum item in Default cell
                string SelectedEnumValue = enumItems.Count != 0 ? enumItems[enumItemSelectionIndex].ToString() : "";
                if (!SelectedEnumValue.IsEmpty())
                    varDataGrid.GetCellBy(varDataGrid.LastRowNo, "Default").ComboBoxSelectByName(SelectedEnumValue);
                //Press(Keys.Enter);
            }
        }

        public PP5DataGrid CreateNewVariable1(VariableTabType tabType, string ShowName, string CallName,
                               VariableDataType DataType, VariableEditType EditType)
        {
            PP5DataGrid varDataGrid = CreateNewVariable1(tabType, ShowName, CallName, DataType);

            // Select "ComboBox" in Edit Type combobox
            varDataGrid.GetCellBy(varDataGrid.LastRowNo, "Edit Type").ComboBoxSelectByName(EditType.ToString());
            return varDataGrid;
        }

        public PP5DataGrid CreateNewVariable1(VariableTabType tabType, string ShowName, string CallName,
                               VariableDataType DataType, VariableEditType EditType, string arrSize1)
        {
            PP5DataGrid varDataGrid = CreateNewVariable1(tabType, ShowName, CallName, DataType, arrSize1);

            // Select "ComboBox" in Edit Type combobox
            varDataGrid.GetCellBy(varDataGrid.LastRowNo, "Edit Type").ComboBoxSelectByName(EditType.ToString());
            return varDataGrid;
        }

        public PP5DataGrid CreateNewVariable1(VariableTabType tabType, string ShowName, string CallName,
                                       VariableDataType DataType, VariableEditType EditType, string arrSize1, string arrSize2)
        {
            PP5DataGrid varDataGrid = CreateNewVariable1(tabType, ShowName, CallName, DataType, arrSize1, arrSize2);

            // Select "ComboBox" in Edit Type combobox
            varDataGrid.GetCellBy(varDataGrid.LastRowNo, "Edit Type").ComboBoxSelectByName(EditType.ToString());
            return varDataGrid;
        }

        public PP5DataGrid CreateNewVariable1(VariableTabType tabType, string ShowName, string CallName, VariableDataType DataType, string arrSize1 = "10", string arrSize2 = "10")
        {
            PP5DataGrid varDataGrid = CreateNewVariable1(tabType, ShowName, CallName);

            // Input "Float" in Data Type cell
            varDataGrid.GetCellBy(varDataGrid.LastRowNo, "Data Type").ComboBoxSelectByName(DataType.GetDescription());

            if (IsVariable1DArrayDataType(DataType))
                VariableTypeSizeSelect1DArray(arrSize1);
            else if (IsVariable2DArrayDataType(DataType))
                VariableTypeSizeSelect2DArray(arrSize1, arrSize2);

            return varDataGrid;
        }

        public void VariableTypeSizeSelect1DArray(string arrSize1)
        {
            //PP5DataGrid varDataGrid = CreateNewVariable1(tabType, ShowName, CallName, DataType);
            //VariableDataType DataType1DArray = VariableDataType.FloatArray | VariableDataType.IntegerArray | 
            //                                   VariableDataType.DoubleArray | VariableDataType.StringArray | 
            //                                   VariableDataType.ByteArray | VariableDataType.HexStringArray | VariableDataType.LongArray;
            //if (DataType == (DataType & DataType1DArray))
            //if (IsVariable1DArrayDataType(DataType))
            //{
            PP5IDEWindow.PerformInput("/ByName[Array Size Editor]/ById[SizeTxtBox]", InputType.SendContent, arrSize1);
            PP5IDEWindow.PerformClick("/ByName[Array Size Editor,Ok]", ClickType.LeftClick);
            //}
            //return varDataGrid;
        }

        public void VariableTypeSizeSelect2DArray(string arrSize1, string arrSize2)
        {
            //PP5DataGrid varDataGrid = CreateNewVariable1(tabType, ShowName, CallName, DataType);
            //VariableDataType DataType2DArray = VariableDataType.Float2DArray | VariableDataType.Integer2DArray |
            //                                   VariableDataType.Double2DArray | VariableDataType.String2DArray |
            //                                   VariableDataType.Byte2DArray | VariableDataType.HexString2DArray | VariableDataType.Long2DArray;
            //if (DataType == (DataType & DataType2DArray))
            //if (IsVariable2DArrayDataType(DataType))
            //{
            if (int.Parse(arrSize1) * int.Parse(arrSize2) > 100000)
            {
                ((PP5Element)PP5IDEWindow).ErrorMessages.Add("/Window[0]/Edit[0]", PP5IDEWindow.PerformGetElement("/Window[0]/Edit[0]").GetToolTipMessage());
                ((PP5Element)PP5IDEWindow).ErrorMessages.Add("/Window[0]/Edit[1]", PP5IDEWindow.PerformGetElement("/Window[0]/Edit[1]").GetToolTipMessage());
                return;
            }
            PP5IDEWindow.PerformInput("/Window[0]/Edit[0]", InputType.SendContent, arrSize1);
            PP5IDEWindow.PerformInput("/Window[0]/Edit[1]", InputType.SendContent, arrSize2);
            PP5IDEWindow.PerformClick("/Window[0]/Button[Ok]", ClickType.LeftClick);
            //}
            //return varDataGrid;
        }

        public PP5DataGrid CreateNewVariable1(VariableTabType tabType, string ShowName, string CallName)
        {
            PP5DataGrid varDataGrid = CreateNewVariable1(tabType, ShowName);

            // Input "a" in Call Name cell
            varDataGrid.GetSelectedCellBy(varDataGrid.LastRowNo, "Call Name").SendText(CallName);
            return varDataGrid;
        }

        public PP5DataGrid CreateNewVariable1(VariableTabType tabType, string ShowName)
        {
            VariableTabNavi(tabType);

            //// Testing get data table element from cache
            // Press page down until last row show up
            PP5DataGrid varDataGrid = new PP5DataGrid((PP5Element)GetDataTableElement((DataTableAutoIDType)tabType));
            //PP5DataGrid varDataGrid = (PP5DataGrid)varDataGridEle;
            //int rowCount = GetRowCount((DataTableAutoIDType)tabType);

            // First Add a new empty row
            //varDataGrid.RefreshSelectedCell();
            varDataGrid.GetCellBy(varDataGrid.LastRowNo + 1, "Show Name").DoubleClick();
            //varDataGrid.RefreshSelectedCell();

            // 20241007, Adam fix the row count > max display count case
            //Press(Keys.Up);
            //if (varDataGrid.RowCount == varDataGrid.LastRowNo)
            //    varDataGrid.LastRowNo--;

            // Input "a" in Show Name cell
            if (!ShowName.IsNullOrEmpty())
                varDataGrid.GetSelectedCellBy(varDataGrid.LastRowNo, "Show Name").SendText(ShowName);

            return varDataGrid;
        }

        public PP5DataGrid InitializeVariableDataGrid(VariableTabType tabType, string ShowName)
        {
            VariableTabNavi(tabType);

            //// Testing get data table element from cache
            // Press page down until last row show up
            PP5DataGrid varDataGrid = new PP5DataGrid((PP5Element)GetDataTableElement((DataTableAutoIDType)tabType));
            varDataGrid.GetCellBy(varDataGrid.LastRowNo + 1, "No").LeftClick();

            Press(Keys.Tab);
            varDataGrid.RefreshSelectedCell();
            varDataGrid.SelectedCellInfo.SelectedCell.DoubleClick();
            Press(Keys.Up);
            if (!ShowName.IsNullOrEmpty())
                SendSingleKeys(ShowName);

            return varDataGrid;
        }

        public PP5DataGrid InitializeVariableDataGrid(VariableTabType tabType, string ShowName, string CallName)
        {
            PP5DataGrid varDataGrid = InitializeVariableDataGrid(tabType, ShowName);

            Press(Keys.Tab);
            if (!CallName.IsNullOrEmpty())
                SendSingleKeys(CallName);
            return varDataGrid;
        }

        public PP5DataGrid InitializeVariableDataGrid(VariableTabType tabType, string ShowName, string CallName,
                                               VariableDataType DataType, string arrSize1 = "10", string arrSize2 = "10")
        {
            PP5DataGrid varDataGrid = InitializeVariableDataGrid(tabType, ShowName, CallName);

            // Input "Float" in Data Type cell
            Press(Keys.Tab);
            SendSingleKeys(DataType.GetDescription());
            Press(Keys.Enter);

            if (IsVariable1DArrayDataType(DataType))
                VariableTypeSizeSelect1DArray(arrSize1);
            else if (IsVariable2DArrayDataType(DataType))
                VariableTypeSizeSelect2DArray(arrSize1, arrSize2);

            RecordVariableInfo(varDataGrid, new VariableInfo(tabType, CallName, DataType));
            return varDataGrid;
        }

        public PP5DataGrid InitializeVariableDataGrid(VariableTabType tabType, string ShowName, string CallName,
                                               VariableDataType DataType, VariableEditType EditType)
        {
            PP5DataGrid varDataGrid = InitializeVariableDataGrid(tabType, ShowName, CallName, DataType);

            // Select "ComboBox" in Edit Type combobox
            Press(Keys.Tab);
            SendSingleKeys(EditType.ToString());
            Press(Keys.Enter);

            RecordVariableInfo(varDataGrid, new VariableInfo(tabType, CallName, DataType, EditType));
            return varDataGrid;
        }

        public PP5DataGrid InitializeVariableDataGrid(VariableTabType tabType, string ShowName, string CallName,
                                               VariableDataType DataType, VariableEditType EditType, string arrSize1)
        {
            PP5DataGrid varDataGrid = InitializeVariableDataGrid(tabType, ShowName, CallName, DataType, arrSize1);

            Press(Keys.Tab);
            SendSingleKeys(EditType.ToString());
            Press(Keys.Enter);

            RecordVariableInfo(varDataGrid, new VariableInfo(tabType, CallName, DataType, EditType, int.Parse(arrSize1)));
            return varDataGrid;
        }

        public PP5DataGrid InitializeVariableDataGrid(VariableTabType tabType, string ShowName, string CallName,
                                                      VariableDataType DataType, VariableEditType EditType, string arrSize1, string arrSize2)
        {
            PP5DataGrid varDataGrid = InitializeVariableDataGrid(tabType, ShowName, CallName, DataType, arrSize1, arrSize2);

            Press(Keys.Tab);
            SendSingleKeys(EditType.ToString());
            Press(Keys.Enter);

            RecordVariableInfo(varDataGrid, new VariableInfo(tabType, CallName, DataType, EditType, int.Parse(arrSize1), int.Parse(arrSize2)));
            return varDataGrid;
        }

        public PP5DataGrid CreateNewVariableWithMinValue(VariableTabType tabType, string showName, string callName, VariableDataType dataType, VariableEditType editType,
                                                         string minValue, string arrSize1 = "10", string arrSize2 = "10")
        {
            return CreateNewVariableWithBoundaryValue(tabType, showName, callName, dataType, editType, minValue, arrSize1, arrSize2, VariableColumnType.Min, VariableCheckCanAddMinValue);
        }

        public PP5DataGrid CreateNewVariableWithMaxValue(VariableTabType tabType, string showName, string callName, VariableDataType dataType, VariableEditType editType,
                                                         string maxValue, string arrSize1 = "10", string arrSize2 = "10")
        {
            PP5DataGrid varDataGrid = CreateNewVariableWithBoundaryValue(tabType, showName, callName, dataType, editType, maxValue, arrSize1, arrSize2, VariableColumnType.Max, VariableCheckCanAddMaxValue);

            // Additional Operation: Refresh selected cell and perform click
            varDataGrid.RefreshSelectedCell();
            var selectedCell = varDataGrid.SelectedCellInfo.SelectedCell;
            selectedCell.MoveToElementAndClick(0, selectedCell.Height, MoveToElementOffsetOrigin.Center);

            return varDataGrid;
        }

        public PP5DataGrid CreateNewVariableWithDefaultValue(VariableTabType tabType, string showName, string callName, VariableDataType dataType, VariableEditType editType,
                                                             string defaultValue, string arrSize1 = "10", string arrSize2 = "10")
        {
            // Step 1: Initialization and Validation
            PP5DataGrid varDataGrid = InitializeAndValidateVariable(tabType, showName, callName, dataType, editType, arrSize1, arrSize2,
                                                                    (value, arr1, arr2, type, edit) => VariableCheckCanAddDefaultValue(value, arr1, arr2, type, edit),
                                                                    defaultValue, arrSize1, arrSize2,
                                                                    "Can't setup Default value with given parameters");

            // Step 2: Modify Default Value if HexString
            if (IsVariableHexStringDataType(dataType))
            {
                defaultValue = string.Concat(Enumerable.Repeat(defaultValue, int.Parse(varDataGrid.GetCellValue(1, "Max."))));
            }

            // Step 3: Add Default Value
            VariableSelectionMoveTo(tabType, dataType, editType, VariableColumnType.EditType, VariableColumnType.Default);
            VariableInfo variableInfo = new VariableInfo(tabType, callName, dataType, editType);
            AddValueToDataGrid(varDataGrid, variableInfo, VariableColumnType.Default, defaultValue);

            return varDataGrid;
        }

        public PP5DataGrid CreateNewVariableWithFormat(VariableTabType tabType, string showName, string callName, VariableDataType dataType, VariableEditType editType,
                                                       VariableDecimalPlaces[] formatArr, string arrSize1 = "10", string arrSize2 = "10")
        {
            // Step 1: Initialization and Validation
            PP5DataGrid varDataGrid = InitializeAndValidateVariable(tabType, showName, callName, dataType, editType, arrSize1, arrSize2,
                                                                    dataType => IsVariableFloatType(dataType),
                                                                    "Only \"Float\" related datatype is supported!");

            // Step 2: Actual Operation
            VariableSelectionMoveTo(tabType, dataType, editType, VariableColumnType.EditType, VariableColumnType.Format);
            ApplyFormatToDataGrid(varDataGrid, formatArr);

            return varDataGrid;
        }

        /// <summary>
        /// For variable: Result, Temp
        /// </summary>
        /// <param name="tabType"></param>
        /// <param name="ShowName"></param>
        /// <param name="CallName"></param>
        /// <param name="DataType"></param>
        /// <param name="arrSize1"></param>
        /// <param name="arrSize2"></param>
        /// <returns></returns>
        public PP5DataGrid CreateNewVariable2(VariableTabType tabType, string ShowName, string CallName, VariableDataType DataType, string arrSize1 = "10", string arrSize2 = "10")
        {
            PP5DataGrid varDataGrid = CreateNewVariable2(tabType, ShowName, CallName);

            // Input "Float" in Data Type cell
            varDataGrid.GetCellBy(varDataGrid.LastRowNo, "Data Type").ComboBoxSelectByName(DataType.GetDescription());
            //Press(Keys.Enter);

            if (IsVariable1DArrayDataType(DataType))
                VariableTypeSizeSelect1DArray(arrSize1);
            else if (IsVariable2DArrayDataType(DataType))
                VariableTypeSizeSelect2DArray(arrSize1, arrSize2);

            return varDataGrid;
        }

        public PP5DataGrid CreateNewVariable2(VariableTabType tabType, string ShowName, string CallName)
        {
            PP5DataGrid varDataGrid = CreateNewVariable2(tabType, ShowName);

            // Input "a" in Call Name cell
            //GetCellBy("CndGrid", 0, "Call Name").LeftClick();
            varDataGrid.GetSelectedCellBy(varDataGrid.LastRowNo, "Call Name").SendText(CallName);
            return varDataGrid;
        }

        public PP5DataGrid CreateNewVariable2(VariableTabType tabType, string ShowName)
        {
            VariableTabNavi(tabType);

            //// Testing get data table element from cache
            // Press page down until last row show up
            PP5DataGrid varDataGrid = new PP5DataGrid((PP5Element)GetDataTableElement((DataTableAutoIDType)tabType));
            //varDataGrid.LastRowNo = varDataGrid.RowCount;
            //if (varDataGrid.LastRowNo > 1 && varDataGrid.GetAttribute("Scroll.VerticallyScrollable") == bool.TrueString
            //                 && varDataGrid.GetAttribute("Scroll.VerticalScrollPercent") != "100")
            //{
            //    varDataGrid.GetCellBy(1, "No").LeftClick();
            //    while (double.Parse(varDataGrid.GetAttribute("Scroll.VerticalScrollPercent")) <= 99.99)
            //        Press(Keys.PageDown);
            //}

            // First Add a new empty row
            //varDataGrid.GetCellBy(varDataGrid.LastRowNo, "Show Name").DoubleClick();
            varDataGrid.GetCellBy(varDataGrid.LastRowNo + 1, "Show Name").DoubleClick();

            // 20241007, Adam fix the row count > max display count case
            //Press(Keys.Up);
            //if (varDataGrid.GetRowCount() == varDataGrid.LastRowNo)
            //    varDataGrid.LastRowNo--;

            // Input "a" in Show Name cell
            //GetCellBy("CndGrid", 0, "Show Name").LeftClick();
            varDataGrid.GetSelectedCellBy(varDataGrid.LastRowNo, "Show Name").SendText(ShowName);
            return varDataGrid;
        }

        private bool VariableCheckCanAddMinValue(string min, int arrSize1, int arrSize2, VariableDataType DataType, VariableEditType EditType)
        {
            //bool needToCheckMinValue = false;
            bool minValueIsValid = false;

            // Skip for vector type, string/hexString/byte related types, only supported in setting default value
            if (min.IsNullOrEmptyOrWhiteSpace() || IsVariableVectorDataType(DataType) || IsVariableStringDataType(DataType) ||
                IsVariableHexStringDataType(DataType) || IsVariableByteArrayDataType(DataType)) { return false; }

            if (EditType == VariableEditType.EditBox)
            {
                if (IsVariableArrayDataType(DataType))
                {
                    if (IsVariableUUTArrayDataType(DataType))
                        minValueIsValid = CheckValueTypeValid(min, DataType, 640);
                    else if (IsVariable1DArrayDataType(DataType))
                        minValueIsValid = CheckValueTypeValid(min, DataType, arrSize1);
                    else if (IsVariable2DArrayDataType(DataType))
                        minValueIsValid = CheckValueTypeValid(min, DataType, arrSize1, arrSize2);
                }
                else
                    // Regular data type (Float, Integer, Long,...)
                    minValueIsValid = CheckValueTypeValid(min, DataType);
            }

            return minValueIsValid;
        }

        private bool VariableCheckCanAddMaxValue(string max, int arrSize1, int arrSize2, VariableDataType DataType, VariableEditType EditType)
        {
            //bool needToCheckMinValue = false;
            bool minValueIsValid = false;

            // Skip for vector type, string/byte related types, byte only supported in setting default value
            if (max.IsNullOrEmptyOrWhiteSpace() || IsVariableVectorDataType(DataType) || IsVariableStringDataType(DataType) || IsVariableByteArrayDataType(DataType)) { return false; }

            if (EditType == VariableEditType.EditBox)
            {
                if (IsVariableHexStringDataType(DataType))
                    minValueIsValid = CheckValueTypeValid(max, DataType);
                else if (IsVariableArrayDataType(DataType))
                {
                    if (IsVariableUUTArrayDataType(DataType))
                        minValueIsValid = CheckValueTypeValid(max, DataType, 640);
                    else if (IsVariable1DArrayDataType(DataType))
                        minValueIsValid = CheckValueTypeValid(max, DataType, arrSize1);
                    else if (IsVariable2DArrayDataType(DataType))
                        minValueIsValid = CheckValueTypeValid(max, DataType, arrSize1, arrSize2);
                }
                else
                    // Regular data type (Float, Integer, Long,...)
                    minValueIsValid = CheckValueTypeValid(max, DataType);
            }

            return minValueIsValid;
        }

        private bool VariableCheckCanAddDefaultValue(string @default, int arrSize1, int arrSize2, VariableDataType DataType, VariableEditType EditType)
        {
            //bool needToCheckMinValue = false;
            bool DefaultValueIsValid = false;

            // Skip for vector type, string/byte related types, byte only supported in setting default value
            if (@default.IsNullOrEmptyOrWhiteSpace()) { return false; }

            if (EditType == VariableEditType.EditBox)
            {
                if (IsVariableHexStringDataType(DataType))
                    DefaultValueIsValid = CheckValueTypeValid(@default, DataType);
                else if (IsVariableArrayDataType(DataType))
                {
                    if (IsVariableUUTArrayDataType(DataType))
                        DefaultValueIsValid = CheckValueTypeValid(@default, DataType, 640);
                    else if (IsVariable1DArrayDataType(DataType))
                        DefaultValueIsValid = CheckValueTypeValid(@default, DataType, arrSize1);
                    else if (IsVariable2DArrayDataType(DataType))
                        DefaultValueIsValid = CheckValueTypeValid(@default, DataType, arrSize1, arrSize2);
                }
                else
                    // Regular data type (Float, Integer, Long,...)
                    DefaultValueIsValid = CheckValueTypeValid(@default, DataType);
            }
            else if (EditType == VariableEditType.External_Signal)
            {
                // For String type only, including String/StringUUTArray/String1DArray/String2DArray
                StaticData.currentDataGrid.RefreshSelectedCell();
                StaticData.currentDataGrid.SelectedCellInfo.SelectedCell.LeftClick();
                return PP5IDEWindow.PerformGetElement("/ByClass[Popup]/ById[ItemTreeViewUserControl,tvCluster]") != null;
            }
            else // ComboBox
            {
                // Vector type, value is the CallName
                if (IsVariableVectorDataType(DataType))
                {
                    PP5DataGrid VectorDataGrid = new PP5DataGrid((PP5Element)PP5IDEWindow.PerformGetElement("/ById[VectorGrid]"));
                    //StaticData.currentDataGrid = VectorDataGrid;
                    return VectorDataGrid != null;
                }
                else
                {
                    // Regular setup default from Enum Items
                }
            }

            return DefaultValueIsValid;
        }

        private bool VariableCheckCanAddFormatValue(string format, int arrSize1, int arrSize2, VariableDataType DataType, VariableEditType EditType)
        {
            if (!IsVariableFloatType(DataType))
            {
                Logger.LogMessage("Only \"Float\" related datatype is supported!");
                return false;
            }
            return true;
        }

        public ValidateVariableValue validateVariable;
        public delegate bool ValidateVariableValue(VariableInfo variableInfo, object value);

        public bool VariableCheckCanAddMinValue(VariableInfo variableInfo, object min)
        {
            //bool needToCheckMinValue = false;
            bool minValueIsValid = false;
            string minStr = min.ToString();

            // Skip for vector type, string/hexString/byte related types, only supported in setting default value
            if (minStr.IsNullOrEmptyOrWhiteSpace() || IsVariableVectorDataType(variableInfo.DataType) || IsVariableStringDataType(variableInfo.DataType) ||
                IsVariableHexStringDataType(variableInfo.DataType) || IsVariableByteArrayDataType(variableInfo.DataType)) { return false; }

            if (variableInfo.EditType == VariableEditType.EditBox)
            {
                if (IsVariableArrayDataType(variableInfo.DataType))
                {
                    if (IsVariableUUTArrayDataType(variableInfo.DataType))
                        minValueIsValid = CheckValueTypeValid(minStr, variableInfo.DataType, 640);
                    else if (IsVariable1DArrayDataType(variableInfo.DataType))
                        minValueIsValid = CheckValueTypeValid(minStr, variableInfo.DataType, variableInfo.ArrSize1);
                    else if (IsVariable2DArrayDataType(variableInfo.DataType))
                        minValueIsValid = CheckValueTypeValid(minStr, variableInfo.DataType, variableInfo.ArrSize1, variableInfo.ArrSize2);
                }
                else
                    // Regular data type (Float, Integer, Long,...)
                    minValueIsValid = CheckValueTypeValid(minStr, variableInfo.DataType);
            }

            return minValueIsValid;
        }

        public bool VariableCheckCanAddMaxValue(VariableInfo variableInfo, object max)
        {
            //bool needToCheckMinValue = false;
            bool minValueIsValid = false;
            string maxStr = max.ToString();

            // Skip for vector type, string/byte related types, byte only supported in setting default value
            if (maxStr.IsNullOrEmptyOrWhiteSpace() || IsVariableVectorDataType(variableInfo.DataType) || IsVariableStringDataType(variableInfo.DataType) || IsVariableByteArrayDataType(variableInfo.DataType)) { return false; }

            if (variableInfo.EditType == VariableEditType.EditBox)
            {
                if (IsVariableHexStringDataType(variableInfo.DataType))
                    minValueIsValid = CheckValueTypeValid(maxStr, variableInfo.DataType);
                else if (IsVariableArrayDataType(variableInfo.DataType))
                {
                    if (IsVariableUUTArrayDataType(variableInfo.DataType))
                        minValueIsValid = CheckValueTypeValid(maxStr, variableInfo.DataType, 640);
                    else if (IsVariable1DArrayDataType(variableInfo.DataType))
                        minValueIsValid = CheckValueTypeValid(maxStr, variableInfo.DataType, variableInfo.ArrSize1);
                    else if (IsVariable2DArrayDataType(variableInfo.DataType))
                        minValueIsValid = CheckValueTypeValid(maxStr, variableInfo.DataType, variableInfo.ArrSize1, variableInfo.ArrSize2);
                }
                else
                    // Regular data type (Float, Integer, Long,...)
                    minValueIsValid = CheckValueTypeValid(maxStr, variableInfo.DataType);
            }

            return minValueIsValid;
        }

        public bool VariableCheckCanAddDefaultValue(VariableInfo variableInfo, object @default)
        {
            //bool needToCheckMinValue = false;
            bool DefaultValueIsValid = false;
            string @defaultStr = @default.ToString();

            // Skip for vector type, string/byte related types, byte only supported in setting default value
            if (@defaultStr.ToString().IsNullOrEmptyOrWhiteSpace()) { return false; }

            if (variableInfo.EditType == VariableEditType.EditBox)
            {
                if (IsVariableHexStringDataType(variableInfo.DataType))
                    DefaultValueIsValid = CheckValueTypeValid(@defaultStr.ToString(), variableInfo.DataType);
                else if (IsVariableArrayDataType(variableInfo.DataType))
                {
                    if (IsVariableUUTArrayDataType(variableInfo.DataType))
                        DefaultValueIsValid = CheckValueTypeValid(@defaultStr.ToString(), variableInfo.DataType, 640);
                    else if (IsVariable1DArrayDataType(variableInfo.DataType))
                        DefaultValueIsValid = CheckValueTypeValid(@defaultStr.ToString(), variableInfo.DataType, variableInfo.ArrSize1);
                    else if (IsVariable2DArrayDataType(variableInfo.DataType))
                        DefaultValueIsValid = CheckValueTypeValid(@defaultStr.ToString(), variableInfo.DataType, variableInfo.ArrSize1, variableInfo.ArrSize2);
                }
                else
                    // Regular data type (Float, Integer, Long,...)
                    DefaultValueIsValid = CheckValueTypeValid(@defaultStr.ToString(), variableInfo.DataType);
            }
            else if (variableInfo.EditType == VariableEditType.External_Signal)
            {
                // For String type only, including String/StringUUTArray/String1DArray/String2DArray
                StaticData.currentDataGrid.RefreshSelectedCell();
                StaticData.currentDataGrid.SelectedCellInfo.SelectedCell.LeftClick();
                return PP5IDEWindow.PerformGetElement("/ByClass[Popup]/ById[ItemTreeViewUserControl,tvCluster]") != null;
            }
            else // ComboBox
            {
                // Vector type, value is the CallName
                if (IsVariableVectorDataType(variableInfo.DataType))
                {
                    //PP5DataGrid VectorDataGrid = new PP5DataGrid((PP5Element)PP5IDEWindow.PerformGetElement("/ById[VectorGrid]"));
                    //StaticData.currentDataGrid = VectorDataGrid;
                    return PP5IDEWindow.PerformGetElement("/ById[VectorGrid]") != null;
                }
                else
                {
                    // Regular setup default from Enum Items
                }
            }

            return DefaultValueIsValid;
        }

        public bool VariableCheckCanAddFormatValue(VariableInfo variableInfo, object format)
        {
            if (!IsVariableFloatType(variableInfo.DataType))
            {
                Logger.LogMessage("Only \"Float\" related datatype is supported!");
                return false;
            }
            return true;
        }

        public bool VariableCheckCanAddValue(VariableInfo variableInfo, VariableColumnType columnType, object value)
        {
            switch (columnType)
            {
                case VariableColumnType.Min:
                    validateVariable = new ValidateVariableValue(VariableCheckCanAddMinValue);
                    break; 
                case VariableColumnType.Max:
                    validateVariable = new ValidateVariableValue(VariableCheckCanAddMaxValue);
                    break;
                case VariableColumnType.Default:
                    validateVariable = new ValidateVariableValue(VariableCheckCanAddDefaultValue);
                    break;
                case VariableColumnType.Format:
                    validateVariable = new ValidateVariableValue(VariableCheckCanAddFormatValue);
                    break;
            }

            return validateVariable.Invoke(variableInfo, value);
        }

        //public void FillIn(PP5DataGrid varDataGrid, VariableColumnType columnType, string value)
        //{
        //    varDataGrid.RefreshSelectedCell();
        //    VariableColumnType currColType = TypeExtension.GetEnumByDescription<VariableColumnType>(varDataGrid.SelectedCellInfo.ColumnName);


        //}

        //private bool VariableCheckCanAddMaxValue(string max, VariableDataType DataType, VariableEditType EditType)
        //{
        //    return false;
        //}

        //private bool VariableCheckCanAddDefaultValue(string max, VariableDataType DataType, VariableEditType EditType)
        //{
        //    string extSig = max;
        //    if (EditType == VariableEditType.External_Signal)
        //        VariableAddExternalSignal(extSig);

        //    return true;
        //}

        //public void VariableAddExternalSignal(string extSig)
        //{
        //    // extSig format: "PhoenixonDSP.DSPSensingOFF(0xFF).DSPSensingOFF_Signal";
        //    var extSigLabels = extSig.Split('.');
        //    if (extSigLabels.Length > 1)
        //    {
        //        PP5IDEWindow.PerformGetElement("/ById[tvCluster]")
        //                    .AddTreeViewItem(extSigLabels);
        //    }
        //}

        //private bool CheckMinValueTypeValidold(string value, VariableDataType DataType)
        //{
        //    //Float = 1L,
        //    //Integer = 2L,
        //    //FloatPercentage = 4L,
        //    //Short = 8L,
        //    //String = 16L,
        //    //Byte = 32L,
        //    //FloatArrayOfSize640 = 64L,
        //    //IntegerArrayOfSize640 = 128L,
        //    //FloatPercentageArrayOfSize640 = 256L,
        //    //HexString = 512L,
        //    //FloatArray = 1024L,
        //    //IntegerArray = 2048L,
        //    //LineInVector = 4096L,
        //    //LoadVector = 8192L,
        //    //SpecVector = 16384L,
        //    //ExtMeasVector = 32768L,
        //    //ACLoadVector = 65536L,
        //    //ACSpecVector = 131072L,
        //    //ConstantVector = 262144L,
        //    //Double = 524288L,
        //    //DoubleArray = 1048576L,
        //    //DoubleArrayOfUUTMaxSize = 2097152L,
        //    //StringArray = 4194304L,
        //    //ByteArray = 8388608L,
        //    //Double2DArray = 16777216L,
        //    //Float2DArray = 33554432L,
        //    //Integer2DArray = 67108864L,
        //    //Byte2DArray = 134217728L,
        //    //String2DArray = 268435456L,
        //    //StringArrayOfUUTMaxSize = 536870912L,
        //    //Long = 1073741824L,
        //    //LongArray = 2147483648L,
        //    //HexStringArray = 4294967296L,
        //    //HexString2DArray = 8589934592L,
        //    //Long2DArray = 17179869184L,
        //    switch (DataType) 
        //    {
        //        case VariableDataType.Float: return float.TryParse(value, out _);
        //        case VariableDataType.Integer: return int.TryParse(value, out _);
        //        case VariableDataType.FloatPercentage: return float.TryParse(value, out _);
        //        case VariableDataType.Short: return short.TryParse(value, out _);
        //        case VariableDataType.String: return true;
        //        case VariableDataType.Byte: return byte.TryParse(value, out _);
        //        case VariableDataType.Double: return double.TryParse(value, out _);

        //        case VariableDataType.HexString:
        //        case VariableDataType.StringArray:
        //        case VariableDataType.String2DArray:
        //        case VariableDataType.ByteArray:
        //        case VariableDataType.Byte2DArray:
        //        case VariableDataType.LineInVector:
        //        case VariableDataType.LoadVector:
        //        case VariableDataType.SpecVector:
        //        case VariableDataType.ExtMeasVector:
        //        case VariableDataType.ACLoadVector:
        //        case VariableDataType.ACSpecVector:
        //        case VariableDataType.ConstantVector:
        //            return false;

        //        case VariableDataType.FloatArrayOfSize640: return float.TryParse(value, out _);
        //        case VariableDataType.IntegerArrayOfSize640: return int.TryParse(value, out _);
        //        case VariableDataType.FloatPercentageArrayOfSize640: return float.TryParse(value, out _);
        //        case VariableDataType.StringArrayOfUUTMaxSize: return true;
        //        case VariableDataType.DoubleArrayOfUUTMaxSize: return double.TryParse(value, out _);
        //        case VariableDataType.FloatArray: return float.TryParse(value, out _);
        //        case VariableDataType.IntegerArray: return int.TryParse(value, out _);
        //        case VariableDataType.DoubleArray: return double.TryParse(value, out _);
        //        case VariableDataType.Double2DArray: return double.TryParse(value, out _);
        //        case VariableDataType.Float2DArray: return float.TryParse(value, out _);
        //        case VariableDataType.Integer2DArray: return int.TryParse(value, out _);
        //    }
        //}

        /// <summary>
        /// Valid array:
        /// 1. "10,10,10,10,10" === "10[5]"      -> 1D array, rank 5 x 1, with same value "10" (value[rank1])
        /// 2. "10,10;10,10;10,10" === "10[2,3]" -> 2D array, rank 2 x 3, with same value "10" (value[rank1,rank2])
        /// 3. "10,2,5,15,0"                     -> 1D array, rank 5 x 1, with different values
        /// 4. "10,2;5,15;0,100"                 -> 2D array, rank 2 x 3, with different values
        ///
        /// Invalid array:
        /// 1. 維度格式不一致的情況
        ///   > "10;10,10"                       -> 混合使用,和;的分隔符，不符合統一維度的要求
        ///   > "10,10;10"                       -> 10,10;10缺少完整的第二維度
        ///   > "10,10;10,10,10"                 -> 第二維度中的rank不匹配，期望每行應該有相同的數量
        /// 2. 分隔符號不正確
        ///   > "10-10,10"                       -> 使用不正確的分隔符-而不是,或;
        ///   > "10:10;10"                       -> 使用:來分隔，應該是,或;
        /// 7. 格式與預期rank不匹配
        ///   > "10,10,10" === "10[2,3]"         -> 給定的rank為2 x 3，但實際只有3個值
        ///   > "10,10;10,10" === "10[2]"        -> 給定rank為1D的[2]，但實際值看起來是2D的格式
        /// 8. 缺少維度分隔符
        ///   > "10,10 10,10"                    -> 缺少維度分隔符;，應該是10,10;10,10
        /// 10. 無法解析的格式
        ///   > "[10,10,10]"                     -> 包含方括號，應為逗號或分號分隔的數值序列
        ///   > "10;10];[10"                     -> 包含不對稱的括號，無法解析
        ///
        /// Should be valid:
        /// 3. 額外的分隔符號
        ///   > "10,,10,10"                      -> 出現多個連續的,，代表空值
        ///   > "10,10;;10,10"                   -> 出現多個連續的;，代表空值
        /// 4. 缺少值
        ///   > ",10,10"                         -> 首位缺少值。
        ///   > "10,10;"                         -> 最後有一個額外的;，且沒有對應的值。
        ///   > "10,;10,10"                      -> 中間缺少值。
        /// 6. 多餘的空格
        ///   > "10, 10, 10"                     -> 數字之間有額外的空格
        ///   > " 10,10;10,10"                   -> 開頭有多餘的空格
        ///   > "10,10;10,10 "                   -> 結尾有多餘的空格
        /// </summary>
        /// <param name="value"></param>
        /// <param name="DataType"></param>
        /// <returns></returns>

        // General Method for Creating New Variables with Boundary Values
        private PP5DataGrid CreateNewVariableWithBoundaryValue(VariableTabType tabType, string showName, string callName, VariableDataType dataType, VariableEditType editType,
                                                               string value, string arrSize1, string arrSize2, VariableColumnType columnType,
                                                               Func<string, int, int, VariableDataType, VariableEditType, bool> checkCanAddValue)
        {
            // Step 1: Initialization and Validation
            PP5DataGrid varDataGrid = InitializeVariableDataGrid(tabType, showName, callName, dataType, editType, arrSize1, arrSize2);

            if (!checkCanAddValue(value, int.Parse(arrSize1), int.Parse(arrSize2), dataType, editType))
            {
                Logger.LogMessage($"Can't setup {columnType.ToString().ToLower()} value with given parameters");
                return varDataGrid;
            }

            // Step 2: Add Value to the DataGrid
            VariableSelectionMoveTo(tabType, dataType, editType, VariableColumnType.EditType, columnType);
            VariableInfo variableInfo = new VariableInfo(tabType, callName, dataType, editType);
            AddValueToDataGrid(varDataGrid, variableInfo, columnType, value);

            return varDataGrid;
        }

        // Method for Initializing and Validating Variables
        private PP5DataGrid InitializeAndValidateVariable(VariableTabType tabType, string showName, string callName, VariableDataType dataType, VariableEditType editType,
                                                          string arrSize1, string arrSize2, Func<VariableDataType, bool> validateFunc, string errorMessage)
        {
            PP5DataGrid varDataGrid = InitializeVariableDataGrid(tabType, showName, callName, dataType, editType, arrSize1, arrSize2);
            if (!validateFunc(dataType))
            {
                throw new ArgumentException(errorMessage, dataType.ToString());
            }
            return varDataGrid;
        }

        // Overloaded Method for Initializing and Validating Variables with Custom Validation Logic
        private PP5DataGrid InitializeAndValidateVariable(VariableTabType tabType, string showName, string callName, VariableDataType dataType, VariableEditType editType,
                                                          string arrSize1, string arrSize2, Func<string, int, int, VariableDataType, VariableEditType, bool> validateFunc,
                                                          string value, string arrSize1Value, string arrSize2Value, string errorMessage)
        {
            PP5DataGrid varDataGrid = InitializeVariableDataGrid(tabType, showName, callName, dataType, editType, arrSize1, arrSize2);
            if (!validateFunc(value, int.Parse(arrSize1Value), int.Parse(arrSize2Value), dataType, editType))
            {
                Logger.LogMessage(errorMessage);
                return varDataGrid;
            }
            return varDataGrid;
        }

        // Add Value to DataGrid
        public void AddValueToDataGrid(PP5DataGrid varDataGrid, VariableInfo variableInfo, string value)
        {
            string valueTemp = value;
            List<string[]> values = ParseArrayValues(ref valueTemp, out _, out _);
            VariableRowInfo vri = varDataGrid.TiVariables.GetVariableRowInfoSet(variableInfo.TabType, variableInfo.CallName);
            long HexStringOrExternalSignal = vri.CanSetHexString ? vri.CanSetExternalSignal ? 
                                             (long)VariableDataType.HexStringCtgy : (long)VariableEditType.External_Signal : -1;
            VariableValueEditorAddValue(valueTemp, values, variableInfo.DataType, HexStringOrExternalSignal);
        }

        // Add Value to DataGrid
        public void AddValueToDataGrid(PP5DataGrid varDataGrid, VariableInfo variableInfo, VariableColumnType columnType, object value)
        {
            // Step 1.1 validation of column type
            if (columnType == VariableColumnType.Min || columnType == VariableColumnType.Max || 
                columnType == VariableColumnType.Default || columnType == VariableColumnType.Format)
            {
                if (value == null || string.IsNullOrWhiteSpace(value.ToString()))
                {
                    throw new ArgumentException($"Invalid value for {columnType}");
                }
            }

            // Step 1.2 validation of value by given variableInfo
            //validateVariable = new ValidateVariableValue(VariableCheckCanAddDefaultValue);
            //validateVariable.Invoke(variableInfo, value);
            if (!VariableCheckCanAddValue(variableInfo, columnType, value))
            {
                string errorMessage = columnType == VariableColumnType.Format ? "Only \"Float\" related datatype is supported!" 
                                                                              : $"Can't setup \"{columnType.ToString()}\" value with given parameters";
                throw new ArgumentException(errorMessage, variableInfo.DataType.ToString());
            }
                

            // Step 2. Move to the given
            VariableSelectionMoveToSpecified(varDataGrid, variableInfo.TabType, variableInfo.DataType, variableInfo.EditType, columnType);

            // Step 3. Add data to cell
            switch (columnType)
            {
                case VariableColumnType.Min:
                case VariableColumnType.Max:
                    AddValueToDataGrid(varDataGrid, variableInfo, value.ToString());
                    break;

                case VariableColumnType.Default:
                    //string valueTemp = value.ToString();
                    //List<string[]> values = ParseArrayValues(ref valueTemp, out _, out _);
                    //VariableValueEditorAddValue(valueTemp, values, dataType);
                    // Step 2: Modify Default Value if HexString
                    if (IsVariableHexStringDataType(variableInfo.DataType))
                        value = string.Concat(Enumerable.Repeat(value, int.Parse(varDataGrid.GetCellValue(1, "Max."))));
                    AddValueToDataGrid(varDataGrid, variableInfo, value.ToString());
                    break;

                case VariableColumnType.Format:
                    ApplyFormatToDataGrid(varDataGrid, (VariableDecimalPlaces[])value);
                    break;
                default:
                    throw new ArgumentException("Unsupported column type for adding value", nameof(columnType));
            }
        }

        // Apply Format to DataGrid
        private void ApplyFormatToDataGrid(PP5DataGrid varDataGrid, VariableDecimalPlaces[] formatArr)
        {
            varDataGrid.RefreshSelectedCell();
            foreach (var format in formatArr)
            {
                string fmExpected = format.GetDescription();
                varDataGrid.SelectedCellInfo.SendText(fmExpected);
                fmExpected.ShouldEqualTo(varDataGrid.GetCellValue(1, "Format"));
            }
        }

        private bool CheckValueTypeValid(List<string[]> arrayValues, VariableDataType DataType)
        {
            //var arrayValues = ParseArrayValues(value, out int rank1, out int rank2);

            if (arrayValues == null)
            {
                return false; // 格式不正確
            }

            // 驗證數值的格式與對應的數據型態
            foreach (var row in arrayValues)
            {
                foreach (var item in row)
                {
                    if (!IsValidForDataType(item, DataType))
                    {
                        return false;
                    }
                }
            }

            // 如果所有數值都符合指定型態，則返回 true
            return true;
        }

        private bool CheckValueTypeValid(string value, VariableDataType DataType)
        {
            //(int rank1, int rank2) = GetArrayRanks(arrayValues, rank1, rank2);
            string valueTemp = value;
            var arrayValues = ParseArrayValues(ref valueTemp, out int rank1, out int rank2);
            return CheckValueTypeValid(arrayValues, DataType) && rank1 == 0 && rank2 == 0;
        }

        private bool CheckValueTypeValid(string value, VariableDataType DataType, int arrSize1)
        {
            string valueTemp = value;
            //(int rank1, int rank2) = GetArrayRanks(arrayValues, rank1, rank2);
            var arrayValues = ParseArrayValues(ref valueTemp, out int rank1, out int rank2);
            return CheckValueTypeValid(arrayValues, DataType) && rank1 == arrSize1 && rank2 == 0;
        }

        private bool CheckValueTypeValid(string value, VariableDataType DataType, int arrSize1, int arrSize2)
        {
            //(int rank1, int rank2) = GetArrayRanks(value);
            string valueTemp = value;
            var arrayValues = ParseArrayValues(ref valueTemp, out int rank1, out int rank2);
            return CheckValueTypeValid(arrayValues, DataType) && rank1 == arrSize1 && rank2 == arrSize2;
        }

        private bool IsValidForDataType(string value, VariableDataType DataType)
        {
            // 驗證是否為空值 (如果空值是有效的，應考慮為 "null")
            if (string.IsNullOrWhiteSpace(value) && (DataType != VariableDataType.String && DataType != VariableDataType.StringArray &&
                                                     DataType != VariableDataType.String2DArray && DataType != VariableDataType.StringArrayOfUUTMaxSize))
            {
                return value.Trim().ToLower() == "null";
            }

            switch (DataType)
            {
                case VariableDataType.Float:
                case VariableDataType.FloatArray:
                case VariableDataType.FloatPercentage:
                case VariableDataType.FloatPercentageArrayOfSize640:
                case VariableDataType.FloatArrayOfSize640:
                case VariableDataType.Float2DArray:
                    return float.TryParse(value, out _);

                case VariableDataType.Integer:
                case VariableDataType.IntegerArray:
                case VariableDataType.IntegerArrayOfSize640:
                case VariableDataType.Integer2DArray:
                    return int.TryParse(value, out _);

                case VariableDataType.Double:
                case VariableDataType.DoubleArray:
                case VariableDataType.DoubleArrayOfUUTMaxSize:
                case VariableDataType.Double2DArray:
                    return double.TryParse(value, out _);

                case VariableDataType.Short:
                    return short.TryParse(value, out _);

                case VariableDataType.String:
                case VariableDataType.StringArray:
                case VariableDataType.String2DArray:
                case VariableDataType.StringArrayOfUUTMaxSize:
                case VariableDataType.HexString:
                case VariableDataType.HexStringArray:
                case VariableDataType.HexString2DArray:
                    // 字串類型可以接受任意值
                    return true;

                case VariableDataType.Byte:
                case VariableDataType.ByteArray:
                case VariableDataType.Byte2DArray:
                    return byte.TryParse(value, out _);

                case VariableDataType.Long:
                case VariableDataType.LongArray:
                case VariableDataType.Long2DArray:
                    return long.TryParse(value, out _);

                default:
                    // 若無法識別類型，視為無效
                    return false;
            }
        }

        private List<string[]> ParseArrayValues(ref string value, out int rank1, out int rank2)
        {
            List<string[]> arrayValues = new List<string[]>();
            rank1 = 0;
            rank2 = 0;

            // 檢查是否符合 rank 格式，如 "10[5]" 或 "10[2,3]"
            if (value.Contains('[') && value.Contains(']'))
            {
                StringBuilder sb = new StringBuilder();
                var rankParts = value.Split('[');
                if (rankParts.Length != 2 || !rankParts[1].EndsWith("]"))
                {
                    return null; // 格式不正確，應包含且僅包含一個 '[' 和對應的 ']'.
                }

                string valuePart = rankParts[0].Trim();
                string rankPart = rankParts[1].TrimEnd(']');
                int[] ranks = rankPart.Split(',').Select(int.Parse).ToArray();

                if (ranks.Length == 1)
                {
                    // 1D 陣列，生成標準格式
                    rank1 = ranks[0];
                    arrayValues.Add(Enumerable.Repeat(valuePart, ranks[0]).ToArray());
                    foreach (var val in arrayValues[0])
                        sb.AppendFormat("{0},", val);
                    sb.Remove(sb.Length - 1, 1);
                }
                else if (ranks.Length == 2)
                {
                    // 2D 陣列，生成標準格式
                    rank1 = ranks[0];
                    rank2 = ranks[1];
                    for (int i = 0; i < ranks[0]; i++)
                    {
                        arrayValues.Add(Enumerable.Repeat(valuePart, ranks[1]).ToArray());
                        foreach (var val in arrayValues[i])
                            sb.AppendFormat("{0},", val);
                        sb.Remove(sb.Length - 1, 1);
                        sb.Append(";");
                    }
                    sb.Remove(sb.Length - 1, 1);
                }
                else
                {
                    return null; // 不支持的 rank 格式
                }

                value = sb.ToString();
            }
            else
            {
                // 將輸入的值分割成陣列
                string[] rows = value.Split(';'); // 分割行數
                foreach (string row in rows)
                {
                    arrayValues.Add(row.Split(',')); // 分割列數
                }

                rank1 = rows.Length;
                rank2 = arrayValues[0].Length;

                foreach (var row in arrayValues)
                {
                    if (row.Length != rank2)
                    {
                        return null; // 所有行應有相同的列長度
                    }
                }

                if (!value.Contains(',') && !value.Contains(';'))
                {
                    //arrayValues = null;
                    rank1 = rank2 = 0;
                }
            }

            return arrayValues;
        }


        private void VariableValueEditorAddValue(string value, List<string[]> valueArr, VariableDataType DataType, long HexStringOrExternalSignal)
        {
            bool manualSetup = false;
            string valueTmp = value;
            manualSetup = !valueArr.ToArray().AreAllElementsEqual();
            if (!manualSetup)
                valueTmp = valueArr[0][0];
            StaticData.isExternalSignal = false;
            StaticData.isHexString = false;

            if (IsVariableHexStringDataType(DataType)) StaticData.isHexString = true;
            if (IsVariableStringDataType(DataType)) StaticData.isExternalSignal = true;

            //string defaultRes = null;
            if (DataType == VariableDataType.HexString)
            {
                // For single HexString setup
                if (HexStringOrExternalSignal == (long)VariableDataType.HexStringCtgy)
                    SetupHexString(new string[] { value });
            }
            else if (DataType == VariableDataType.String)
            {
                // For String extern signal setup
                if (HexStringOrExternalSignal == (long)VariableEditType.External_Signal)
                    SetupExternalSignal(value);
                else
                    SendSingleKeys(value);
            }
            else if (IsVariable1DArrayDataType(DataType) || IsVariableUUTArrayDataType(DataType))
            {
                // For 1D array setup
                Setup1DArray(valueTmp, manualSetup, HexStringOrExternalSignal);
            }
            else if (IsVariable2DArrayDataType(DataType))
            {
                // For 2D array setup
                Setup2DArray(valueTmp, manualSetup, HexStringOrExternalSignal);
            }
            else if (IsVariableVectorDataType(DataType))
            {
                // For Vector setup
                SetupVector();
            }
            else
            {
                SendSingleKeys(value);
            }
        }

        private void Setup2DArray(string value, bool manualSetup, long HexStringOrExternalSignal)
        {
            var variableEditorWindow = PP5IDEWindow.PerformGetElement("/ByName[Variables Value Editor]");
            if (manualSetup)
            {
                variableEditorWindow.PerformClick("/ByClass[GroupBox]/ById[TC_ParameterApply]/ByName[Outer]", ClickType.LeftClick);
                variableEditorWindow.PerformGetElement("/ByName[Outer]/ById[svOuter,cb_1Dchar]").ComboBoxSelectByName(Rank1SplitChar.Comma.GetDescription());
                variableEditorWindow.PerformGetElement("/ByName[Outer]/ById[svOuter,cb_2Dchar]").ComboBoxSelectByName(Rank2SplitChar.Semicolon.GetDescription());
                List<string[]> arr = new List<string[]>();
                string valTemp = value;
                if (HexStringOrExternalSignal == (int)VariableDataType.HexString)
                    SetupHexString(ParseArrayValues(ref valTemp, out _, out _).ToArray());
                else if (HexStringOrExternalSignal == (int)VariableEditType.External_Signal)
                    SetupExternalSignal(ParseArrayValues(ref valTemp, out _, out _).ToArray());
                else
                    variableEditorWindow.PerformInput("/ById[svOuter]/ByClass[PopupBaseEdit]/ById[PART_Editor]", InputType.SendContent, value);
                variableEditorWindow.PerformClick("/ByName[Outer]/ById[svOuter]/Button[1]", ClickType.LeftClick);
            }
            else
            {
                var edtBxInner = variableEditorWindow.PerformGetElement("/ById[svInner,popEdit,PART_Editor]");
                if (HexStringOrExternalSignal == (int)VariableDataType.HexString)
                {
                    edtBxInner.PerformClick("/", ClickType.LeftClick);
                    SetupHexString(new string[] { value });
                }
                else if (HexStringOrExternalSignal == (int)VariableEditType.External_Signal)
                {
                    edtBxInner.PerformClick("/", ClickType.LeftClick);
                    SetupExternalSignal(value);
                }
                else
                    edtBxInner.PerformInput("/", InputType.SendContent, value);
                variableEditorWindow.PerformClick("/ByName[Inner]/ById[svInner]/ByName[Apply]", ClickType.LeftClick);
                variableEditorWindow.PerformClick("/ByName[Ok]", ClickType.LeftClick);
            }
        }

        private void Setup1DArray(string value, bool manualSetup, long HexStringOrExternalSignal)
        {
            var variableEditorWindow = PP5IDEWindow.PerformGetElement("/ByName[Variables Value Editor]");
            if (manualSetup)
            {
                variableEditorWindow.PerformClick("/ByClass[GroupBox]/ById[TC_ParameterApply]/ByName[Outer]", ClickType.LeftClick);
                variableEditorWindow.PerformGetElement("/ByName[Outer]/ById[svOuter,cb_char]").ComboBoxSelectByName(Rank1SplitChar.Comma.GetDescription());
                if (HexStringOrExternalSignal == (int)VariableDataType.HexString)
                    SetupHexString(value.Split(','));
                else if (HexStringOrExternalSignal == (int)VariableEditType.External_Signal)
                    SetupExternalSignal(value.Split(','));
                else
                    variableEditorWindow.PerformInput("/ById[svOuter]/ByClass[PopupBaseEdit]/ById[PART_Editor]", InputType.SendContent, value);
                variableEditorWindow.PerformClick("/ByName[Outer]/ById[svOuter]/Button[1]", ClickType.LeftClick);
            }
            else
            {
                var edtBxInner = variableEditorWindow.PerformGetElement("/ById[svInner,popEdit,PART_Editor]");
                if (HexStringOrExternalSignal == (int)VariableDataType.HexString)
                {
                    edtBxInner.PerformClick("/", ClickType.LeftClick);
                    SetupHexString(new string[] { value });
                }
                else if (HexStringOrExternalSignal == (int)VariableEditType.External_Signal)
                {
                    edtBxInner.PerformClick("/", ClickType.LeftClick);
                    SetupExternalSignal(value);
                }
                else
                    edtBxInner.PerformInput("/", InputType.SendContent, value);
                variableEditorWindow.PerformClick("/ByName[Inner]/ById[svInner]/ByName[Apply]", ClickType.LeftClick);
                variableEditorWindow.PerformClick("/ByName[Ok]", ClickType.LeftClick);
            }
            variableEditorWindow.PerformClick("/ByName[Ok]", ClickType.LeftClick);
        }

        private void SetupVector()
        {
            var vectorEditorWindow = PP5IDEWindow.PerformGetElement("/ByName[Vector Editor]");

            //vectorEditorWindow.PerformClick($"/ByCell[2,#{callName}]", ClickType.LeftClick);
            if (StaticData.currentDataGrid == null)
                StaticData.currentDataGrid = new PP5DataGrid((PP5Element)vectorEditorWindow.PerformGetElement("/ById[VectorGrid]"));

            StaticData.currentDataGrid.ScrollToSpecificColumn("Default");
            StaticData.currentDataGrid.cellValueCache = String.Join("@", StaticData.currentDataGrid.GetSingleColumnValues(7).ToArray());

            vectorEditorWindow.PerformClick("/ByName[Ok]", ClickType.LeftClick);
        }

        private void SetupHexString(IElement HexStringEditorWindow, string hexValuePattern)
        {
            if (!hexValuePattern.IsHexString() || hexValuePattern.Length > 2048)
                throw new ArgumentException($"hexValuePattern \"{hexValuePattern}\" not valid or length is over 2048");

            var hexValues = hexValuePattern.SplitStringIntoIntervals(2);
            //IElement HexStringEditorWindow;
            //if (StaticData.currentElement == null)
            //    HexStringEditorWindow = PP5IDEWindow.PerformGetElement("/ByName[HexString Editor]");
            //else
            //    HexStringEditorWindow = StaticData.currentElement;

            for (var i = 0; i < hexValues.Count; i++)
            {
                HexStringEditorWindow.PerformInput($"/ByClass[ListBox]/ListItem[{i}]/Edit[0]", InputType.SendContent, hexValues[i]);
            }
            //HexStringEditorWindow.PerformInput("/ByClass[ListBox]/ListItem[1]/Edit[0]", InputType.SendContent, new string(value, 2));
            HexStringEditorWindow.PerformClick("/ByName[Ok]", ClickType.LeftClick);

        }

        private void SetupHexString(string[] hexValuesPattern)
        {
            PP5Element HexStringEditorWindow = (PP5Element)PP5IDEWindow.PerformGetElement("/ByName[HexString Editor]");
            //StaticData.currentElement = HexStringEditorWindow;
            foreach (var hexValuePattern in hexValuesPattern)
            {
                SetupHexString(HexStringEditorWindow, hexValuePattern);
            }
        }

        private void SetupHexString(string[][] hexValuesPattern)
        {
            PP5Element HexStringEditorWindow = (PP5Element)PP5IDEWindow.PerformGetElement("/ByName[HexString Editor]");
            //StaticData.currentElement = HexStringEditorWindow;
            foreach (var row in hexValuesPattern)
            {
                foreach (var hexValuePattern in row)
                    SetupHexString(HexStringEditorWindow, hexValuePattern);
            }
        }

        private void SetupExternalSignal(string externalSignalPattern)
        {
            //varDataGrid.GetCellBy(varDataGrid.LastRowNo, "Default").LeftClick();
            // externalSignalPattern: "PhoenixonDSP", "DSPSensingOFF(0xFF)", "DSPSensingOFF_Signal"
            PP5IDEWindow.PerformGetElement("/ById[tvCluster]")
                        .AddTreeViewItem(externalSignalPattern.Split('.'));

            //string defaultValue = varDataGrid.GetCellBy(varDataGrid.LastRowNo, "Default").GetText();
            //defaultValue.ShouldContain("PhoenixonDSP");
            //defaultValue.ShouldContain("DSPSensingOFF");
            //defaultValue.ShouldContain("(0xFF)");
            //defaultValue.ShouldContain("DSPSensingOFF_Signal");
        }

        private void SetupExternalSignal(string[] externalSignalsPattern)
        {
            foreach (var externalSignalPattern in externalSignalsPattern)
                SetupExternalSignal(externalSignalPattern);
        }

        private void SetupExternalSignal(string[][] externalSignalsPattern)
        {
            foreach (var row in externalSignalsPattern)
                foreach (var externalSignalPattern in row)
                    SetupExternalSignal(externalSignalPattern);
        }

        public void VariableSelectionMoveToSpecified(PP5DataGrid varDataGrid, VariableTabType tabType, VariableDataType DataType, VariableEditType EditType, VariableColumnType toColumn)
        {
            varDataGrid.RefreshSelectedCell();
            var currSelectedCol = TypeExtension.GetEnumByDescription<VariableColumnType>(varDataGrid.SelectedCellInfo.ColumnName);
            int stepCount = toColumn - currSelectedCol;
            VariableSelectionMoveTo(tabType, DataType, EditType, currSelectedCol, currSelectedCol + stepCount);
        }

        public void VariableSelectionMoveForward(PP5DataGrid varDataGrid, VariableTabType tabType, VariableDataType DataType, VariableEditType EditType, int stepCount)
        {
            var currSelectedCol = TypeExtension.GetEnumByDescription<VariableColumnType>(varDataGrid.SelectedCellInfo.ColumnName);
            VariableSelectionMoveTo(tabType, DataType, EditType, currSelectedCol, currSelectedCol + stepCount);
        }

        public void VariableSelectionMoveBackward(PP5DataGrid varDataGrid, VariableTabType tabType, VariableDataType DataType, VariableEditType EditType, int stepCount)
        {
            var currSelectedCol = TypeExtension.GetEnumByDescription<VariableColumnType>(varDataGrid.SelectedCellInfo.ColumnName);
            VariableSelectionMoveTo(tabType, DataType, EditType, currSelectedCol, currSelectedCol - stepCount);
        }

        public void VariableSelectionMovePrev(PP5DataGrid varDataGrid, VariableTabType tabType, VariableDataType DataType, VariableEditType EditType)
        {
            var currSelectedCol = TypeExtension.GetEnumByDescription<VariableColumnType>(varDataGrid.SelectedCellInfo.ColumnName);
            VariableSelectionMoveTo(tabType, DataType, EditType, currSelectedCol, currSelectedCol - 1);
        }

        public void VariableSelectionMoveTo(VariableTabType tabType, VariableDataType DataType, VariableEditType EditType, VariableColumnType fromColumn, VariableColumnType toColumn)
        {
            List<VariableColumnType> columns = new List<VariableColumnType>();
            switch (tabType)
            {
                case VariableTabType.Condition:
                    columns = new List<VariableColumnType> { VariableColumnType.No, VariableColumnType.ShowName, VariableColumnType.CallName, VariableColumnType.DataType, VariableColumnType.EditType,
                                                             VariableColumnType.Min, VariableColumnType.Max, VariableColumnType.Default,
                                                             VariableColumnType.Format, VariableColumnType.EnumItem };
                    break;
                case VariableTabType.Result:
                    columns = new List<VariableColumnType> { VariableColumnType.No, VariableColumnType.ShowName, VariableColumnType.CallName, VariableColumnType.DataType,
                                                             VariableColumnType.MinimumSpec, VariableColumnType.DefectCodeMin, VariableColumnType.MaximumSpec, VariableColumnType.DefectCodeMax };
                    break;
                case VariableTabType.Temp:
                    columns = new List<VariableColumnType> { VariableColumnType.No, VariableColumnType.ShowName, VariableColumnType.CallName, VariableColumnType.DataType,
                                                             VariableColumnType.Max, VariableColumnType.Default };
                    break;
                case VariableTabType.Global:
                    columns = new List<VariableColumnType> { VariableColumnType.Lock, VariableColumnType.No, VariableColumnType.ShowName, VariableColumnType.CallName,
                                                             VariableColumnType.DataType, VariableColumnType.EditType,
                                                             VariableColumnType.Min, VariableColumnType.Max, VariableColumnType.Default,
                                                             VariableColumnType.Format, VariableColumnType.EnumItem };
                    break;
                default:
                    break;
            }

            if (!columns.Contains(fromColumn))
                throw new ArgumentException($"Given \"from\" column not existed in specified VariableTabType \"{tabType}\"", fromColumn.ToString());
            if (!columns.Contains(toColumn))
                throw new ArgumentException($"Given \"to\" column not existed in specified VariableTabType \"{tabType}\"", toColumn.ToString());
            if (fromColumn == toColumn)
                throw new ArgumentException($"Given \"from\" is same as the \"to\" column", toColumn.ToString());

            int fromColumnIdx = columns.IndexOf(fromColumn);
            int toColumnIdx = columns.IndexOf(toColumn);

            bool isLast = false;
            if (fromColumnIdx < toColumnIdx)  // left to right
            {
                for (int i = fromColumnIdx;  i < toColumnIdx; i++)
                {
                    if (i == toColumnIdx - 1) { isLast = true; }
                    Press(Keys.Tab);
                    VariableOnSelected(DataType, EditType, columns[i + 1], isLast);
                }
            }
            else                                // right to left
            {
                for (int i = fromColumnIdx; i > toColumnIdx; i--)
                {
                    if (i == toColumnIdx + 1) { isLast = true; }
                    Press(Keys.Shift, Keys.Tab);
                    VariableOnSelected(DataType, EditType, columns[i - 1], isLast);
                }
            }
        }

        public void VariableOnSelected(VariableDataType DataType, VariableEditType EditType, VariableColumnType col, bool isLast)
        {
            if (isLast) return;
            switch (col)
            {
                case VariableColumnType.Lock:
                case VariableColumnType.No:
                case VariableColumnType.ShowName:
                case VariableColumnType.CallName:
                case VariableColumnType.DataType:
                case VariableColumnType.EditType:
                case VariableColumnType.Format:
                case VariableColumnType.MinimumSpec:
                case VariableColumnType.DefectCodeMin:
                case VariableColumnType.MaximumSpec:
                case VariableColumnType.DefectCodeMax:
                    break;
                case VariableColumnType.Min:
                case VariableColumnType.Max:
                    if (EditType == VariableEditType.EditBox && (IsVariableArrayDataType(DataType) && !IsVariableHexStringDataType(DataType)
                                                                                                    && !IsVariableStringDataType(DataType)
                                                                                                    && !IsVariableByteArrayDataType(DataType)))
                            PP5IDEWindow.PerformClick("/ByName[Variables Value Editor,Cancel]", ClickType.LeftClick);
                    break;
                case VariableColumnType.Default:
                    if (EditType == VariableEditType.EditBox)
                    {
                        if (IsVariableArrayDataType(DataType) || DataType == VariableDataType.HexString || IsVariableVectorDataType(DataType))
                            PP5IDEWindow.CloseWindow(0);
                    }
                    else if (EditType == VariableEditType.ComboBox)
                    {
                        if (IsVariableVectorDataType(DataType) || IsVariableArrayDataType(DataType))
                            PP5IDEWindow.CloseWindow(0);
                    }
                    break;
                case VariableColumnType.EnumItem:
                    if (EditType == VariableEditType.ComboBox && !IsVariableVectorDataType(DataType))
                        PP5IDEWindow.CloseWindow(0);
                    break;
                default:
                    break;
            }
        }

        public void RecordVariableInfo(PP5DataGrid varDataGrid, VariableInfo variableInfo)
        {
            VariableRowInfo variableRowInfo = new VariableRowInfo(new DataTypeInfo(variableInfo.DataType, variableInfo.ArrSize1, variableInfo.ArrSize2), variableInfo.EditType);
            VariableRowInfoSet variableRowInfoSet = new VariableRowInfoSet(variableInfo.TabType);
            variableRowInfoSet.variableRowInfos.Add(variableInfo.CallName, variableRowInfo);
            varDataGrid.TiVariables.variableRowInfoSets.Add(variableRowInfoSet);
            //switch (tabType)
            //{
            //    case VariableTabType.Condition:
            //        StaticData.conditionVariables = variableInfoList;
            //        break;
            //    case VariableTabType.Result:
            //        StaticData.resultVariables = variableInfoList;
            //        break;
            //    case VariableTabType.Temp:
            //        StaticData.tempVariables = variableInfoList;
            //        break;
            //    case VariableTabType.Global:
            //        StaticData.globalVariables = variableInfoList;
            //        break;
            //}
        }

        public void RemoveVariableInfo(PP5DataGrid varDataGrid, VariableTabType tabType, string callName)
        {
            //StaticData.TiVariables.Remove(tabType, callName);
            varDataGrid.TiVariables.Remove(tabType, callName);
        }

        public IWebElement GetCommandIsSelected(string cmdName, string GroupNameToSearch = "", bool findExactSameCommand = true)
        {
            ////var CommandsMapMemoryCache = new MemoryCache<string, IWebElement>(CommandsMapCache);

            //// If row index within current CommandsMap cache row count, query element from cache
            //if (!CommandsMapCache.IsEmpty() && CommandsMapCache.Keys.Contains(cmdName))
            //    return CommandsMapCache.Get(cmdName);

            //else  // If row index out of current CommandsMap cache, resave the table
            //{
            //    SaveCommandMap(cmdName, GroupNameToSearch, findExactSameCommand);
            //    return CommandsMapCache.Get(cmdName);
            //}

            var cmdTreeItems = GetExpandedCommandGroup(GetCommandTreeByGroupName(QueryGroupName(cmdName)));
            return GetCommandIsSelected(cmdTreeItems);
        }

        public void SaveCommandMap(string cmdNameToSearch, string GroupNameToSearch = "", bool findExactSameCommand = true)
        {
            //bool RestartCommandSearch;
            int CommandSearchCount = 1;
            IElement cmdFound;

            //do
            //{
            //RestartCommandSearch = false;

            //Dictionary<string, AppiumWebElement> CommandsMap = new Dictionary<string, AppiumWebElement>();

            //WindowsElement testCmdSearchBox = CurrentDriver.GetElementFromWebElement(5000, By.ClassName("CmdTreeView"), MobileBy.AccessibilityId("searchBox"));
            //testCmdSearchBox.ClearContent();
            //testCmdSearchBox.SendComboKeys(CommandName, OpenQA.Selenium.Keys.Enter);

            IElement CommandTreeView = CurrentDriver.GetExtendedElement(PP5By.ClassName("CmdTreeView"));
            IElement testCmdSearchBox = CommandTreeView.GetExtendedElement(PP5By.Id("searchBox"));
            IReadOnlyCollection<IElement> subCmdTreeViews = CommandTreeView.GetExtendedElements(PP5By.ClassName("TreeView"));

            Assert.AreEqual(2, subCmdTreeViews.Count);
            testCmdSearchBox.ClearContent();
            testCmdSearchBox.SendComboKeys(cmdNameToSearch, OpenQA.Selenium.Keys.Enter);

            foreach (var subCmdTreeView in subCmdTreeViews)
            {
                //string CmdTreeViewID = subCmdTreeView.Id;

                var CmdGroupsTreeViewItems = subCmdTreeView.GetChildElements();  // Collapsed (0)

                foreach (var cmdGroup in CmdGroupsTreeViewItems)
                {
                    if (cmdGroup.isElementCollapsed())
                        continue;

                    string GroupName = cmdGroup.GetFirstTextContent();

                    if (CommandSearchCount > 3 && !GroupNameToSearch.IsNullOrEmpty() && GroupName == GroupNameToSearch)
                    {
                        // Maximum search limit (3 times) reached, open group tree to add command directly
                        cmdGroup.LeftClick();
                        cmdFound = cmdGroup.GetExtendedElements(PP5By.XPath($".//TreeItem[@ClassName='TreeViewItem']")).ToList()
                                           .Find(cmd => cmd.GetFirstTextContent() == cmdNameToSearch);

                        if (cmdFound != null)
                            CommandsMapCache.Add(cmdNameToSearch, cmdFound);
                        return;
                    }

                    var cmdGroupTemp = cmdGroup.GetExtendedElements(PP5By.XPath($".//TreeItem[@ClassName='TreeViewItem']")).ToList();
                    cmdGroupTemp.RemoveAt(0);

                    cmdFound = cmdGroupTemp.Find(c => c.Selected == true);

                    string cmdName = cmdFound.GetFirstTextContent();

                    //ReadOnlyCollection<AppiumWebElement> cmdsTreeViewItems = new ReadOnlyCollection<AppiumWebElement>(cmdGroupTemp);

                    if (cmdName != cmdNameToSearch)
                    {
                        // Command searched not matched with given command name, press F3 to continue to the next matched command
                        // Then restart the search action
                        if (!findExactSameCommand)
                        {
                            CommandsMapCache.Add(cmdNameToSearch, cmdFound);
                            return;
                        }
                        SendSingleKeys(Keys.F3);
                        continue;
                        //RestartCommandSearch = true;
                        //break;
                    }
                    else
                    {
                        CommandsMapCache.Add(cmdName, cmdFound);
                        return;
                    }

                    /* Legacy
                    //foreach (var cmd in cmdsTreeViewItems)
                    //{
                        //if (cmd.GetAttribute("SelectionItem.IsSelected") == "false")
                        //    continue;

                        //string cmdName = ((WindowsElement)cmd).GetSubElementText();

                        //if (cmdName != cmdNameToSearch)
                        //{
                        //    // Command searched not matched with given command name, press F3 to continue to the next matched command
                        //    // Then restart the search action
                        //    SendKeys(Keys.F3);
                        //    RestartCommandSearch = true;
                        //    break;
                        //}
                        //else
                        //{
                        //    CommandsMapCache.Add(cmdName, cmd);
                        //    return;
                        //}
                    //}
                    //if (RestartCommandSearch)
                    //    continue;
                    */
                }
            }
            CommandSearchCount++;
            //}
            //while (RestartCommandSearch);

            //return CommandsMapCache;
        }

        //public void RefreshDataTable(DataTableAutoIDType DataGridType)
        //{
        //    SaveGridTable(GetDataTableElement(DataGridType), DataGridType);
        //}

        public IElement GetCellBy(string DataGridAutomationID, int rowNo, string colName)
        {
            return GetDataTableElement(DataGridAutomationID).GetCellBy(rowNo, colName);

            //// If row index within current datatable, query element from cache
            //if (!dataTableTemp.IsEmpty() && dataTableTemp.Keys.Contains((rowIdx, colName)))
            //    return dataTableTemp.Get((rowIdx, colName));

            //else  // If row index out of current datatable cache, resave the table
            //{
            //    SaveGridCell(DataGridAutomationID, rowIdx, colName);
            //    return dataTableTemp.Get((rowIdx, colName));
            //}
        }

        public IEnumerable<IElement> GetColumnBy(DataTableAutoIDType DataGridType, int colNo, bool excludeLastRow = true)
        {
            return GetDataTableElement(DataGridType).GetColumnBy(colNo);

            //return excludeLastRow ? dataTableLookUp[colName].ToList().GetRange(0, colElements.Count - 1) // Exclude the last column (empty)
            //                      : dataTableLookUp[colName].ToList();  

            //// If Column name in current datatable, query single column from cache
            //ILookup<string, IWebElement> dataTableLookUp;
            //if (!dataTableTemp.IsEmpty() && dataTableTemp.Keys.Select(k => k.Item2).Contains(colName))
            //{
            //    // Skip action
            //}
            //else  // If Column name out of current datatable cache, resave the table and return single column from cache
            //{
            //    SaveGridTable(GetDataTableElement(DataGridType), DataGridType);
            //}

            //dataTableLookUp = dataTableTemp.CacheData.ToLookup(x => x.Key.Item2, x => x.Value);
            //List<IWebElement> colElements = dataTableLookUp[colName].ToList();
            //return excludeLastRow ? dataTableLookUp[colName].ToList().GetRange(0, colElements.Count - 1) // Exclude the last column (empty)
            //                      : dataTableLookUp[colName].ToList();    
        }

        public IEnumerable<IWebElement> GetRowCellElementsBy(DataTableAutoIDType DataGridType, int rowNo)
        {
            return GetDataTableElement(DataGridType).GetDataTableRowElements()[rowNo - 1].GetCellElementsOfRow();

            //// If row index within current datatable, query single row from cache
            //ILookup<int?, IWebElement> dataTableLookUp;
            //if (!dataTableTemp.IsEmpty() && dataTableTemp.Keys.Select(k => k.Item1).Contains(rowIdx))
            //{
            //    // Skip action
            //}
            //else  // If row index out of current datatable cache, resave the table and return single row from cache
            //{
            //    SaveGridTable(GetDataTableElement(DataGridType), DataGridType);
            //}

            //dataTableLookUp = dataTableTemp.CacheData.ToLookup(x => x.Key.Item1, x => x.Value);
            //return dataTableLookUp[rowIdx].ToList();
        }

        public string GetCellValue(string DataGridAutomationID, int rowIdx, string colName)
        {
            IElement cell = GetCellBy(DataGridAutomationID, rowIdx, colName);
            return cell.GetCellValue();
        }

        public int GetRowCount(DataTableAutoIDType DataGridType)
        {
            string DataGridAutomationID = DataGridType.ToString();
            return GetRowCount(DataGridAutomationID);
        }

        public int GetRowCount(string DataGridAutomationID)
        {
            return GetDataTableElement(DataGridAutomationID).GetDataTableRowElements().Count;
        }

        public List<string> GetSingleColumnValues(DataTableAutoIDType DataGridType, int colNo, bool excludeLastRow = true)
        {
            IEnumerable<IWebElement> column = GetColumnBy(DataGridType, colNo, excludeLastRow);
            List<string> columnValues = new List<string>();

            if (column == null)
                return null;
            else
            {
                columnValues = column.Select(c => c.GetAttribute("Value.Value")).ToList();
                columnValues.RemoveAll(s => s.IsNullOrEmpty());
                return columnValues;
            }
        }

        public List<string> GetSingleRowValues(DataTableAutoIDType DataGridType, int rowNo)
        {
            IEnumerable<IWebElement> row = GetRowCellElementsBy(DataGridType, rowNo);
            List<string> rowValues = new List<string>();

            if (row == null)
                return null;
            else
            {
                rowValues = row.Select(c => c.GetAttribute("Value.Value")).ToList();
                rowValues.RemoveAll(s => s.IsNullOrEmpty());
                return rowValues;
            }
        }

        #region Test Command/Test Item operations

        public string GetGroupNameByEnum(TestCmdGroupType cmdGrpType)
        {
            return cmdGrpType.GetDescription();
        }

        // 2024/07/12, Adam, add checking command source type
        public void AddCommandBy(TestCmdGroupType cmdGrpType, int cmdNumber = 1, int addCount = 1)
        {
            /* Legacy method of getting cmdtree by group name
            //IWebElement commandTree;
            //if (!GetCommandSourceType(groupName, out bool IsDevice))
            //    return;
            
            //if (IsDevice)
            //{
            //    commandTree = PP5IDEWindow.GetElementFromWebElement(MobileBy.AccessibilityId("DeivceCmdTree"));
            //}
            //else
            //{
            //    commandTree = PP5IDEWindow.GetElementFromWebElement(MobileBy.AccessibilityId("SystemCmdTree"));
            //}
            */

            IElement groupTreeItem = GetCmdGroupTreeItemByGroupName(cmdGrpType);      // Expand the group item tree
            AddCommandBy(groupTreeItem, cmdNumber, addCount, out IElement cmdToAdd);     // Add the command by given parameters
            TreeViewCollapseAll(cmdToAdd);                                                  // Press left arrow key twice to Close the group tree view

            /* Legacy method
            //// Get the element that matching the given groupname directly by XPath (Faster)
            //groupTreeItem = commandTree.GetElementFromWebElement(By.XPath($".//TreeItem[@ClassName='TreeViewItem']/Text[@Name='{groupName}']/parent::node()"), 3000);

            //if (groupTreeItem == null && !cmdListIsFocused)
            //{
            //    commandTree.LeftClick();
            //    cmdListIsFocused = true; // Set the flag to true after the left click on the command list
            //}

            //// If element if out of screen, press page down to find the element
            //if (groupTreeItem == null)
            //{
            //    Press(Keys.PageDown);
            //    Thread.Sleep(1);

            //    // If scroll to end of the command list, group item still not found, throw exception
            //    foreach (var cmdList in commandTree.GetElements(By.ClassName("TreeView")))
            //    {
            //        if (cmdList.GetAttribute("Scroll.VerticallyScrollable") == bool.FalseString)
            //            continue;

            //        if (cmdList.GetAttribute("Scroll.VerticalScrollPercent") == "100")
            //            throw new GroupNameNotExistedException(groupName);
            //    }
            //}
            */
        }

        // 2024/07/12, Adam, add checking command source type
        public void AddCommandBy(TestCmdGroupType cmdGrpType, string cmdName, int addCount = 1)
        {
            /* Legacy method of getting cmdtree by group name
            //IWebElement commandTree;
            //if (!GetCommandSourceType(groupName, out bool IsDevice))
            //    return;
            
            //if (IsDevice)
            //{
            //    commandTree = PP5IDEWindow.GetElementFromWebElement(MobileBy.AccessibilityId("DeivceCmdTree"));
            //}
            //else
            //{
            //    commandTree = PP5IDEWindow.GetElementFromWebElement(MobileBy.AccessibilityId("SystemCmdTree"));
            //}
            */

            IElement groupTreeItem = GetCmdGroupTreeItemByGroupName(cmdGrpType);
            AddCommandBy(groupTreeItem, cmdName, addCount, out IElement cmdToAdd);       // Add the command by given parameters
            TreeViewCollapseAll(cmdToAdd);                                                  // Press left arrow key twice to Close the group tree view
        }

        //// Current used method for adding command
        //public void AddCommandBy(string groupName, int cmdNumber = 1, TestCommandSourceType tcType = TestCommandSourceType.System, int addCount = 1)
        //{
        //    //IWebElement commandTree = CurrentDriver.GetElementFromWebElement(By.ClassName("CmdTreeView"));
        //    IWebElement commandTree;
        //    if (tcType == TestCommandSourceType.Device)
        //    {
        //        commandTree = CurrentDriver.GetElementFromWebElement(MobileBy.AccessibilityId("DeivceCmdTree"));
        //    }
        //    else
        //    {
        //        commandTree = CurrentDriver.GetElementFromWebElement(MobileBy.AccessibilityId("SystemCmdTree"));
        //    }

        //    //Console.WriteLine($"LeftClick on Text \"{groupName}\"");

        //    bool cmdListIsFocused = false;
        //    IWebElement groupTreeItem = null;
        //    while (groupTreeItem == null)
        //    {
        //        // Get the element that matching the given groupname directly by XPath (Faster)
        //        groupTreeItem = commandTree.GetElementFromWebElement(By.XPath($".//TreeItem[@ClassName='TreeViewItem']/Text[@Name='{groupName}']/parent::node()"), 3000);

        //        if (groupTreeItem == null && !cmdListIsFocused)
        //        {
        //            commandTree.LeftClick();
        //            cmdListIsFocused = true; // Set the flag to true after the left click on the command list
        //        }

        //        // If element if out of screen, press page down to find the element
        //        if (groupTreeItem == null)
        //        {
        //            Press(Keys.PageDown);
        //            Thread.Sleep(1);

        //            // If scroll to end of the command list, group item still not found, throw exception
        //            foreach (var cmdList in commandTree.GetElements(By.ClassName("TreeView")))
        //            {
        //                if (cmdList.GetAttribute("Scroll.VerticallyScrollable") == bool.FalseString)
        //                    continue;

        //                if (cmdList.GetAttribute("Scroll.VerticalScrollPercent") == "100")
        //                    throw new GroupNameNotExistedException(groupName);
        //            }
        //        }
        //    }

        //    // 2024/07/09, Adam, Expand the command group
        //    groupTreeItem.ExpandTreeView();

        //    //// Get all elements, then find element that matching the given groupname (Longer time required)
        //    //var groupTreeItem = commandTree.GetElements(By.XPath($".//TreeItem[@ClassName='TreeViewItem']")).ToList()
        //    //                   .Find(e => e.GetSubElementText() == groupName);

        //    //// Use attribute: "ExpandCollapse.ExpandCollapseState" to check the expand/collapse state, where: Expanded (1), Collapsed (0)
        //    //if (groupTreeItem.isElementCollapsed())
        //    //    groupTreeItem.DoubleClick();

        //    var cmdTreeItems = groupTreeItem.GetElements(By.XPath($".//TreeItem[@ClassName='TreeViewItem']"));

        //    if (cmdTreeItems.Count == 0)
        //        return;
        //    if (cmdNumber > cmdTreeItems.Count || cmdNumber <= -1)
        //        throw new CommandNumberOutOfRangeException(cmdNumber.ToString());

        //    IWebElement cmdToAdd = cmdNumber == -1 ? cmdTreeItems.Last() : cmdTreeItems[cmdNumber];

        //    //var cmdTreeItem = groupTreeItem.GetElementFromWebElement(By.XPath($"(.//TreeItem[@ClassName='TreeViewItem'])[{cmdNumber + 1}]"));
        //    Console.WriteLine($"LeftClick on Text \"{cmdToAdd.GetFirstTextContent()}\"");

        //    //// If element if out of screen, move to the element first
        //    while (bool.Parse(cmdToAdd.GetAttribute("IsOffscreen")))
        //    {
        //        Press(Keys.PageDown);
        //        Thread.Sleep(50);
        //    }

        //    // Add the command
        //    for (int i = 0; i < addCount; i++)
        //        cmdToAdd.DoubleClick();

        //    // Press left arrow key twice to Close the group tree view
        //    Press(Keys.Left);
        //    Press(Keys.Left);
        //}

        // Legacy method adding repeated commands with command name and repeat count
        //public void AddCommandBy(string groupName, string cmdName = "Arithmetic", TestCommandSourceType tcType = TestCommandSourceType.System, int addCount = 1)
        //{
        //    //IWebElement commandTree = CurrentDriver.GetElementFromWebElement(By.ClassName("CmdTreeView"));
        //    IWebElement commandTree;
        //    if (tcType == TestCommandSourceType.Device)
        //    {
        //        commandTree = CurrentDriver.GetElementFromWebElement(MobileBy.AccessibilityId("DeivceCmdTree"));
        //    }
        //    else
        //    {
        //        commandTree = CurrentDriver.GetElementFromWebElement(MobileBy.AccessibilityId("SystemCmdTree"));
        //    }

        //    //Console.WriteLine($"LeftClick on Text \"{groupName}\"");

        //    bool cmdListIsFocused = false;
        //    IWebElement groupTreeItem = null;
        //    while (groupTreeItem == null)
        //    {
        //        // Get the element that matching the given groupname directly by XPath (Faster)
        //        groupTreeItem = commandTree.GetElementFromWebElement(By.XPath($".//TreeItem[@ClassName='TreeViewItem']/Text[@Name='{groupName}']/parent::node()"), 3000);

        //        if (groupTreeItem == null && !cmdListIsFocused)
        //        {
        //            commandTree.LeftClick();
        //            cmdListIsFocused = true; // Set the flag to true after the left click on the command list
        //        }

        //        // If element if out of screen, press page down to find the element
        //        if (groupTreeItem == null)
        //        {
        //            Press(Keys.PageDown);
        //            Thread.Sleep(1);

        //            // If scroll to end of the command list, group item still not found, throw exception
        //            foreach (var cmdList in commandTree.GetElements(By.ClassName("TreeView")))
        //            {
        //                if (cmdList.GetAttribute("Scroll.VerticallyScrollable") == bool.FalseString)
        //                    continue;

        //                if (cmdList.GetAttribute("Scroll.VerticalScrollPercent") == "100")
        //                    throw new GroupNameNotExistedException(groupName);
        //            }
        //        }
        //    }


        //    // 2024/07/09, Adam, Expand the command group
        //    groupTreeItem.ExpandTreeView();

        //    //// Use attribute: "ExpandCollapse.ExpandCollapseState" to check the expand/collapse state, where: Expanded (1), Collapsed (0)
        //    //if (groupTreeItem.isElementCollapsed())
        //    //    groupTreeItem.DoubleClick();

        //    Console.WriteLine($"LeftClick on Text \"{cmdName}\"");

        //    // Get all elements, then find element that matching the given command name (Longer time required, can use)
        //    var cmdTreeItem = groupTreeItem.GetElements(By.XPath($".//TreeItem[@ClassName='TreeViewItem']"))
        //                                   .FirstOrDefault(e => e.GetFirstTextContent() == cmdName) 
        //                                   ?? throw new CommandNameNotExistedException(cmdName);

        //    //*[contains(@label,"text you want to find")]
        //    //var cmdTreeItem = groupTreeItem.GetElementFromWebElement(By.XPath($".//TreeItem[@ClassName='TreeViewItem']/Text[(@Name='{cmdName}')]/parent::node()"))
        //    //                               ?? throw new CommandNameNotExistedException(cmdName);

        //    //// Get the element that matching the given command name directly by XPath (Faster, but not executable)
        //    //var cmdTreeItem = groupTreeItem.GetElementFromWebElement(By.XPath($".//TreeItem[@ClassName='TreeViewItem']/Text[@Name='{cmdName}']/parent::node()"));

        //    // If element if out of screen, move to the element first
        //    while (bool.Parse(cmdTreeItem.GetAttribute("IsOffscreen")))
        //    {
        //        Press(Keys.PageDown);
        //        Thread.Sleep(50);
        //    }

        //    // Add the command
        //    for (int i = 0; i < addCount; i++)
        //        cmdTreeItem.DoubleClick();

        //    // Press left arrow key twice to Close the group tree view
        //    Press(Keys.Left);
        //    Press(Keys.Left);
        //}

        // Legacy method adding commands with command names

        public void AddCommandsBy(TestCmdGroupType cmdGrpType, params string[] cmdNames)
        {
            /* Legacy method of getting cmdtree by group name
            //IWebElement commandTree;
            //if (!GetCommandSourceType(groupName, out bool IsDevice))
            //    return;
            
            //if (IsDevice)
            //{
            //    commandTree = PP5IDEWindow.GetElementFromWebElement(MobileBy.AccessibilityId("DeivceCmdTree"));
            //}
            //else
            //{
            //    commandTree = PP5IDEWindow.GetElementFromWebElement(MobileBy.AccessibilityId("SystemCmdTree"));
            //}
            */

            IElement groupTreeItem = GetCmdGroupTreeItemByGroupName(cmdGrpType);
            AddCommandsBy(groupTreeItem, cmdNames, out IElement cmdToAdd);               // Add the command by given parameters
            TreeViewCollapseAll(cmdToAdd);                                                  // Press left arrow key twice to Close the group tree view

            /* Legacy methods 
            //// Use attribute: "ExpandCollapse.ExpandCollapseState" to check the expand/collapse state, where: Expanded (1), Collapsed (0)
            //if (groupTreeItem.isElementCollapsed())
            //    groupTreeItem.DoubleClick();

            // Get all elements, then find element that matching the given command name (Longer time required, can use)
            //var cmdTreeItems = groupTreeItem.GetElements(By.XPath($".//TreeItem[@ClassName='TreeViewItem']"));
            //HashSet<IWebElement> cmdTreeItemsHash = new HashSet<IWebElement>(cmdTreeItems);
            //var h = new HashSet<int>(Enumerable.Range(1, 17519).Select(i => i * 17519));

            //var cmdTreeItems = groupTreeItem.GetElements(By.XPath(".//TreeItem[@ClassName='TreeViewItem']"));

            // Get commands in the list
            //Stopwatch sw = Stopwatch.StartNew();
            //sw.Start();
            //var cmdNamesFound = cmdTreeItems.Select(e => e.GetSubElementText());
            //sw.Stop();
            //var GetSubElementTextTime = sw.ElapsedMilliseconds;
            //ArrayList cmdNamesFound = new ArrayList();
            //foreach (string text in cmdTreeItems.Select(e => e.GetSubElementText()))
            //    cmdNamesFound.Add(text);

            //List<string> cmdNamesFound = new List<string>(cmdTreeItems.Select(e => e.GetSubElementText()));

            //for (int i = 0; i < cmdTreeItems.Count(); i++)
            //{
            //    ((WindowsElement)cmdTreeItems[i]).SetCacheValues(new Dictionary<string, object> { ["text"] = cmdNamesFound.ElementAt(i) });
            //}

            //// Convert cmdNamesFound to a HashSet for faster lookups using Distinct and HashSet constructor
            //HashSet<string> cmdNamesFoundSet = new HashSet<string>(cmdTreeItems.Select(e => e.GetSubElementText()), new StringEqualityComparer());

            //var cmdTreeItems = groupTreeItem.GetTreeViewItems();
            //foreach (string cmdName in cmdNames.OrderBy(n => n))
            //{
            //    Console.WriteLine($"LeftClick on Text \"{cmdName}\"");

            //    //sw.Restart();
            //    //if (!cmdNamesFound.Contains(cmdName))
            //    //    throw new CommandNameNotExistedException(cmdName);
            //    //sw.Stop();
            //    //var sListContainsTime = sw.ElapsedMilliseconds;

            //    // Reuse the previously found elements
            //    //IWebElement cmdTreeItem = cmdTreeItems.FirstOrDefault(e => e.GetSubElementText() == cmdName);
            //    //IWebElement cmdTreeItem = cmdTreeItems.FirstOrDefault(e => ((WindowsElement)e).Text == cmdName);
            //    //IWebElement cmdTreeItem = cmdTreeItems[cmdNamesFound.IndexOf(cmdName)];

            //    IWebElement cmdTreeItem = cmdTreeItems.FirstOrDefault(e => e.GetFirstTextContent() == cmdName)
            //                                          ?? throw new CommandNameNotExistedException(cmdName);

            //    // If the element is out of screen, move to the element first
            //    while (bool.Parse(cmdTreeItem.GetAttribute("IsOffscreen")))
            //    {
            //        Press(Keys.PageDown);
            //        Thread.Sleep(1);

            //        //// Find the element again if it shows up
            //        //cmdTreeItem = groupTreeItem.GetElementFromWebElement(By.XPath($".//TreeItem[@ClassName='TreeViewItem']/Text[@Name='{cmdName}']/parent::node()"), 3000);
            //    }

            //    // Add the command
            //    cmdTreeItem.DoubleClick();
            //}

            //// Close the group tree view
            //while (!groupTreeItem.Selected)
            //{
            //    PressUp();
            //    Thread.Sleep(1);
            //}
            */
        }

        // Legacy method adding commands with command indeces in the given group
        public void AddCommandsBy(TestCmdGroupType cmdGrpType, params int[] cmdNumbers)
        {
            /* Legacy method of getting cmdtree by group name
            //IWebElement commandTree;
            //if (!GetCommandSourceType(groupName, out bool IsDevice))
            //    return;
            
            //if (IsDevice)
            //{
            //    commandTree = PP5IDEWindow.GetElementFromWebElement(MobileBy.AccessibilityId("DeivceCmdTree"));
            //}
            //else
            //{
            //    commandTree = PP5IDEWindow.GetElementFromWebElement(MobileBy.AccessibilityId("SystemCmdTree"));
            //}
            */

            if (cmdNumbers.Any(n => n <= 0))                                                // Check cmdNumbers has no negative values
                throw new CommandNumberOutOfRangeException(cmdNumbers.Min());

            IElement groupTreeItem = GetCmdGroupTreeItemByGroupName(cmdGrpType);
            AddCommandsBy(groupTreeItem, cmdNumbers, out IElement cmdToAdd);             // Add the command by given parameters
            TreeViewCollapseAll(cmdToAdd);                                                  // Press left arrow key twice to Close the group tree view

            /* Legacy methods  
            //// Get all elements, then find element that matching the given groupname (Longer time required)
            //var groupTreeItem = commandTree.GetElements(By.XPath($".//TreeItem[@ClassName='TreeViewItem']")).ToList()
            //                   .Find(e => e.GetSubElementText() == groupName);

            //// Use attribute: "ExpandCollapse.ExpandCollapseState" to check the expand/collapse state, where: Expanded (1), Collapsed (0)
            //if (groupTreeItem.isElementCollapsed())
            //    groupTreeItem.DoubleClick();

            //IWebElement cmdToAdd = null;

            ////var cmdTreeItems = groupTreeItem.GetElements(By.XPath($".//TreeItem[@ClassName='TreeViewItem']"));
            //var cmdTreeItems = groupTreeItem.GetTreeViewItems();
            //if (cmdNumbers.Any(n => n > cmdTreeItems.Count))
            //    throw new CommandNumberOutOfRangeException(cmdNumbers.Max().ToString());

            ////cmdNumbers.ToList().ForEach(num =>
            ////{
            ////    if (cmdTreeItems.Count == 0)
            ////        return;
            ////    if (num >= cmdTreeItems.Count || num <= -1)
            ////        throw new CommandNumberOutOfRangeException(num.ToString());
            ////});

            ////if (cmdNumbers.ToList().Any(num => num >= cmdTreeItems.Count || num <= 0))
            ////{
            ////    throw new CommandNumberOutOfRangeException(cmdNumbers.ToList().First().ToString());
            ////}

            //foreach (int cmdNumber in cmdNumbers.OrderBy(n => n))
            //{
            //    cmdToAdd = cmdNumber == -1 ? cmdTreeItems.Last() : cmdTreeItems[cmdNumber - 1];

            //    //var cmdTreeItem = groupTreeItem.GetElementFromWebElement(By.XPath($"(.//TreeItem[@ClassName='TreeViewItem'])[{cmdNumber + 1}]"));
            //    Console.WriteLine($"LeftClick on Text \"{cmdToAdd.GetFirstTextContent()}\"");

            //    //// If element if out of screen, move to the element first
            //    while (bool.Parse(cmdToAdd.GetAttribute("IsOffscreen")))
            //    {
            //        Press(Keys.PageDown);
            //        Thread.Sleep(1);
            //    }

            //    // Add the command
            //    cmdToAdd.DoubleClick();
            //}

            // Close the group tree view
            //while (!groupTreeItem.Selected)
            //{
            //    PressUp();
            //    Thread.Sleep(1);
            //}
            */
        }


        // Base method to add command by command indeces
        public void AddCommandsBy(IElement groupTreeItem, int[] cmdNumbers, out IElement cmdToAdd)
        {
            var cmdTreeItems = groupTreeItem.GetTreeViewItems();
            if (cmdNumbers.Any(n => n > cmdTreeItems.Count))
                throw new CommandNumberOutOfRangeException(cmdNumbers.Max());

            AddCommands(cmdTreeItems, cmdNumbers);
            cmdToAdd = cmdTreeItems[cmdNumbers[0]];
        }

        // Base method to add command by command Names
        public void AddCommandsBy(IElement groupTreeItem, string[] cmdNames, out IElement cmdToAdd)
        {
            var cmdTreeItems = groupTreeItem.GetTreeViewItems();
            //var cmdNamesActual = groupTreeItem.GetSpecificChildrenContentOfControlType(ElementControlType.TextBlock, e => e.GetFirstTextContent() != null);
            //var cmdNamesActual = cmdTreeItems.ToArray().Select(e => e.GetFirstTextContent()).ToArray();
            List<string> cmdNamesActual = QueryCommandNames(groupTreeItem.GetFirstTextContent());

            if (cmdNames.Except(cmdNamesActual).Count() > 0)
                throw new CommandNameNotExistedException(cmdNames.Except(cmdNamesActual).First());

            var cmdNumbers = cmdNames.Select(x => cmdNamesActual.ToList().IndexOf(x) + 1).ToArray();
            AddCommands(cmdTreeItems, cmdNumbers);
            cmdToAdd = cmdTreeItems[cmdNumbers[0]];

            /* Legacy using cmdNames
            //foreach (string cmdName in cmdNames.OrderBy(n => n))
            //{
            //    Console.WriteLine($"LeftClick on Text \"{cmdName}\"");

            //    //sw.Restart();
            //    //if (!cmdNamesFound.Contains(cmdName))
            //    //    throw new CommandNameNotExistedException(cmdName);
            //    //sw.Stop();
            //    //var sListContainsTime = sw.ElapsedMilliseconds;

            //    cmdToAdd = cmdTreeItems.FirstOrDefault(e => e.GetFirstTextContent() == cmdName)
            //                           ?? throw new CommandNameNotExistedException(cmdName);

            //    // If the element is out of screen, move to the element first
            //    while (bool.Parse(cmdToAdd.GetAttribute("IsOffscreen")))
            //    {
            //        Press(Keys.PageDown);
            //        Thread.Sleep(1);
            //    }

            //    // Add the command
            //    cmdToAdd.DoubleClick();
            //}
            */
        }

        // Base method to get command by command Name
        public IElement GetCommandBy(TestCmdGroupType cmdGrpType, string cmdName, bool collapseTreeView = false)
        {
            IElement groupTreeItem = GetCmdGroupTreeItemByGroupName(cmdGrpType);
            IElement cmdToAdd = GetCommand(groupTreeItem, GetGroupNameByEnum(cmdGrpType), cmdName);                                       // Add the command by given parameters

            if (collapseTreeView)
                TreeViewCollapseAll(cmdToAdd);                                                                          // Press left arrow key twice to Close the group tree view

            return cmdToAdd;
        }

        // Base method to get command by command Number
        public IElement GetCommandBy(TestCmdGroupType cmdGrpType, int cmdNumber, bool collapseTreeView = false)
        {
            IElement groupTreeItem = GetCmdGroupTreeItemByGroupName(cmdGrpType);
            IElement cmdToAdd = GetCommand(groupTreeItem, cmdNumber);                                                   // Add the command by given parameters

            if (collapseTreeView)
                TreeViewCollapseAll(cmdToAdd);                                                                          // Press left arrow key twice to Close the group tree view

            return cmdToAdd;
        }

        public void AddCommandBy(IElement groupTreeItem, string cmdName, int addCount, out IElement cmdToAdd)
        {
            // Get all elements, then find element that matching the given command name (Longer time required, can use)
            int cmdNumber;
            var cmdTreeItems = groupTreeItem.GetTreeViewItems();
            List<string> cmdNamesActual = QueryCommandNames(groupTreeItem.GetFirstTextContent());
            if (!cmdNamesActual.Contains(cmdName))
                throw new CommandNameNotExistedException(cmdName);

            cmdNumber = cmdNamesActual.IndexOf(cmdName) + 1;
            AddCommand(cmdTreeItems, cmdNumber, addCount);
            cmdToAdd = cmdTreeItems[cmdNumber];
        }

        public void AddCommandBy(IElement groupTreeItem, int cmdNumber, int addCount, out IElement cmdToAdd)
        {
            cmdToAdd = null;
            var cmdTreeItems = groupTreeItem.GetTreeViewItems();

            if (cmdTreeItems.Count == 0)
                return;
            if (cmdNumber > cmdTreeItems.Count || cmdNumber <= -1)
                throw new CommandNumberOutOfRangeException(cmdNumber);

            AddCommand(cmdTreeItems, cmdNumber, addCount);
            cmdToAdd = cmdTreeItems[cmdNumber];
            /*
            //cmdToAdd = cmdNumber == -1 ? cmdTreeItems.Last() : cmdTreeItems[cmdNumber];

            ////var cmdTreeItem = groupTreeItem.GetElementFromWebElement(By.XPath($"(.//TreeItem[@ClassName='TreeViewItem'])[{cmdNumber + 1}]"));
            ////Console.WriteLine($"LeftClick on Text \"{cmdToAdd.GetFirstTextContent()}\"");

            ////// If element if out of screen, move to the element first
            //while (bool.Parse(cmdToAdd.GetAttribute("IsOffscreen")))
            //{
            //    Press(Keys.PageDown);
            //    Thread.Sleep(50);
            //}

            //// Add the command
            //if (isAddCommand)
            //{
            //    for (int i = 0; i < addCount; i++)
            //        cmdToAdd.DoubleClick();
            //}
            */
        }

        public void SelectCommandBy(IElement groupTreeItem, int cmdNumber, out IElement cmdToAdd)
        {
            cmdToAdd = null;
            var cmdTreeItems = groupTreeItem.GetTreeViewItems();

            if (cmdTreeItems.Count == 0)
                return;
            if (cmdNumber > cmdTreeItems.Count || cmdNumber <= -1)
                throw new CommandNumberOutOfRangeException(cmdNumber);

            SelectCommand(cmdTreeItems, cmdNumber);
        }

        private IElement GetCommand(IElement groupTreeItem, string groupName, string cmdName)
        {
            if (!HasCommand(groupName, cmdName))
                throw new CommandNameNotExistedException(cmdName);

            var cmdTreeItems = groupTreeItem.GetTreeViewItems();
            //Logger.LogMessage($"In GetCommand method, cmdTreeItems.Count:{cmdTreeItems.Count}");
            int cmdNumber = QueryCommandNames(groupName).IndexOf(cmdName) + 1;
            //Logger.LogMessage($"In GetCommand method, cmdNumber:{cmdNumber}");
            return GetCommand(cmdTreeItems, cmdNumber);
        }

        public IElement GetCommand(IElement groupTreeItem, int cmdNumber)
        {
            var cmdTreeItems = groupTreeItem.GetTreeViewItems();

            if (cmdTreeItems.Count == 0)
                return null;
            if (cmdNumber > cmdTreeItems.Count || cmdNumber <= -1)
                throw new CommandNumberOutOfRangeException(cmdNumber);

            return GetCommand(cmdTreeItems, cmdNumber);
        }

        // Common base method to add command by command indeces
        private void AddCommands(ReadOnlyCollection<IElement> cmdTreeItems, int[] cmdNumbers)
        {
            foreach (int cmdNumber in cmdNumbers.OrderBy(n => n))
            {
                AddCommand(cmdTreeItems, cmdNumber, 1);
            }
        }

        // Base method to add a single command (used by all other AddCommand/AddCommands methods)
        private void AddCommand(ReadOnlyCollection<IElement> cmdTreeItems, int cmdNumber, int addCount = 1)
        {
            IElement cmdToAdd = cmdNumber == -1 ? cmdTreeItems.Last() : cmdTreeItems[cmdNumber];

            //// If element if out of screen, move to the element first
            while (bool.Parse(cmdToAdd.GetAttribute("IsOffscreen")))
            {
                Press(Keys.PageDown);
                Thread.Sleep(1);
            }

            // Add the command
            for (int i = 0; i < addCount; i++)
                cmdToAdd.DoubleClick();
        }

        // Base method to select a single command (used by all other AddCommand/AddCommands methods)
        private void SelectCommand(ReadOnlyCollection<IElement> cmdTreeItems, int cmdNumber)
        {
            IElement cmdToAdd = cmdNumber == -1 ? cmdTreeItems.Last() : cmdTreeItems[cmdNumber];

            //// If element if out of screen, move to the element first
            while (bool.Parse(cmdToAdd.GetAttribute("IsOffscreen")))
            {
                Press(Keys.PageDown);
                Thread.Sleep(1);
            }

            // Select the command
            cmdToAdd.LeftClick();
        }

        // Base method to get a single command (used by all other AddCommand/AddCommands methods)
        private IElement GetCommand(ReadOnlyCollection<IElement> cmdTreeItems, int cmdNumber)
        {
            IElement cmdToAdd = cmdNumber == -1 ? cmdTreeItems.Last() : cmdTreeItems[cmdNumber];

            //// If element if out of screen, move to the element first
            while (bool.Parse(cmdToAdd.GetAttribute("IsOffscreen")))
            {
                Press(Keys.PageDown);
                Thread.Sleep(1);
            }

            return cmdToAdd;
        }

        private IElement GetCommandIsSelected(ReadOnlyCollection<IElement> cmdTreeItems)
        {
            IElement cmdToAdd = cmdTreeItems.FirstOrDefault(cmd => cmd.Selected);

            //// If element if out of screen, move to the element first
            while (bool.Parse(cmdToAdd.GetAttribute("IsOffscreen")))
            {
                Press(Keys.PageDown);
                Thread.Sleep(1);
            }

            return cmdToAdd;
        }


        public IElement GetCmdGroupTreeItemByGroupName(TestCmdGroupType cmdGrpType)
        {
            string groupName = GetGroupNameByEnum(cmdGrpType);
            IElement commandTree = GetCommandTreeByGroupName(groupName);
            if (commandTree == null)
                return null;

            CommandTreeViewScrollToTop(commandTree);                                                                    // Scroll to the top if not
            ExpandCommandGroup(commandTree, groupName, out IElement groupTreeItem);                                     // Expand the group item tree
            return groupTreeItem;
        }

        public IElement GetCommandTreeByGroupName(string groupName)
        {
            IElement commandTree;
            if (!GetCommandSourceType(groupName, out bool IsDevice))
                throw new GroupNameNotExistedException(groupName);

            if (IsDevice)
            {
                commandTree = PP5IDEWindow.GetElementFromPP5Element(PP5By.Id("DeivceCmdTree"));
            }
            else
            {
                commandTree = PP5IDEWindow.GetElementFromPP5Element(PP5By.Id("SystemCmdTree"));
            }
            return commandTree;
        }

        public void CommandTreeViewScrollToTop(IElement commandTree)
        {
            // Scroll to the top if not
            //string ScrollBarPosPerc = commandTree.GetAttribute("Scroll.VerticalScrollPercent");
            if (commandTree.GetAttribute("Scroll.VerticalScrollPercent") != "-1")
            {
                while (double.Parse(commandTree.GetAttribute("Scroll.VerticalScrollPercent")) >= 0.00001)
                {
                    //MoveToElementAndLeftClick(commandTree);
                    commandTree.LeftClick();
                    Press(Keys.Home);
                }
            }
        }

        private void SearchForCommandGroup(IElement commandTree, string groupName, out IElement groupTreeItem)
        {
            commandTree.LeftClick();
            do
            {
                // Get the group element by groupname
                groupTreeItem = commandTree.GetTreeViewItems()
                                           .FirstOrDefault(e => e.GetTextElement(groupName)?.Text == groupName);

                // If element if out of screen, press page down to find the element
                if (groupTreeItem == null)
                    return;
                else if (!groupTreeItem.Displayed)
                {
                    Press(Keys.PageDown);
                    Thread.Sleep(1);
                }
            } while (bool.Parse(groupTreeItem?.GetAttribute("IsOffscreen")));
        }

        private void ExpandCommandGroup(IElement commandTree, string groupName, out IElement groupTreeItem)
        {
            SearchForCommandGroup(commandTree, groupName, out groupTreeItem);

            // 2024/07/09, Adam, Expand the command group
            if (groupTreeItem != null)
                groupTreeItem.ExpandTreeView();
        }

        public ReadOnlyCollection<IElement> GetExpandedCommandGroup(IElement commandTree)
        {
            // Get the group element by groupname
            IElement groupTreeItem = commandTree.GetTreeViewItems()
                                                   .FirstOrDefault(e => e.isElementExpanded());
            return groupTreeItem.GetTreeViewItems();
        }

        public void AddTIBy(string groupName, int tiIndex = 1, TestItemSourceType tiType = TestItemSourceType.System, int addCount = 1)
        {
            IElement commandTree;
            if (tiType == TestItemSourceType.System)
            {
                TPExecuteAction(TPAction.SwitchToSystemTIPage);
                commandTree = CurrentDriver.GetExtendedElementBySingleWithRetry(PP5By.ClassName("SysUUTCmdTreeView")).GetFirstTreeViewElement();
            }
            else
            {
                TPExecuteAction(TPAction.SwitchToUserDefinedTIPage);
                commandTree = CurrentDriver.GetExtendedElementBySingleWithRetry(PP5By.ClassName("UserUUTCmdTreeView")).GetFirstTreeViewElement();
            }

            Logger.LogMessage($"LeftClick on Text \"{groupName}\"");

            bool cmdListIsFocused = false;
            IElement groupTreeItem = null;
            while (groupTreeItem == null)
            {
                // Get the element that matching the given groupname directly by XPath (Faster)
                groupTreeItem = commandTree.GetExtendedElementBySingleWithRetry(PP5By.XPath($".//TreeItem[@ClassName='TreeViewItem']/Text[@Name='{groupName}']/parent::node()"), 3000);

                if (groupTreeItem == null && !cmdListIsFocused)
                {
                    commandTree.LeftClick();
                    cmdListIsFocused = true; // Set the flag to true after the left click on the command list
                }

                // If element if out of screen, press page down to find the element
                if (groupTreeItem == null)
                {
                    Press(Keys.PageDown);
                    Thread.Sleep(1);

                    // If scroll to end of the command list, group item still not found, throw exception
                    foreach (var cmdList in commandTree.GetWebElementsFromWebElement(By.ClassName("TreeView")))
                    {
                        if (cmdList.GetAttribute("Scroll.VerticallyScrollable") == bool.FalseString)
                            continue;

                        if (cmdList.GetAttribute("Scroll.VerticalScrollPercent") == "100")
                            throw new GroupNameNotExistedException(groupName);
                    }
                }
            }

            // 2024/07/09, Adam, Expand the command group
            groupTreeItem.ExpandTreeView();

            //// Get all elements, then find element that matching the given groupname (Longer time required)
            //var groupTreeItem = commandTree.GetElements(By.XPath($".//TreeItem[@ClassName='TreeViewItem']")).ToList()
            //                   .Find(e => e.GetSubElementText() == groupName);

            //// Use attribute: "ExpandCollapse.ExpandCollapseState" to check the expand/collapse state, where: Expanded (1), Collapsed (0)
            //if (groupTreeItem.isElementCollapsed())
            //    groupTreeItem.DoubleClick();

            var cmdTreeItems = groupTreeItem.GetExtendedElements(PP5By.XPath($".//TreeItem[@ClassName='TreeViewItem']"));

            if (cmdTreeItems.Count == 0)
                return;
            if (tiIndex > cmdTreeItems.Count || tiIndex < -1 || tiIndex == 0)
                throw new ArgumentOutOfRangeException(tiIndex.ToString());

            IElement tiToAdd = tiIndex == -1 ? cmdTreeItems.Last() : cmdTreeItems[tiIndex];

            //var cmdTreeItem = groupTreeItem.GetElementFromWebElement(By.XPath($"(.//TreeItem[@ClassName='TreeViewItem'])[{cmdNumber + 1}]"));
            Logger.LogMessage($"LeftClick on Text \"{tiToAdd.GetFirstTextContent()}\"");

            //// If element if out of screen, move to the element first
            while (bool.Parse(tiToAdd.GetAttribute("IsOffscreen")))
            {
                Press(Keys.PageDown);
                Thread.Sleep(50);
            }

            // Add the TI
            for (int i = 0; i < addCount; i++)
                tiToAdd.DoubleClick();

            // Press left arrow key twice to Close the group tree view
            Press(Keys.Left);
            Press(Keys.Left);
        }

        public void AddTIBy(string groupName, int tiIndex = 1, int addCount = 1)
        {
            // Switch to System or user-defined page by checking groupname

            IElement commandTree = CurrentDriver.GetExtendedElement(By.ClassName("DeivceCmdTree"));

            Logger.LogMessage($"LeftClick on Text \"{groupName}\"");

            bool cmdListIsFocused = false;
            IElement groupTreeItem = null;
            while (groupTreeItem == null)
            {
                // Get the element that matching the given groupname directly by XPath (Faster)
                groupTreeItem = commandTree.GetExtendedElementBySingleWithRetry(PP5By.XPath($".//TreeItem[@ClassName='TreeViewItem']/Text[@Name='{groupName}']/parent::node()"), 3000);

                if (groupTreeItem == null && !cmdListIsFocused)
                {
                    commandTree.LeftClick();
                    cmdListIsFocused = true; // Set the flag to true after the left click on the command list
                }

                // If element if out of screen, press page down to find the element
                if (groupTreeItem == null)
                {
                    Press(Keys.PageDown);
                    Thread.Sleep(1);

                    // If scroll to end of the command list, group item still not found, throw exception
                    foreach (var cmdList in commandTree.GetWebElementsFromWebElement(By.ClassName("TreeView")))
                    {
                        if (cmdList.GetAttribute("Scroll.VerticallyScrollable") == bool.FalseString)
                            continue;

                        if (cmdList.GetAttribute("Scroll.VerticalScrollPercent") == "100")
                            throw new GroupNameNotExistedException(groupName);
                    }
                }
            }

            // 2024/07/09, Adam, Expand the command group
            groupTreeItem.ExpandTreeView();

            //// Get all elements, then find element that matching the given groupname (Longer time required)
            //var groupTreeItem = commandTree.GetElements(By.XPath($".//TreeItem[@ClassName='TreeViewItem']")).ToList()
            //                   .Find(e => e.GetSubElementText() == groupName);

            //// Use attribute: "ExpandCollapse.ExpandCollapseState" to check the expand/collapse state, where: Expanded (1), Collapsed (0)
            //if (groupTreeItem.isElementCollapsed())
            //    groupTreeItem.DoubleClick();

            var cmdTreeItems = groupTreeItem.GetExtendedElements(PP5By.XPath($".//TreeItem[@ClassName='TreeViewItem']"));

            if (cmdTreeItems.Count == 0)
                return;
            if (tiIndex > cmdTreeItems.Count || tiIndex < -1 || tiIndex == 0)
                throw new ArgumentOutOfRangeException(tiIndex.ToString());

            IElement tiToAdd = tiIndex == -1 ? cmdTreeItems.Last() : cmdTreeItems[tiIndex];

            //var cmdTreeItem = groupTreeItem.GetElementFromWebElement(By.XPath($"(.//TreeItem[@ClassName='TreeViewItem'])[{cmdNumber + 1}]"));
            Logger.LogMessage($"LeftClick on Text \"{tiToAdd.GetFirstTextContent()}\"");

            //// If element if out of screen, move to the element first
            while (bool.Parse(tiToAdd.GetAttribute("IsOffscreen")))
            {
                Press(Keys.PageDown);
                Thread.Sleep(50);
            }

            // Add the TI
            for (int i = 0; i < addCount; i++)
                tiToAdd.DoubleClick();

            // Press left arrow key twice to Close the group tree view
            Press(Keys.Left);
            Press(Keys.Left);
        }

        public void SelectColorSettingItem(IElement ColorSettingPage, ColorSettingPageType csPageType, TestCmdGroupType cmdGrpType, int idx = 1, bool collapseTreeView = false)
        {
            IElement colorSettingItem = SelectColorSettingItem(ColorSettingPage, csPageType, cmdGrpType, idx);
            if (collapseTreeView)
                TreeViewCollapseAll(colorSettingItem);
        }

        public IElement SelectColorSettingItem(IElement ColorSettingPage, ColorSettingPageType csPageType, TestCmdGroupType cmdGrpType, int idx = 1)
        {
            IElement colorSettingItem = GetColorSettingItem(ColorSettingPage, csPageType, cmdGrpType, idx);
            colorSettingItem.GetFirstTextElement().LeftClick();
            return colorSettingItem;
        }


        public IElement GetColorSettingItem(IElement ColorSettingPage, ColorSettingPageType csPageType, TestCmdGroupType cmdGrpType, int idx = 1)
        {
            IElement commandTree = ColorSettingPage.GetExtendedElement(PP5By.Id(csPageType.GetDescription()));
            //Console.WriteLine($"LeftClick on Text \"{groupName}\"");

            CommandTreeViewScrollToTop(commandTree);                                                            // Scroll to the top if not
            ExpandCommandGroup(commandTree, GetGroupNameByEnum(cmdGrpType), out IElement groupTreeItem);        // Expand the group item tree
            return GetCommand(groupTreeItem, idx);                                                              // Add the command by given parameters

            /* Legacy methods  
            //bool cmdListIsFocused = false;
            //IWebElement groupTreeItem = null;
            //while (groupTreeItem == null)
            //{
            //    // Get the element that matching the given groupname directly by XPath (Faster)
            //    groupTreeItem = commandTree.GetElementFromWebElement(By.XPath($".//TreeItem[@ClassName='TreeViewItem']/Text[@Name='{groupName}']/parent::node()"), 3000);

            //    if (groupTreeItem == null && !cmdListIsFocused)
            //    {
            //        commandTree.LeftClick();
            //        cmdListIsFocused = true; // Set the flag to true after the left click on the command list
            //    }

            //    // If element if out of screen, press page down to find the element
            //    if (groupTreeItem == null)
            //    {
            //        Press(Keys.PageDown);
            //        Thread.Sleep(1);

            //        // If scroll to end of the command list, group item still not found, throw exception
            //        foreach (var cmdList in commandTree.GetElements(By.ClassName("TreeView")))
            //        {
            //            if (cmdList.GetAttribute("Scroll.VerticallyScrollable") == bool.FalseString)
            //                continue;

            //            if (cmdList.GetAttribute("Scroll.VerticalScrollPercent") == "100")
            //                throw new GroupNameNotExistedException(groupName);
            //        }
            //    }
            //}

            //// Use attribute: "ExpandCollapse.ExpandCollapseState" to check the expand/collapse state, where: Expanded (1), Collapsed (0)
            //if (groupTreeItem.isElementCollapsed())
            //    groupTreeItem.GetFirstTextElement().DoubleClick();

            //var cmdTreeItems = groupTreeItem.GetElements(By.XPath($".//TreeItem[@ClassName='TreeViewItem']"));

            //var cmdTreeItems = groupTreeItem.GetTreeViewItems();

            //if (cmdTreeItems.Count == 0)
            //    return null;
            //if (idx > cmdTreeItems.Count || idx < -1 || idx == 0)
            //    throw new ArgumentOutOfRangeException(idx.ToString());

            //IWebElement itemToSetColor = idx == -1 ? cmdTreeItems.Last() : cmdTreeItems[idx];

            ////var cmdTreeItem = groupTreeItem.GetElementFromWebElement(By.XPath($"(.//TreeItem[@ClassName='TreeViewItem'])[{cmdNumber + 1}]"));
            ////Console.WriteLine($"LeftClick on Text \"{itemToSetColor.GetFirstTextContent()}\"");

            ////// If element if out of screen, move to the element first
            //while (bool.Parse(itemToSetColor.GetAttribute("IsOffscreen")))
            //{
            //    Press(Keys.PageDown);
            //    Thread.Sleep(50);
            //}

            ////// Click on the item
            ////itemToSetColor.LeftClick();

            //return itemToSetColor;
            */
        }


        public void TreeViewCollapseAll(IElement treeItemEle)
        {
            if (treeItemEle.CanExpandOrCollapse)
            {
                Press(Keys.Left);       // Press left arrow key to move selection one tree level up
                Press(Keys.Left);       // Press left arrow key again to collapse current selection tree view
            }
        }

        #endregion

        //public IWebElement GetComboBoxElementByID(string comboBoxID)
        //{
        //    return CurrentDriver.GetElementFromWebElement(MobileBy.AccessibilityId(comboBoxID));
        //    //return CurrentDriver.GetElementFromWebElement(By.XPath($".//ComboBox[@AutomationId=\"{comboBoxID}\"]"));
        //    //return CurrentDriver.GetElements(By.ClassName(ElementControlType.ComboBox.ToString()))
        //    //                    .FirstOrDefault(e => e.CheckElementHasNameOrId() == comboBoxID);
        //}

        //public IWebElement GetListBoxElementByID(string listBoxID)
        //{
        //    return CurrentDriver.GetElementFromWebElement(MobileBy.AccessibilityId(listBoxID));
        //    //return CurrentDriver.GetElementFromWebElement(By.XPath($".//List[@AutomationId=\"{listBoxID}\"]"));
        //    //return CurrentDriver.GetElements(By.ClassName(ElementControlType.ListBox.ToString()))
        //    //                    .FirstOrDefault(e => e.CheckElementHasNameOrId() == listBoxID);
        //}

        //public IWebElement GetCustomComboBoxElementByID(string comboBoxID)
        //{
        //    //return CurrentDriver.GetElementFromWebElement(MobileBy.AccessibilityId(listBoxID));
        //    return CurrentDriver.GetElementFromWebElement(By.XPath($".//Pane[@AutomationId=\"{comboBoxID}\"]"));
        //    //return CurrentDriver.GetElements(By.ClassName(ElementControlType.ListBox.ToString()))
        //    //                    .FirstOrDefault(e => e.CheckElementHasNameOrId() == listBoxID);
        //}

        public void GetComboBoxItems(string comboBoxID, out ReadOnlyCollection<IWebElement> cmbItems)
        {
            //IWebElement comboBox = CurrentDriver.GetElementFromWebElement(MobileBy.AccessibilityId(comboBoxID));
            //if (comboBox.TagName == ElementControlType.ComboBox.GetDescription())
            //    return comboBox.GetElements(By.ClassName("ComboBoxItem"));
            //else if (comboBox.TagName == ElementControlType.ListBox.GetDescription())
            //    return comboBox.GetElements(By.ClassName("ListBoxItem"));
            //else
            //    return null;
            
            PP5IDEWindow.GetWebElementFromWebElement(MobileBy.AccessibilityId(comboBoxID)).LeftClick();
            cmbItems = PP5IDEWindow.GetWebElementFromWebElement(By.ClassName("Popup"))
                                    .GetWebElementsFromWebElement(By.ClassName("ListBoxItem"));
        }

        // 先暫時分成兩個方法(給/不給comboBoxID)
        // 給comboBoxID作法: 用comboBoxID定位combobox後再做動作
        public void ComboBoxSelectByName(string comboBoxID, string name)
        {
            //IWebElement comboBox = PP5IDEWindow.GetElementFromWebElement(MobileBy.AccessibilityId(comboBoxID));
            ////IWebElement comboBox = GetComboBoxElementByID(comboBoxID);

            //comboBox.SelectComboBoxItemByName2(name);
            PP5IDEWindow.GetExtendedElement(PP5By.Id(comboBoxID))
                        .SelectComboBoxItemByName2(name);
        }

        public void ComboBoxSelectByIndex(string comboBoxID, int index, bool supportKeyInputSearch = true)
        {
            IElement comboBox = PP5IDEWindow.GetExtendedElement(PP5By.Id(comboBoxID));
            //IWebElement comboBox = GetComboBoxElementByID(comboBoxID);
            if (comboBox.isElementCollapsed())
                comboBox.LeftClick();
            comboBox.SelectComboBoxItemByIndex(index, supportKeyInputSearch);
        }

        //public void ListBoxSelectByName(string listBoxID, string name)
        //{
        //    IWebElement listBox = CurrentDriver.GetElementFromWebElement(MobileBy.AccessibilityId(listBoxID));
        //    //IWebElement listBox = GetListBoxElementByID(listBoxID);
        //    if (listBox.isElementCollapsed())
        //        listBox.LeftClick();
        //    ListBoxSelectByName(listBox, name);
        //}

        //public void ListBoxSelectByIndex(string listBoxID, int index)
        //{
        //    IWebElement listBox = CurrentDriver.GetElementFromWebElement(MobileBy.AccessibilityId(listBoxID));
        //    //IWebElement listBox = GetListBoxElementByID(listBoxID);
        //    if (listBox.isElementCollapsed())
        //        listBox.LeftClick();
        //    ListBoxSelectByIndex(listBox, index);
        //}

        //public void ListBoxSelectByName(IWebElement listBox, string name)
        //{
        //    if (CheckComboBoxHasItemByName(listBox, name))
        //    {
        //        SendComboKeys(name, Keys.Enter);
        //    }
        //}

        //public void ListBoxSelectByIndex(IWebElement listBox, int index)
        //{
        //    if (GetComboBoxItems(listBox).Count() >= index + 1)
        //    {
        //        var listBoxItems = GetComboBoxItems(listBox);
        //        string name = listBoxItems[index].Text;
        //        listBox.SendComboKeys(name, Keys.Enter);
        //    }
        //}



        //public int GetComboBoxSelectionIndex(string comboBoxID)
        //{
        //    IWebElement comboBox = CurrentDriver.GetElementFromWebElement(MobileBy.AccessibilityId(comboBoxID));
        //    //IWebElement comboBox = GetComboBoxElementByID(comboBoxID);

        //    string selectionName = comboBox.GetAttribute("Selection.Selection");
        //    comboBox.DoubleClickWithDelay(10);
        //    comboBox.GetComboBoxItems(out ReadOnlyCollection<IWebElement> cmbItems);
        //    return cmbItems.Select(e => e.Text).ToList().IndexOf(selectionName);
        //}

        //public bool CheckListBoxHasItemByName(IWebElement listBox, string name)
        //{
        //    return GetComboBoxItems(listBox).Select(e => e.Text).Contains(name);
        //}

        // cell headers: "ShowName", "CallName", "DisplayedType", "DisplayedEditType", "Minimum",
        //               "Maximum", "Default", "Format", "DisplayedEnum"
        // DataGridByInfo classNames: PGGrid, CndGrid, RstGrid, TmpGrid, GlbGrid, DefectCodeGrid, ParameterGrid, LoginGrid
        //public void SaveGridTable(IWebElement gridTableElement, DataTableAutoIDType DataGridType)
        //{
        //    //WindowsElement HeaderPanel = CurrentDriver.GetElementFromWebElement(MobileBy.AccessibilityId(DataGridAutomationID)).GetElementFromWebElement(By.Name("HeaderPanel"));

        //    //List<string> headers = HeaderPanel.GetChildElements().Where(ape => ape.TagName == "ControlType.HeaderItem")
        //    //                                                     .Select(tg => ((WindowsElement)tg).GetSubElementText())
        //    //                                                     .ToList();

        //    //WindowsElement DataPanel = CurrentDriver.GetElementFromWebElement(MobileBy.AccessibilityId(DataGridAutomationID))
        //    //                                        .GetElementFromWebElement(By.Name("DataPanel"));
        //    //ReadOnlyCollection<AppiumWebElement> rows = DataPanel.GetChildElements();

        //    //List<string> rowsss = CurrentDriver.GetElementFromWebElement(MobileBy.AccessibilityId(DataGridAutomationID))
        //    //                                   .GetElementFromWebElement(By.Name("DataPanel"))
        //    //                                   .GetChildElements().Select(e => e.TagName).ToList();

        //    //List<string> headers = CurrentDriver.GetElementFromWebElement(MobileBy.AccessibilityId(DataGridAutomationID))
        //    //                                    .GetElements(By.XPath(".//HeaderItem"))
        //    //                                    .Select(tg => tg.GetElementFromWebElement(By.XPath(".//Text")).Text).ToList();

        //    List<string> headers = gridTableElement.GetDataTableHeaders(DataGridType.ToString()).ToList();
        //    ReadOnlyCollection<IWebElement> rows = gridTableElement.GetDataTableRowElements(DataGridType.ToString());

        //    //int rowIdx = 0;
        //    foreach (var row in rows)
        //    {
        //        var column = row.GetChildElements();
        //        Dictionary<string, IWebElement> kvps = new Dictionary<string, IWebElement>();
        //        string header;
        //        IWebElement cell;
                
        //        foreach (var headerAndCol in headers.Zip(column, Tuple.Create))
        //        {
        //            header = headerAndCol.Item1;
        //            cell = headerAndCol.Item2;
        //            int rowIdx = int.Parse(cell.GetAttribute("GridItem.Row")); // Get GridItem.Row index
        //            GetDataTableElement(DataGridType.ToString()).Add((rowIdx, header), cell);
        //        }
        //        //rowIdx++;
        //    }
        //}

        //public void SaveGridCell(string DataGridAutomationID, int rowIdx, string colName)
        //{
        //    List<string> headers = CurrentDriver.GetDataTableHeaders(DataGridAutomationID).ToList();
        //    ReadOnlyCollection<IWebElement> rows = CurrentDriver.GetDataTableRowElements(DataGridAutomationID);

        //    IWebElement cell = rows[rowIdx].GetChildElements()[headers.IndexOf(colName)];
        //    GetDataTableElement(DataGridAutomationID).Add((rowIdx, colName), cell);
        //}

        public IElement GetDataTableElement(string DataGridAutomationID)
        {
            if (dataTablesCache.Get(DataGridAutomationID) == null)
                dataTablesCache.Add(DataGridAutomationID, PP5IDEWindow.GetExtendedElement(PP5By.Id(DataGridAutomationID)));

            return dataTablesCache.Get(DataGridAutomationID);

            /* Legacy method
            //switch (DataGridAutomationID)
            //{
            //    case "PGGrid":
            //        // Test Item Context
            //        return dataTableTestItemsCache;

            //    case "CndGrid":
            //        // Condition
            //        return dataTableConditionCache;

            //    case "RstGrid":
            //        // Result
            //        return dataTableResultCache;

            //    case "TmpGrid":
            //        // Temp
            //        return dataTableTempCache;

            //    case "GlbGrid":
            //        // Global
            //        return dataTableGlobalCache;

            //    case "DefectCodeGrid":
            //        // Defect Code
            //        return dataTableDefectCodeCache;

            //    case "ParameterGrid":
            //        // Test command Parameter
            //        return dataTableTestCmdParamCache;

            //    case "LoginGrid":
            //        // Open Test Item
            //        return dataTableAllTestItemsCache;

            //    default:
            //        return null;
            //}

            /* if else method
            //if (DataGridAutomationID == "PGGrid")
            //// Test Item Context
            //    return new MemoryCache<(int?, string), IWebElement>(dataTableTestItemsCache);

            ////// Variable
            //else if (DataGridAutomationID == "CndGrid")
            //// Condition
            //    return new MemoryCache<(int?, string), IWebElement>(dataTableConditionCache);

            //else if (DataGridAutomationID == "RstGrid")
            //// Result
            //    return new MemoryCache<(int?, string), IWebElement>(dataTableResultCache);

            //else if (DataGridAutomationID == "TmpGrid")
            //// Temp
            //    return new MemoryCache<(int?, string), IWebElement>(dataTableTempCache);

            //else if (DataGridAutomationID == "GlbGrid")
            //// Global
            //    return new MemoryCache<(int?, string), IWebElement>(dataTableGlobalCache);

            //else if (DataGridAutomationID == "DefectCodeGrid")
            //// Defect Code
            //    return new MemoryCache<(int?, string), IWebElement>(dataTableDefectCodeCache);

            //else if (DataGridAutomationID == "ParameterGrid")
            //// Test command Parameter
            //    return new MemoryCache<(int?, string), IWebElement>(dataTableTestCmdParamCache);

            //else if (DataGridAutomationID == "LoginGrid")
            //    // Test command Parameter
            //    return new MemoryCache<(int?, string), IWebElement>(dataTableAllTestItemsCache);

            //else return null;
            */
        }

        public IElement GetDataTableElement(DataTableAutoIDType DataGridType)
        {
            string DataGridAutomationID = DataGridType.ToString();

            if (dataTablesCache.Get(DataGridAutomationID) == null)
                dataTablesCache.Add(DataGridAutomationID, PP5IDEWindow.GetExtendedElement(PP5By.Id(DataGridAutomationID)));

            return dataTablesCache.Get(DataGridAutomationID);

            /* Legacy method  
            //switch (DataGridType)
            //{
            //    case DataTableAutoIDType.PGGrid:
            //        // Test Item Context
            //        return CurrentDriver.GetElementFromWebElement(MobileBy.AccessibilityId(DataTableAutoIDType.PGGrid.ToString()));

            //    case DataTableAutoIDType.CndGrid:
            //        // Condition
            //        return CurrentDriver.GetElementFromWebElement(MobileBy.AccessibilityId(DataTableAutoIDType.CndGrid.ToString()));

            //    case DataTableAutoIDType.RstGrid:
            //        // Result
            //        return CurrentDriver.GetElementFromWebElement(MobileBy.AccessibilityId(DataTableAutoIDType.RstGrid.ToString()));

            //    case DataTableAutoIDType.TmpGrid:
            //        // Temp
            //        return CurrentDriver.GetElementFromWebElement(MobileBy.AccessibilityId(DataTableAutoIDType.TmpGrid.ToString()));

            //    case DataTableAutoIDType.GlbGrid:
            //        // Global
            //        return CurrentDriver.GetElementFromWebElement(MobileBy.AccessibilityId(DataTableAutoIDType.GlbGrid.ToString()));

            //    case DataTableAutoIDType.DefectCodeGrid:
            //        // Defect Code
            //        return CurrentDriver.GetElementFromWebElement(MobileBy.AccessibilityId(DataTableAutoIDType.DefectCodeGrid.ToString()));

            //    case DataTableAutoIDType.ParameterGrid:
            //        // Test command Parameter
            //        return CurrentDriver.GetElementFromWebElement(MobileBy.AccessibilityId(DataTableAutoIDType.ParameterGrid.ToString()));

            //    case DataTableAutoIDType.LoginGrid:
            //        // Open Test Item
            //        return CurrentDriver.GetElementFromWebElement(MobileBy.AccessibilityId(DataTableAutoIDType.LoginGrid.ToString()));

            //    default:
            //        return null;
            //}
            */
        }

        internal static bool CheckLoginPageIsOpened()
        {
            //AutoUIExecutor.SwitchTo(SessionType.PP5IDE);
            //WindowsElement PP5_TIEditorWindow= CurrentDriver.FindElement(By.Name(PowerPro5Config.IDE_TIEditorWindowName));
            //return PP5_TIEditorWindow != null;
            string WindowTitle = PowerPro5Config.LoginWindowName;
            return CheckWindowOpened(WindowTitle);
        }

        internal static bool CheckMainPanelIsOpened()
        {
            //AutoUIExecutor.SwitchTo(SessionType.MainPanel);
            //WindowsElement PP5_MainPanel = CurrentDriver.FindElement(By.Name(PowerPro5Config.MainPanelWindowName));
            //return PP5_MainPanel != null;
            string WindowTitle = PowerPro5Config.MainPanelWindowName;
            return CheckWindowOpened(WindowTitle);
        }

        internal static bool CheckTIEditorWindowIsOpened()
        {
            //AutoUIExecutor.SwitchTo(SessionType.PP5IDE);
            //WindowsElement PP5_TIEditorWindow= CurrentDriver.FindElement(By.Name(PowerPro5Config.IDE_TIEditorWindowName));
            //return PP5_TIEditorWindow != null;
            string WindowTitle = PowerPro5Config.IDE_TIEditorWindowName;
            return CheckWindowOpened(WindowTitle);
        }

        internal static bool CheckGUIEditorWindowIsOpenedByName(string GUIFileName = "Demo")
        {
            string GUIWindowTitle = PowerPro5Config.IDE_GUIEditorWindowName;
            if (GUIFileName != "Demo")
                GUIWindowTitle = PowerPro5Config.IDE_GUIEditorWindowName.Replace("Demo", GUIFileName);

            return CheckWindowOpened(GUIWindowTitle);
        }

        internal static bool CheckGUIEditorWindowIsOpened()
        {
            IWebElement currWindow = PP5IDEWindow;
            return Regex.IsMatch(currWindow.Text, @"Chroma ATS IDE - \[GUI Editor - .+\]");
        }

        internal static bool CheckPP5WindowIsOpened(WindowType windowType)
        {
            if (windowType == WindowType.MainPanelMenu)
                return CurrentDriver.PerformGetElement($"/ByName[{PowerPro5Config.MainPanelWindowName}]") != null;

            IElement currWindow = PP5IDEWindow;
            if (currWindow == null) return false;

            if (windowType == WindowType.GUIEditor || windowType == WindowType.OnlineControl || windowType == WindowType.Report)
            {
                return Regex.IsMatch(currWindow.Text, @$"Chroma ATS IDE - \[{windowType.GetDescription()} - .+\]");
            }
            else
            {
                return Regex.IsMatch(currWindow.Text, @$"Chroma ATS IDE - \[{windowType.GetDescription()}\]");
            }
        }

        internal static bool CheckPP5WindowIsOpenedNew(WindowType windowType)
        {
            if (windowType == WindowType.MainPanelMenu)
                return CurrentDriver.PerformGetElement($"/ByName[{PowerPro5Config.MainPanelWindowName}]") != null;

            IElement currWindow = PP5IDEWindow;
            if (currWindow == null) 
                return false;

            return windowType.GetDescription() == PP5IDEWindow.ModuleName;
        }

        internal static bool IsPP5IDEWindow(string windowName)
        {
            return Regex.IsMatch(windowName, @$"Chroma ATS IDE - \[.+\]");
        }

        internal static IElement GetPP5Window()
        {
            var currWindow = CurrentDriver.SwitchToWindow(out bool _);

            if (currWindow == null || !currWindow.Title.Contains("Chroma ATS")) return null;

            //return currWindow.GetElementFromWebElement(By.Name(currWindow.Title));
            return currWindow.GetExtendedElement(PP5By.Name(currWindow.Title));
            
            //if (windowType == WindowType.GUIEditor || windowType == WindowType.OnlineControl || windowType == WindowType.Report)
            //{
            //    return Regex.IsMatch(currWindow.Title, @$"Chroma ATS IDE - \[{windowType.GetDescription()} - .+\]");
            //}
            //else
            //{
            //    return Regex.IsMatch(currWindow.Title, @$"Chroma ATS IDE - \[{windowType.GetDescription()}\]");
            //}
        }

        static bool GetCommandSourceType(string groupName, out bool IsDevice)
        {
            IsDevice = false;
            //if (groupName == "Thread TI" || groupName == "Sub TI")
            //{
            //    IsDevice = false;
            //    return true;
            //}

            if (!cmdGroupDataDict.AsDictionary().TryGetValue(groupName, out CommandGroupData cgdata))
            {
                //Console.WriteLine($"GroupName with name '{groupName}' not found.");
                return false;
            }
            IsDevice = cgdata.IsDevice;
            return true;
        }

        internal void PP5IDEWindowRefresh()
        {
            AutoUIExecutor.SwitchTo(SessionType.PP5IDE);
            _PP5IDEWindow = GetPP5Window();
        }

        public static string GetCommandFileFullPath()
        {
           return string.Format(PowerPro5Config.SubPathPattern, PowerPro5Config.ReleaseDataFolder, PowerPro5Config.SystemCommandFileName);
        }


        public void SetColor(IElement colorTabItem, ColorSettingType colorSettingType, string colorCode = "default" /*default color: transparent White (#00FFFFFF)*/)
        {
            var FontColorEditBtn = colorTabItem.GetCustomElement(colorSettingType.GetDescription(), (e) => e.Enabled);
            FontColorEditBtn.LeftClick();

            SetColor(colorCode);
        }

        public void SetColor(IElement colorTabItem, ColorSettingType colorSettingType, Colors colorType /*default color: transparent White (#00FFFFFF)*/)
        {
            var FontColorEditBtn = colorTabItem.GetCustomElement(colorSettingType.GetDescription(), (e) => e.Enabled);
            FontColorEditBtn.LeftClick();

            string colorCode = colorType.GetDescription();
            SetColor(colorCode);
        }

        private void SetColor(string colorCode)
        {
            //string[] colorNames = Enum.GetNames(typeof(Colors));
            if (!colorCode.Contains('#') && colorCode == "default")
            {
                PP5IDEWindow.GetExtendedElement(PP5By.Id("DefaultColor"))
                            .GetFirstListBoxItemElement()
                            .LeftClick();
            }
            else
            {
                List<string> colorCodeList = TypeExtension.GetEnumDescriptions<Colors>();
                Logger.LogMessage($"colorCode:{colorCode}");
                int colorIndex = colorCodeList.IndexOf(colorCode);
                Logger.LogMessage($"colorIndex:{colorIndex}");

                //PP5IDEWindow.GetElementFromWebElement(MobileBy.AccessibilityId("DefaultPicker"))
                //            .SelectComboBoxItemByIndex(colorIndex, supportKeyInputSearch: false);

                PP5IDEWindow.GetExtendedElement(PP5By.Id("DefaultPicker"))
                            .ComboBoxSelectByIndex(colorIndex);
            }
        }

        public static void WaitUntilTIFinishedSaving()
        {
            if (CheckPP5WindowIsOpenedNew(WindowType.TIEditor))
                WaitUntil(() => PP5IDEWindow.PerformGetElement("/ByCondition#ToolBar[Visible]/RadioButton[4]").Enabled);
        }

        public bool AddNewEnumItem(string name, string value)
        {
            List<string> enumNames = PP5IDEWindow.PerformGetElement("/ByName[Enum Item Editor]/ById[EnumItemGrid]").GetSingleColumnValues(1/*Name*/);
            PP5IDEWindow.PerformClick("/ByName[Enum Item Editor,New]", ClickType.LeftClick);

            PP5IDEWindow.PerformInput("/ByName[Enum Item Creater]/Edit[0]", InputType.SendContent, name);
            if (enumNames.Contains(name))
            {
                Logger.LogMessage($"{PP5IDEWindow.PerformGetElement("/ByName[Enum Item Creater]/Edit[0]").GetToolTipMessage()}");
                return false;
            }

            PP5IDEWindow.PerformInput("/ByName[Enum Item Creater]/Edit[1]", InputType.SendContent, value);
            PP5IDEWindow.PerformClick("/ByName[Enum Item Creater,Ok]", ClickType.LeftClick);
            var nameCellAdded = PP5IDEWindow.PerformGetElement($"/ByName[Enum Item Editor]/ById[EnumItemGrid]/ByCell[1,#{name}]");
            if (nameCellAdded != null)
                return true;
            else 
                return false;
        }

        public bool AddNewEnumItem(PP5DataGrid dataGrid, int rowNo, string name, string value)
        {
            dataGrid.GetCellBy(rowNo, VariableColumnType.EnumItem.GetDescription()).LeftClick();
            bool AddNewEnumItemSuccess = AddNewEnumItem(name, value);
            PP5IDEWindow.PerformClick("/ByName[Enum Item Editor,Ok]", ClickType.LeftClick);
            return AddNewEnumItemSuccess;
        }

        public void AddNewEnumItems(PP5DataGrid dataGrid, int rowNo, string[] names, string[] values)
        {
            dataGrid.GetCellBy(rowNo, VariableColumnType.EnumItem.GetDescription()).LeftClick();

            for (int i = 0; i < names.Length; i++) 
            {
                AddNewEnumItem(names[i], values[i]);
            }

            PP5IDEWindow.PerformClick("/ByName[Enum Item Editor,Ok]", ClickType.LeftClick);
        }

        public void AddDefectCode(int defectCode, string desc, int customerDefectCode)
        {
            // 1. Open management window, switch to MISC > Defect Code
            // 2. Add 1 defect code
            //   2-1. First check the defectCode/customerDefectCode to be added is not in the existing defect code table
            //   2-2. Click button "Add", "Add Defect Code" window show up
            //   2-3. In "Add Defect Code" window, input defectCode, desc, customerDefectCode respectively
            //   2-4. Click button "Ok" to confirm to add the defect code
            // 3. Switch back to original module's window

            string orgModule = PP5IDEWindow.ModuleName;
            SwitchToModule(WindowType.Management);

            PP5IDEWindow.PerformClick("/ByCondition#ToolBar[Visible]/RadioButton[3]", ClickType.LeftClick); // Click on MISC button
            IElement mainTab = PP5IDEWindow.PerformGetElement("/Tab[mainTab]");                             
            IElement DefectCodeTabItem = mainTab.TabSelect(3, "Defect Code");                               // Click on Defect Code tab

            var dfs = DefectCodeTabItem.PerformGetElement("/ById[DefectCodeDataGrid]").GetSingleColumnValues(1/*Defect Code*/);
            var cdfs = DefectCodeTabItem.PerformGetElement("/ById[DefectCodeDataGrid]").GetSingleColumnValues(3/*Customer Defect Code*/);
            if (dfs.Contains(defectCode.ToString()) || cdfs.Contains(customerDefectCode.ToString()))
            {
                MenuSelect("Windows", orgModule); return;
            }   

            DefectCodeTabItem.PerformClick("/ByName[Ok]", ClickType.LeftClick);

            AutoUIExecutor.SwitchTo(SessionType.Desktop);
            PP5IDEWindow.PerformInput("/ByName[Add Defect Code]/ById[ErrorCode]", InputType.SendContent, defectCode.ToString());
            PP5IDEWindow.PerformInput("/ByName[Add Defect Code]/ById[Description]", InputType.SendContent, desc);
            PP5IDEWindow.PerformInput("/ByName[Add Defect Code]/ById[CustomerDefectCode]", InputType.SendContent, customerDefectCode.ToString());
            PP5IDEWindow.PerformClick("/ByName[Add Defect Code]/ById[OKButton]", ClickType.LeftClick);
            AutoUIExecutor.SwitchTo(SessionType.PP5IDE);

            WaitUntil(() => PP5IDEWindow.PerformGetElement($"/ByCell@DefectCodeDataGrid[1,#{defectCode}]").GetText() == defectCode.ToString());
            MenuSelect("Windows", orgModule);
        }

        public bool SwitchToModule(WindowType windowType)
        {
            WindowType NotSupportedWindowType = WindowType.None | WindowType.TIEditor | WindowType.TPEditor | WindowType.Execution | WindowType.MainPanelMenu | WindowType.OnlineControl;
            if (windowType == NotSupportedWindowType)
                return false;

            if (MenuSelect("Functions", windowType.GetDescription()))
                WaitUntil(() => PP5IDEWindow.ModuleName == windowType.GetDescription(), 10000);
            else
                MenuSelect("Windows", windowType.GetDescription());

            return true;
        }

        /// <summary>
        /// Get all scrollbar states in a test item window
        /// 0: get horizontal, 1: get Vertical
        /// </summary>
        /// <param name="horizontalOrVertical"></param>
        /// <returns></returns>
        public Dictionary<string, bool> TIWindowGetAllScrollBarEnabledStates(int horizontalOrVertical)
        {
            Dictionary<string, bool> canScrollDict = new Dictionary<string, bool>();

            if (horizontalOrVertical == 0)
            {
                canScrollDict.Add("TIContext", PP5IDEWindow.PerformGetElement("/ByClass[PGGridAeraView]/ById[PGGrid,HeaderPanel]").CanScrollHorizontally);              // Test Item datagrid

                // Test Command Parameter Page
                canScrollDict.Add("TCParamEdit", PP5IDEWindow.PerformGetElement("/ById[editParamAreaView,ParameterGrid,HeaderPanel]").CanScrollHorizontally);           // Parameter datagrid
                canScrollDict.Add("TCParamDesc", PP5IDEWindow.PerformGetElement("/ById[TCParameterPanel,editParamAreaView]/Pane[0]").CanScrollHorizontally);            // Parameter Type
                canScrollDict.Add("ParamType", PP5IDEWindow.PerformGetElement("/ById[TCParameterPanel,editParamAreaView]/Pane[1]").CanScrollHorizontally);              // Parameter Type
                canScrollDict.Add("ConstType", PP5IDEWindow.PerformGetElement("/ById[TCParameterPanel,editParamAreaView]/Pane[2]").CanScrollHorizontally);              // Constant Type

                // Variable Page
                canScrollDict.Add("Variable", PP5IDEWindow.PerformGetElement("/ByClass[CndGridView]/ById[CndGrid,HeaderPanel]").CanScrollHorizontally);                 // Variable datagrid

                // Test Command Page
                canScrollDict.Add("TCSrchBx", PP5IDEWindow.PerformGetElement("/ById[TCListPanel]/ByClass[CmdTreeView,ScrollViewer]").CanScrollHorizontally);            // Test Command SearchBox
                canScrollDict.Add("TCDev", PP5IDEWindow.PerformGetElement("/ById[TCListPanel]/ByClass[CmdTreeView]/ById[DeivceCmdTree]").CanScrollHorizontally);        // Device command treeview
                canScrollDict.Add("TCSys", PP5IDEWindow.PerformGetElement("/ById[TCListPanel]/ByClass[CmdTreeView]/ById[SystemCmdTree]").CanScrollHorizontally);        // System command treeview
            }
            else if (horizontalOrVertical == 1)
            {
                canScrollDict.Add("TIContext", PP5IDEWindow.PerformGetElement("/ByClass[PGGridAeraView]/ById[PGGrid,HeaderPanel]").CanScrollVertically);             // Test Item datagrid

                // Test Command Parameter Page
                canScrollDict.Add("TCParam", PP5IDEWindow.PerformGetElement("/ById[editParamAreaView,ParameterGrid,HeaderPanel]").CanScrollVertically);              // Parameter datagrid
                canScrollDict.Add("ParamType", PP5IDEWindow.PerformGetElement("/ById[TCParameterPanel,editParamAreaView]/Pane[0]").CanScrollVertically);             // Parameter Type
                canScrollDict.Add("ConstType", PP5IDEWindow.PerformGetElement("/ById[TCParameterPanel,editParamAreaView]/Pane[1]").CanScrollVertically);             // Constant Type

                // Variable Page
                canScrollDict.Add("TIContext", PP5IDEWindow.PerformGetElement("/ByClass[CndGridView]/ById[CndGrid,HeaderPanel]").CanScrollVertically);               // Variable datagrid

                // Test Command Page
                canScrollDict.Add("TCSrchBx", PP5IDEWindow.PerformGetElement("/ById[TCListPanel]/ByClass[CmdTreeView,ScrollViewer]").CanScrollVertically);           // Test Command SearchBox
                canScrollDict.Add("TCDev", PP5IDEWindow.PerformGetElement("/ById[TCListPanel]/ByClass[CmdTreeView]/ById[DeivceCmdTree]").CanScrollVertically);       // Device command treeview
                canScrollDict.Add("TCSys", PP5IDEWindow.PerformGetElement("/ById[TCListPanel]/ByClass[CmdTreeView]/ById[SystemCmdTree]").CanScrollVertically);       // System command treeview
            }

            return canScrollDict;
        }

        public void RestoreDefaultDocking(WindowType windowType)
        {
            string mainMenuItemName;
            if (windowType == WindowType.HardwareConfig)
                mainMenuItemName = "File";
            else if (windowType == WindowType.Execution)
                mainMenuItemName = "Settings";
            else if (windowType == WindowType.TIEditor || windowType == WindowType.TPEditor)
                mainMenuItemName = "Edit";
            else
                throw new ArgumentException($"Not supported module: \"{windowType}\"");

            string moduleName = windowType.GetDescription();
            MenuSelect(mainMenuItemName, "Default Dock");
            PP5IDEWindow.PerformClick($"/ByName[{moduleName},Yes]", ClickType.LeftClick);
        }

        //// For TP Editor, later to do
        //private bool getTISourceType(string groupName, out bool IsDevice)
        //{
        //    if (!cmdGroupSourceTypeDict.TryGetValue(groupName, out IsDevice))
        //    {
        //        Console.WriteLine($"GroupName with name '{groupName}' not found.");
        //        return false;
        //    }
        //    return true;
        //}

        //private static void SetTimer(string procName, int updateMilliseconds)
        //{
        //    // Create a timer with a 1 minute interval.
        //    _timer = new System.Timers.Timer(updateMilliseconds); // 60000 milliseconds = 1 minute
        //    _timer.Elapsed += (sender, e) => OnTimedEvent(sender, e, procName);
        //    _timer.AutoReset = true;
        //    _timer.Enabled = true;
        //}

        //private static void OnTimedEvent(Object source, ElapsedEventArgs e, string procName)
        //{
        //    Console.WriteLine("The Elapsed event was raised at {0:HH:mm:ss.fff}", e.SignalTime);
        //    // Call your method here
        //    CaptureApplicationScreenshot(procName);
        //}

        //internal bool CheckTreeViewIsCollapsed(IWebElement TreeItem)
        //{
        //    return TreeItem.GetAttribute("ExpandCollapse.ExpandCollapseState") == ExpandCollapseState.Collapsed.ToString();
        //}
    }
}
