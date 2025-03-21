﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Automation;
using System.Windows.Threading;
using Chroma.UnitTest.Common;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Tab;
using static PP5AutoUITests.AutoUIActionHelper;
using static PP5AutoUITests.AutoUIExtension;
using static PP5AutoUITests.ControlElementExtension;
using static PP5AutoUITests.ThreadHelper;

namespace PP5AutoUITests
{
    [TestClass]
    public class TestCases_Management : TestBase
    {
        string saveFolder = Path.Combine(Directory.GetCurrentDirectory(), "TestCases_Management");
        double updateInterval = 60000;

        [TestInitialize]
        public void Management_TestMethodSetup()
        {
            OpenDefaultManagementWindow();

            _scrnShotTimer = new CaptureAppScreenshotTimer(CaptureApplicationScreenshot, "dotMemory.UI.64", saveFolder, updateInterval); // 60000 milliseconds = 1 minute
            _scrnShotTimer.Start();
        }

        [TestCleanup]
        public void Management_TestMethodCleanup()
        {
            _scrnShotTimer.Stop();
        }

        /// <summary>
        /// EmptyTest
        /// </summary>
        [TestMethod("Test")]
        public void EmptyTest()
        {
            Assert.IsTrue(true);
        }

        /// <summary>
        /// TestCase: E1-9
        /// In User Account, add User for 5000 times until messagebox popup (user accounts max limit is 5000)
        /// </summary>
        [TestMethod("In User Account, add User for 5000 times until messagebox popup (user accounts max limit is 5000)")]
        [TestCategory("帳號管理(E1)")]
        public void Management_AddUserFor5000Times_UntilWarningBoxPopup()
        {
            // Click on Security button
            CurrentDriver.GetWebElementFromWebDriver(By.ClassName("ToolBar"))
                         .GetWebElementsFromWebElement(By.ClassName("RadioButton"))[0].LeftClick();

            // Click on User ID tab
            IWebElement SecurityTabItem = CurrentDriver.GetWebElementFromWebDriver(MobileBy.AccessibilityId("mainTab"))
                                                       .GetWebElementsFromWebElement(By.ClassName("TabItem"))[0];

            IWebElement UserIDTabItem = SecurityTabItem.GetWebElementFromWebElement(By.ClassName("TabControl"))
                                                       .GetWebElementsFromWebElement(By.ClassName("TabItem"))[0];

            UserIDTabItem.LeftClick();

            bool AddSuccess = true;
            //for (int i = 20000; i < 2000; i++)
            int i = 20000;
            while (AddSuccess)
            {
                CurrentDriver.GetWebElementFromWebDriver(By.Name("Add")).LeftClick();
                AddSuccess = AddNewUser(userID: i.ToString(), password: i.ToString());
                i++;
            }

            // Check "can't add more than 5000 user IDs" warning message box popuped
            //CurrentDriver.GetElementFromWebElement(By.Name("Add")).LeftClick();
            Assert.AreEqual("The total number of User ID cannot exceed 5000, so it cannot be added!", CurrentDriver.GetExtendedElement(PP5By.Name("Chroma ATS IDE - [Management]"))
                                                                                                                   .GetSpecificChildOfControlType(ElementControlType.Window, "Notice")
                                                                                                                   .GetFirstEditContent());

            // Close the warning message box
            CurrentDriver.GetExtendedElement(PP5By.Name("Chroma ATS IDE - [Management]"))
                         .GetSpecificChildOfControlType(ElementControlType.Window, "Notice")
                         .GetWebElementFromWebElement(By.Name("OK"))
                         .LeftClick();
        }

        /// <summary>
        /// TestCase: E2-3
        /// In Role Page, add Role for 5000 times until messagebox popup (role max limit is 5000)
        /// </summary>
        [TestMethod("In Role Page, add Role for 5000 times until messagebox popup (role max limit is 5000)")]
        [TestCategory("角色權限管理(E2)")]
        public void Management_AddRoleFor5000Times_UntilWarningBoxPopup()
        {
            // Click on Security button
            CurrentDriver.GetWebElementFromWebDriver(By.ClassName("ToolBar"))
                         .GetWebElementsFromWebElement(By.ClassName("RadioButton"))[0].LeftClick();

            // Click on Role tab
            IWebElement SecurityTabItem = CurrentDriver.GetWebElementFromWebDriver(MobileBy.AccessibilityId("mainTab"))
                                                       .GetWebElementsFromWebElement(By.ClassName("TabItem"))[0];

            IWebElement roleTabItem = SecurityTabItem.GetWebElementFromWebElement(By.ClassName("TabControl"))
                                                     .GetWebElementsFromWebElement(By.ClassName("TabItem"))[1];

            roleTabItem.LeftClick();

            bool AddSuccess = true;
            //for (int i = 0; i < 5000; i++)
            int i = 20000;
            while (AddSuccess)
            {
                CurrentDriver.GetWebElementFromWebDriver(By.Name("Add")).LeftClick();
                AddSuccess = AddNewRole(roleName: i.ToString());
                i++;
            }
            //"/Pane[@ClassName=\"#32769\"][@Name=\"桌面 1\"]/Window[@ClassName=\"Window\"][@Name=\"Chroma ATS IDE - [Management]\"]/Window[@ClassName=\"Window\"][@Name=\"Notice\"]/Edit[@ClassName=\"TextBox\"]"
            // Check "can't add more than 5000 roles" warning message box popuped
            //CurrentDriver.GetElementFromWebElement(By.Name("Add")).LeftClick();
            Assert.AreEqual("The total number of Role cannot exceed 5000, so it cannot be added!", CurrentDriver.GetExtendedElement(PP5By.Name("Chroma ATS IDE - [Management]"))
                                                                                                                .GetSpecificChildOfControlType(ElementControlType.Window, "Notice")
                                                                                                                .GetFirstEditContent());

            // Close the warning message box
            CurrentDriver.GetExtendedElement(PP5By.Name("Chroma ATS IDE - [Management]"))
                         .GetSpecificChildOfControlType(ElementControlType.Window, "Notice")
                         .GetWebElementFromWebElement(By.Name("OK"))
                         .LeftClick();
        }

        ///// <summary>
        ///// TestCase: E3-10
        ///// In Activity Log Page, only show latest 100000 data
        ///// </summary>
        //[TestMethod("In Activity Log Page, only show latest 100000 data")]
        //public void Management_AddActivityLogByEditUserAccountActiveState100000Times_CheckLatest100000Data()
        //{
        //    // Click on Security button
        //    CurrentDriver.GetElementFromWebElement(By.ClassName("ToolBar"))
        //                 .GetElements(By.ClassName("RadioButton"))[0].LeftClick();

        //    // Click on User ID tab
        //    IWebElement SecurityTabItem = CurrentDriver.GetElementFromWebElement(MobileBy.AccessibilityId("mainTab"))
        //                                               .GetElements(By.ClassName("TabItem"))[0];

        //    IWebElement UserIDTabItem = SecurityTabItem.GetElementFromWebElement(By.ClassName("TabControl"))
        //                                               .GetElements(By.ClassName("TabItem"))[0];

        //    UserIDTabItem.LeftClick();

        //    //bool AddSuccess = true;
        //    bool isActive = false;
        //    //while (AddSuccess)
        //    for (int i = 0; i < 100000; i++)
        //    {
        //        CurrentDriver.GetElementFromWebElement(By.Name("Edit")).LeftClick();
        //        EditUser(isActive);
        //        System.Threading.Thread.Sleep(500);
        //        isActive = !isActive;
        //    }

        //    // Click on Activity Log tab
        //    IWebElement activityLogTabItem = SecurityTabItem.GetElementFromWebElement(By.ClassName("TabControl"))
        //                                                    .GetElements(By.ClassName("TabItem"))[2];

        //    activityLogTabItem.LeftClick();

        //    // Check only latest 100000 data in the datagrid
        //    Assert.AreEqual(100000, activityLogTabItem.GetFirstDataGridElement()
        //                                              .GetElements(By.TagName(ElementControlType.DataItem.GetDescription())));
        //}

        /// <summary>
        /// TestCase: E3-10
        /// In Activity Log Page, only show latest 100000 data
        /// </summary>
        [TestMethod("In Activity Log Page, only show latest 100000 data")]
        [TestCategory("使用者操作記錄(E3)")]
        public void Management_AddActivityLogByEditUserAccountActiveState100000Times_CheckLatest100000Data()
        {
            // Click on TP/TI button
            CurrentDriver.GetWebElementFromWebDriver(By.ClassName("ToolBar"))
                         .GetWebElementsFromWebElement(By.ClassName("RadioButton"))[1].LeftClick();

            // Click on Test Item tab
            IElement TPTITabItem = CurrentDriver.GetExtendedElement(PP5By.Id("mainTab"))
                                                   .GetExtendedElements(PP5By.ClassName("TabItem"))[1];

            IElement TestItemTabItem = TPTITabItem.GetExtendedElement(PP5By.ClassName("TabControl"))
                                                     .GetExtendedElements(PP5By.ClassName("TabItem"))[1];

            TestItemTabItem.LeftClick();

            IElement tiDataGrid = TestItemTabItem.GetFirstDataGridElement();

            bool remarkBool = false;
            for (int i = 0; i < 100000; i++)
            {
                tiDataGrid.GetCellBy(1, 5).SendText(remarkBool.ToString());
                Press(Keys.Enter);
                System.Threading.Thread.Sleep(1200);
                remarkBool = !remarkBool;
            }

            // Click on Activity Log tab
            CurrentDriver.GetWebElementFromWebDriver(By.ClassName("ToolBar"))
                         .GetWebElementsFromWebElement(By.ClassName("RadioButton"))[0].LeftClick();

            // Click on Security tab
            IElement SecurityTabItem = CurrentDriver.GetExtendedElement(PP5By.Id("mainTab"))
                                                    .GetExtendedElements(PP5By.ClassName("TabItem"))[0];

            IElement activityLogTabItem = SecurityTabItem.GetExtendedElement(PP5By.ClassName("TabControl"))
                                                         .GetExtendedElements(PP5By.ClassName("TabItem"))[2];

            activityLogTabItem.LeftClick();

            // Check only latest 100000 data in the datagrid
            Assert.AreEqual(100000, activityLogTabItem.GetFirstDataGridElement()
                                                      .GetWebElementsFromWebElement(By.TagName(ElementControlType.DataItem.GetDescription())));
        }

        /// <summary>
        /// TestCase: E13-17
        /// In MISC > Defect Code Page, add defect code for 1000 times until messagebox popup (max limit is 1000)
        /// 總數量超過1000個，彈出無法新增之訊息
        /// </summary>
        [TestMethod("In MISC > Defect Code Page, add Defect Code for 1000 times until messagebox popup (max limit is 1000)")]
        [TestCategory("不良代碼管理(E13)")]
        public void Management_AddDefectCodeFor1000Times_UntilWarningBoxPopup()
        {
            // Click on MISC button
            CurrentDriver.GetWebElementFromWebDriver(By.ClassName("ToolBar"))
                         .GetWebElementsFromWebElement(By.ClassName("RadioButton"))[3].LeftClick();

            // Click on Defect Code tab
            IWebElement MISCTabItem = CurrentDriver.GetWebElementFromWebDriver(MobileBy.AccessibilityId("mainTab"))
                                                   .GetWebElementsFromWebElement(By.ClassName("TabItem"))[3];

            IWebElement dfTabItem = MISCTabItem.GetWebElementFromWebElement(By.ClassName("TabControl"))
                                               .GetWebElementsFromWebElement(By.ClassName("TabItem"))[6];

            dfTabItem.LeftClick();

            bool AddSuccess = true;

            int i = 200;
            int j = i;
            while (AddSuccess)
            {
                CurrentDriver.GetWebElementFromWebDriver(By.Name("Add")).LeftClick();
                AddSuccess = AddDefectCode(defectCode: i, customerDefectCode: i);
                i++;
                if (j - i >= 1000)
                    break;
            }

            Assert.AreEqual("The total number of Defect Code cannot exceed 1000, so it cannot be added!", CurrentDriver.GetExtendedElement(PP5By.Id("Chroma ATS IDE - [Management]"))
                                                                                                                        .GetSpecificChildOfControlType(ElementControlType.Window, "Notice")
                                                                                                                        .GetFirstEditContent());

            // Close the warning message box
            CurrentDriver.GetExtendedElement(PP5By.Id("Chroma ATS IDE - [Management]"))
                         .GetSpecificChildOfControlType(ElementControlType.Window, "Notice")
                         .GetWebElementFromWebElement(By.Name("OK"))
                         .LeftClick();
        }

        /// <summary>
        /// TestCase: J1-12
        /// In User Account, add and delete User for 300 times
        /// </summary>
        [TestMethod("In User Account, add and delete User for 300 times")]
        [TestCategory("長時間動作(J1)")]
        [DataRow(300)] // (Estimated testing time: 150 mins)
        public void Management_AddAndDeleteUser_For300Times(int repeatCount)
        {
            // Click on Security button
            CurrentDriver.GetWebElementFromWebDriver(By.ClassName("ToolBar"))
                         .GetWebElementsFromWebElement(By.ClassName("RadioButton"))[0].LeftClick();

            // Click on User ID tab
            IWebElement SecurityTabItem = CurrentDriver.GetWebElementFromWebDriver(MobileBy.AccessibilityId("mainTab"))
                                                       .GetWebElementsFromWebElement(By.ClassName("TabItem"))[0];

            IWebElement UserIDTabItem = SecurityTabItem.GetWebElementFromWebElement(By.ClassName("TabControl"))
                                                       .GetWebElementsFromWebElement(By.ClassName("TabItem"))[0];

            UserIDTabItem.LeftClick();

            bool AddSuccess = true;
            for (int i = 0; i < repeatCount; i++)
            {
                GetPP5Window().GetWebElementFromWebElement(By.Name("Add")).LeftClick();
                AddSuccess = AddNewUser(userID: "0", password: "0", valid: false);
                if (AddSuccess)
                {
                    WaitUntil(() => GetPP5Window().GetWebElementFromWebElement(By.Name("Delete")) != null);
                    GetPP5Window().GetWebElementFromWebElement(By.Name("Delete")).LeftClick();
                    GetPP5Window().GetElementWithRetry<IElement, IElement>(5000, By.Name("Notice"), By.Name("Yes"))
                                  .LeftClick();

                    // Close error message box if it shows up
                    if (WaitUntil(() => GetPP5Window().GetWebElementFromWebElement(By.Name("Error")) != null))
                        GetPP5Window().GetElementWithRetry<IElement, IElement>(5000, By.Name("Error"), MobileBy.AccessibilityId("Close"))
                                      .LeftClick();
                }
            }
        }

        /// <summary>
        /// TestCase: J1-13
        /// In role, repeatly modify the authority setting for 1000 times
        /// </summary>
        [TestMethod("In role, repeatly modify the authority setting for 1000 times")]
        [TestCategory("長時間動作(J1)")]
        [DataRow(1000)] // (Estimated testing time: 366 mins)
        public void Management_ModifyUserAuthority_For1000Times(int repeatCount)
        {
            // Click on Security button
            CurrentDriver.GetWebElementFromWebDriver(By.ClassName("ToolBar"))
                         .GetWebElementsFromWebElement(By.ClassName("RadioButton"))[0].LeftClick();

            // Click on role tab
            IWebElement SecurityTabItem = CurrentDriver.GetWebElementFromWebDriver(MobileBy.AccessibilityId("mainTab"))
                                                       .GetWebElementsFromWebElement(By.ClassName("TabItem"))[0];

            IWebElement roleTabItem = SecurityTabItem.GetWebElementFromWebElement(By.ClassName("TabControl"))
                                                     .GetWebElementsFromWebElement(By.ClassName("TabItem"))[1];

            roleTabItem.LeftClick();

            roleTabItem.GetWebElementFromWebElement(By.Name("Add")).LeftClick();
            AddNewRole("99999");

            bool isUserChecked = true;
            for (int i = 0; i < repeatCount; i++)
            {
                WaitUntil(() => GetPP5Window().GetWebElementFromWebElement(By.Name("Edit")) != null);
                GetPP5Window().GetWebElementFromWebElement(By.Name("Edit")).LeftClick();
                EditRole(new Dictionary<string, bool>() { { "User", isUserChecked } });
                isUserChecked = !isUserChecked;
            }

            WaitUntil(() => GetPP5Window().GetWebElementFromWebElement(By.Name("Delete")) != null);
            roleTabItem.GetWebElementFromWebElement(By.Name("Delete")).LeftClick();
            GetPP5Window().GetElementWithRetry<IElement, IElement>(5000, By.Name("Notice"), By.Name("Yes"))
                          .LeftClick();
        }

        /// <summary>
        /// TestCase: J1-14
        /// In activity role, repeatly query the activity log for 1000 times
        /// </summary>
        [TestMethod("In activity role, repeatly query the activity log for 1000 times")]
        [TestCategory("長時間動作(J1)")]
        [DataRow(1000)] // (Estimated testing time: 333 mins)
        public void Management_QueryActivityLog_For1000Times(int repeatCount)
        {
            // Click on Security tab
            CurrentDriver.GetWebElementFromWebDriver(By.ClassName("ToolBar"))
                         .GetWebElementsFromWebElement(By.ClassName("RadioButton"))[0].LeftClick();

            // Click on Activity Log tab
            IElement SecurityTabItem = CurrentDriver.GetExtendedElement(PP5By.Id("mainTab"))
                                                    .GetExtendedElements(PP5By.ClassName("TabItem"))[0];

            IElement activityLogTabItem = SecurityTabItem.GetExtendedElement(PP5By.ClassName("TabControl"))
                                                         .GetExtendedElements(PP5By.ClassName("TabItem"))[2];

            activityLogTabItem.LeftClick();

            IElement cmbFunction = activityLogTabItem.GetFirstCustomElement()
                                                        .GetFirstPaneElement()
                                                        .GetFirstComboBoxElement();

            int functionCountForQuery = cmbFunction.GetComboBoxItemCount();

            for (int i = 0; i < repeatCount; i++)
            {
                // 隨機選取Function中任一選項
                cmbFunction.SelectComboBoxItemByIndex(new Random().Next(0, functionCountForQuery + 1)); // 生成介於0和functionCountForQuery之間的隨機整數

                // Query activity Log
                activityLogTabItem.GetBtnElement("Query").LeftClick();
            }
        }

        /// <summary>
        /// TestCase: J1-15
        /// In MISC > Defect Code Page, add and delete defect code for 1000 times
        /// </summary>
        [TestMethod("In MISC > Defect Code Page, add and delete Defect Code for 1000 times")]
        [TestCategory("長時間動作(J1)")]
        [DataRow(1000)] // (Estimated testing time: 366 mins)
        public void Management_AddAndDeleteDefectCode_For1000Times(int repeatCount)
        {
            // Click on MISC button
            CurrentDriver.GetWebElementFromWebDriver(By.ClassName("ToolBar"))
                         .GetWebElementsFromWebElement(By.ClassName("RadioButton"))[3].LeftClick();

            // Click on Defect Code tab
            IElement MISCTabItem = CurrentDriver.GetExtendedElement(PP5By.Id("mainTab"))
                                                .GetExtendedElements(PP5By.ClassName("TabItem"))[3];

            IElement dfTabItem = MISCTabItem.GetExtendedElement(PP5By.ClassName("TabControl"))
                                            .GetExtendedElements(PP5By.ClassName("TabItem"))[6];

            dfTabItem.LeftClick();

            IElement dgDfCode = dfTabItem.GetFirstCustomElement()
                                            .GetFirstPaneElement()
                                            .GetFirstDataGridElement();

            int dfCode = 77777;
            for (int i = 0; i < repeatCount; i++)
            {
                WaitUntil(() => PP5IDEWindow.GetWebElementFromWebElement(By.Name("Add")) != null);
                PP5IDEWindow.GetWebElementFromWebElement(By.Name("Add")).LeftClick();
                AddDefectCode(defectCode: dfCode, customerDefectCode: dfCode);

                //GetPP5Window().GetCellByName(1, dfCode.ToString()).LeftClick();
                //dgDfCode.SelectDataGridItemByRowIndex(1);

                IElement dfCodeItem = null;
                while (dfCodeItem == null)
                {
                    dfCodeItem = PP5IDEWindow.GetExtendedElement(PP5By.Id("DefectCodeDataGrid"))
                                               .GetRowByName(colNo: 1, cellName: dfCode.ToString());
                }
                dfCodeItem.LeftClick();

                PP5IDEWindow.GetWebElementFromWebElement(By.Name("Delete")).LeftClick();
                PP5IDEWindow.GetElementWithRetry<IElement, IElement>(5000, By.Name("Notice"), By.Name("Yes"))
                              .LeftClick();
            }
        }

        /// <summary>
        /// TestCase: J1-16
        /// In MISC > Defect Code Page, import and delete defect code for 1000 times
        /// </summary>
        [TestMethod("In MISC > Defect Code Page, import and delete Defect Code for 1000 times")]
        [TestCategory("長時間動作(J1)")]
        [DataRow(1000)] // (Estimated testing time: 566 mins)
        public void Management_ImportAndDeleteDefectCode_For1000Times(int repeatCount)
        {
            // Click on MISC button
            CurrentDriver.GetWebElementFromWebDriver(By.ClassName("ToolBar"))
                         .GetWebElementsFromWebElement(By.ClassName("RadioButton"))[3].LeftClick();

            // Click on Defect Code tab
            IElement MISCTabItem = CurrentDriver.GetExtendedElement(PP5By.Id("mainTab"))
                                                .GetExtendedElements(PP5By.ClassName("TabItem"))[3];

            IElement dfTabItem = MISCTabItem.GetExtendedElement(PP5By.ClassName("TabControl"))
                                            .GetExtendedElements(PP5By.ClassName("TabItem"))[6];

            dfTabItem.LeftClick();

            IElement dgDfCode = dfTabItem.GetFirstCustomElement()
                                            .GetFirstPaneElement()
                                            .GetFirstDataGridElement();

            int dfCode = 77777;

            // Add a new df code item
            dfTabItem.GetWebElementFromWebElement(By.Name("Add")).LeftClick();
            AddDefectCode(defectCode: dfCode, customerDefectCode: dfCode);

            // Export the new df code item
            PP5IDEWindow.GetExtendedElement(PP5By.Id("DefectCodeDataGrid"))
                          .SelectDataGridItemByRowIndex(-1);
            PP5IDEWindow.GetWebElementFromWebElement(By.Name("Export")).LeftClick();
            PP5IDEWindow.GetWebElementFromWebElement(By.Name("檔案名稱:")).SendText(dfCode.ToString());
            PP5IDEWindow.GetWebElementFromWebElement(By.Name("存檔(S)")).LeftClick();

            // Delete the new df code item
            WaitUntil(() => PP5IDEWindow.GetWebElementFromWebElement(By.Name("Delete")) != null);
            PP5IDEWindow.GetWebElementFromWebElement(By.Name("Delete")).LeftClick();
            PP5IDEWindow.GetElementWithRetry<IElement, IElement>(5000, By.Name("Notice"), By.Name("Yes")).LeftClick();

            // Import and delete the exported df code item
            for (int i = 0; i < repeatCount; i++)
            {
                PP5IDEWindow.GetWebElementFromWebElement(By.Name("Import")).LeftClick();
                PP5IDEWindow.GetElementWithRetry<IElement, IElement>(5000, By.ClassName("UIItemsView"), By.Name(dfCode + ".dfl"))
                              .LeftClick();
                PP5IDEWindow.GetWebElementFromWebElement(By.Name("開啟(O)")).LeftClick();

                Executor.GetInstance().SwitchTo(SessionType.Desktop);
                WaitUntil(() => CurrentDriver.GetWebElementFromWebDriver(By.Name("Select Import Item")) != null);
                CurrentDriver.GetWebElementFromWebDriver(5000, By.Name("Select Import Item"), By.Name("OK"))
                             .LeftClick();
                Executor.GetInstance().SwitchTo(SessionType.PP5IDE);

                //GetPP5Window().SelectDataGridItemByRowIndex(-1);

                IWebElement dfCodeItem = null;
                while (dfCodeItem == null)
                {
                    dfCodeItem = PP5IDEWindow.GetExtendedElement(PP5By.Id("DefectCodeDataGrid"))
                                               .GetRowByName(colNo: 1, cellName: dfCode.ToString());
                }
                dfCodeItem.LeftClick();

                PP5IDEWindow.GetWebElementFromWebElement(By.Name("Delete")).LeftClick();
                PP5IDEWindow.GetElementWithRetry<IElement, IElement>(5000, By.Name("Notice"), By.Name("Yes"))
                              .LeftClick();
            }
        }

        /// <summary>
        /// TestCase: J1-17
        /// In System Setup > Color Page, Set Color Of Same Test Command for 1000 times
        /// </summary>
        [TestMethod("In System Setup > Color Page, Set Color Of Same Test Command for 1000 times")]
        [TestCategory("長時間動作(J1)")]
        [DataRow(1)] // (Estimated testing time: 280 mins)
        public void Management_SetColorOfSameTestCommand_For1000Times(int repeatCount)
        {
            // Click on System Setup button
            PP5IDEWindow.GetWebElementFromWebElement(By.ClassName("ToolBar"))
                         .GetWebElementsFromWebElement(By.ClassName("RadioButton"))[2].LeftClick();

            // Click on Color tab
            IElement systemSetupTabItem = PP5IDEWindow.GetExtendedElement(PP5By.Id("mainTab"))
                                                      .GetExtendedElements(PP5By.ClassName("TabItem"))[2];

            IElement colorTabItem = systemSetupTabItem.GetExtendedElement(PP5By.ClassName("TabControl"))
                                                      .GetExtendedElements(PP5By.ClassName("TabItem"))[3];

            colorTabItem.LeftClick();

            //IWebElement TestCmdTV = colorTabItem.GetTreeViewElement("ColorTCGroupTreeView");

            // Select the first command of group: "AC Source"
            SelectColorSettingItem(colorTabItem, ColorSettingPageType.TestCommand, TestCmdGroupType.AC_Source, 1, true);

            string colorCodeFont = "#FFFF00FF";
            string colorCodeBg = "#FFFF00FF";

            // Repeat setting the test command's color by alternate between \"#FFFF00FF\" and default color
            for (int i = 0; i < repeatCount; i++)
            {
                //var FontColorEditBtn = colorTabItem.GetCustomElement((e) => e.Text == "Font Color Edit" && e.Enabled);
                //FontColorEditBtn.LeftClick();
                //GetPP5Window().GetElementFromWebElement(MobileBy.AccessibilityId("DefaultPicker"))
                //              .SelectComboBoxItemByName("#FFFF00FF", supportKeyInputSearch: false);

                //FontColorEditBtn.LeftClick();
                //GetPP5Window().GetElementFromWebElement(MobileBy.AccessibilityId("DefaultColor"))
                //              .GetFirstListBoxItemElement()
                //              .LeftClick();

                SetColor(colorTabItem, ColorSettingType.Font, colorCodeFont);
                SetColor(colorTabItem, ColorSettingType.Font, "default");
                SetColor(colorTabItem, ColorSettingType.Background, colorCodeBg);
                SetColor(colorTabItem, ColorSettingType.Background, "default");
            }
        }

        /// <summary>
        /// TestCase: J1-18
        /// In Ex-Function > DLL Page, Add And Delete Same Dll for 1000 times
        /// </summary>
        [TestMethod("In Ex-Function > DLL Page, Add And Delete Same Dll for 1000 times")]
        [TestCategory("長時間動作(J1)")]
        [DataRow(1000)] // (Estimated testing time: 666 mins)
        public void Management_AddAndDeleteSameDll_For1000Times(int repeatCount)
        {
            // Click on Ex-Function button
            CurrentDriver.GetWebElementFromWebDriver(By.ClassName("ToolBar"))
                         .GetWebElementsFromWebElement(By.ClassName("RadioButton"))[4].LeftClick();

            // Click on DLL tab
            IWebElement ExFunctionTabItem = CurrentDriver.GetWebElementFromWebDriver(MobileBy.AccessibilityId("mainTab"))
                                                         .GetWebElementsFromWebElement(By.ClassName("TabItem"))[4];

            IWebElement DllTabItem = ExFunctionTabItem.GetWebElementFromWebElement(By.ClassName("TabControl"))
                                                      .GetWebElementsFromWebElement(By.ClassName("TabItem"))[0];

            DllTabItem.LeftClick();

            string DllFilename = "Managed.dll";
            string DllFolder = @"D:\My docu\新人訓練\專案\Debug\Dll";

            // Repeat add and delete same dll item
            for (int i = 0; i < repeatCount; i++)
            {
                WaitUntil(() => GetPP5Window().GetWebElementFromWebElement(By.Name("Add")) != null);
                GetPP5Window().GetWebElementFromWebElement(By.Name("Add")).LeftClick();

                GetPP5Window().GetBtnElement("先前的位置").LeftClick();
                SendSingleKeys(DllFolder);
                Press(Keys.Enter);
                GetPP5Window().GetWebElementFromWebElement(By.Name(DllFilename)).DoubleClick();

                Executor.GetInstance().SwitchTo(SessionType.Desktop);
                WaitUntil(() => CurrentDriver.GetWebElementFromWebDriver(By.Name("Add DLL - Managed DLL")) != null);
                CurrentDriver.GetExtendedElement(PP5By.Id("TopDataGrid"))
                             .GetCellBy(1, 2).GetFirstCheckBoxElement()
                             .UnTickCheckBox();
                CurrentDriver.GetWebElementFromWebDriver(5000, By.Name("Add DLL - Managed DLL"), By.Name("OK"))
                             .LeftClick();
                Executor.GetInstance().SwitchTo(SessionType.PP5IDE);

                // Find the added Dll and uncheck "Active"
                //WaitUntil(() => DllTabItem.GetElementFromWebElement(By.Name("Add")) != null);
                //System.Threading.Thread.Sleep(6000);

                //GetPP5Window().GetElementFromWebElement(MobileBy.AccessibilityId("DllDataGrid"))
                //          .GetRowByName(colNo: 5, cellName: "Managed.dll")
                //          .LeftClick();

                IWebElement dllItem = null;
                while (dllItem == null)
                {
                    dllItem = PP5IDEWindow.GetExtendedElement(PP5By.Id("DllDataGrid"))
                                            .GetRowByName(colNo: 5, cellName: DllFilename);
                }

                //dllItem.GetCellBy(2).GetFirstCheckBoxElement()
                //       .UnTickCheckBox();

                dllItem.LeftClick();

                PP5IDEWindow.GetWebElementFromWebElement(By.Name("Delete")).LeftClick();
                PP5IDEWindow.GetElementWithRetry<IElement, IElement>(5000, By.Name("Notice"), By.Name("Yes"))
                              .LeftClick();
            }
        }

        /// <summary>
        /// TestCase: J1-19
        /// In Ex-Function > Python Page, Add And Delete Same Python for 300 times
        /// </summary>
        [TestMethod("In Ex-Function > Python Page, Add And Delete Same Python for 300 times")]
        [TestCategory("長時間動作(J1)")]
        [DataRow(300)] // (Estimated testing time: 240 mins)
        public void Management_AddAndDeleteSamePython_For300Times(int repeatCount)
        {
            // Click on Ex-Function button
            CurrentDriver.GetWebElementFromWebDriver(By.ClassName("ToolBar"))
                         .GetWebElementsFromWebElement(By.ClassName("RadioButton"))[4].LeftClick();

            // Click on Python tab
            IWebElement ExFunctionTabItem = CurrentDriver.GetWebElementFromWebDriver(MobileBy.AccessibilityId("mainTab"))
                                                         .GetWebElementsFromWebElement(By.ClassName("TabItem"))[4];

            IWebElement PythonTabItem = ExFunctionTabItem.GetWebElementFromWebElement(By.ClassName("TabControl"))
                                                         .GetWebElementsFromWebElement(By.ClassName("TabItem"))[1];

            PythonTabItem.LeftClick();

            string pythonFilename = "ExternalPyTest_OrderedList.py";
            string pythonFolder = @"D:\My docu\新人訓練\專案\Debug\Python";

            // Repeat add and delete same python item
            for (int i = 0; i < repeatCount; i++)
            {
                WaitUntil(() => PP5IDEWindow.GetWebElementFromWebElement(By.Name("Add")) != null);
                PP5IDEWindow.GetWebElementFromWebElement(By.Name("Add")).LeftClick();

                GetPP5Window().GetBtnElement("先前的位置").LeftClick();
                SendSingleKeys(pythonFolder);
                PP5IDEWindow.GetWebElementFromWebElement(By.Name(pythonFilename)).DoubleClick();

                Executor.GetInstance().SwitchTo(SessionType.Desktop);
                WaitUntil(() => CurrentDriver.GetWebElementFromWebDriver(By.Name("Add Python")) != null);
                var AddPythonDg = CurrentDriver.GetExtendedElement(PP5By.Id("TopDataGrid"));

                // Find the added Python and uncheck "Active"
                AddPythonDg.GetCellBy(rowNo: 1, colNo: 2)
                           .GetFirstCheckBoxElement()
                           .UnTickCheckBox();

                AddPythonDg.GetCellBy(rowNo: 1, colNo: 1)
                           .SendText("test");

                CurrentDriver.GetWebElementFromWebDriver(5000, By.Name("Add Python"), By.Name("OK"))
                    .LeftClick();

                Executor.GetInstance().SwitchTo(SessionType.PP5IDE);

                IWebElement pythonItem = null;
                while (pythonItem == null)
                {
                    pythonItem = PP5IDEWindow.GetExtendedElement(PP5By.Id("PythonDataGrid"))
                                               .GetRowByName(colNo: 5, cellName: pythonFilename);
                }
                pythonItem.LeftClick();

                PP5IDEWindow.GetWebElementFromWebElement(By.Name("Delete")).LeftClick();
                PP5IDEWindow.GetElementWithRetry<IElement, IElement>(5000, By.Name("Notice"), By.Name("Yes"))
                              .LeftClick();
            }
        }

        bool AddNewUser(string userID, string password, string role = "Administrator", string expiredDate = "2099-12-31", bool valid = true, string remark = "")
        {
            Executor.GetInstance().SwitchTo(SessionType.Desktop);
            //try
            //{
            WaitUntil(() => CurrentDriver.GetExtendedElementBySingleWithRetry(PP5By.Name("Add User")) != null);
                var AddUserDialog = CurrentDriver.GetExtendedElement(PP5By.Name("Add User"));
                if (AddUserDialog == null)
                {
                    return false;
                }
                AddUserDialog.GetFirstEditElement().SendText(userID);
                AddUserDialog.GetEditElement(1).SendText(password);
                AddUserDialog.GetFirstComboBoxElement().SelectComboBoxItemByName(role);
                AddUserDialog.GetCustomElement(0)
                             .GetFirstEditElement()
                             .SendText(expiredDate);
                if (valid)
                    AddUserDialog.GetRdoBtnElement("Yes").LeftClick();
                else
                    AddUserDialog.GetRdoBtnElement("No").LeftClick();
                AddUserDialog.GetEditElement(2).SendText(remark);
                AddUserDialog.GetBtnElement("Add").LeftClick();
                Executor.GetInstance().SwitchTo(SessionType.PP5IDE);
                return true;
            //}
            //catch
            //{
            //    if (CurrentDriver.GetElementFromWebElement(MobileBy.AccessibilityId("MessageBoxExDialog")) != null)
            //        return false;
            //    else
            //        throw;
            //}
        }

        bool AddNewRole(string roleName)
        {
            Executor.GetInstance().SwitchTo(SessionType.Desktop);
            //try
            //{
            WaitUntil(() => CurrentDriver.GetExtendedElementBySingleWithRetry(PP5By.Name("Add Role")) != null);
            var AddRoleDialog = CurrentDriver.GetExtendedElementBySingleWithRetry(PP5By.Name("Add Role"));
                if (AddRoleDialog == null)
                {
                    return false;
                }
                AddRoleDialog.GetFirstEditElement().SendText(roleName);
                AddRoleDialog.GetBtnElement("Add").LeftClick();
                Executor.GetInstance().SwitchTo(SessionType.PP5IDE);
                return true;
            //}
            //catch (Exception ex)
            //{
            //    if (CurrentDriver.GetElementFromWebElement(MobileBy.AccessibilityId("MessageBoxExDialog")) != null)
            //        return false;
            //    else
            //        throw;
            //}
        }

        bool AddDefectCode(int defectCode, string desc = "0", int customerDefectCode = 0)
        {
            Executor.GetInstance().SwitchTo(SessionType.Desktop);

            WaitUntil(() => CurrentDriver.GetExtendedElementBySingleWithRetry(PP5By.Name("Add Defect Code")) != null);
            var DefectCodeDialog = CurrentDriver.GetExtendedElementBySingleWithRetry(PP5By.Name("Add Defect Code"));
            if (DefectCodeDialog == null)
            {
                return false;
            }
            DefectCodeDialog.GetSpecificChildOfControlType(ElementControlType.TextBox, 0).SendText(defectCode.ToString());
            DefectCodeDialog.GetSpecificChildOfControlType(ElementControlType.TextBox, 1).SendText(desc);
            DefectCodeDialog.GetSpecificChildOfControlType(ElementControlType.TextBox, 2).SendText(customerDefectCode.ToString());
            DefectCodeDialog.GetBtnElement("OK").LeftClick();

            Executor.GetInstance().SwitchTo(SessionType.PP5IDE);
            return true;
        }

        bool EditUser(bool valid = true)
        {
            Executor.GetInstance().SwitchTo(SessionType.Desktop);

            var EditUserDialog = CurrentDriver.GetExtendedElementBySingleWithRetry(PP5By.Name("Edit User"));
            if (EditUserDialog == null)
            {
                return false;
            }

            if (valid)
                EditUserDialog.GetRdoBtnElement("Yes").LeftClick();
            else
                EditUserDialog.GetRdoBtnElement("No").LeftClick();

            EditUserDialog.GetBtnElement("OK").LeftClick();
            Executor.GetInstance().SwitchTo(SessionType.PP5IDE);

            return true;
        }

        bool EditRole(Dictionary<string, bool> authOptions)
        {
            Executor.GetInstance().SwitchTo(SessionType.Desktop);

            WaitUntil(() => CurrentDriver.GetExtendedElementBySingleWithRetry(PP5By.Name("Edit Role")) != null);
            var EditRoleDialog = CurrentDriver.GetExtendedElementBySingleWithRetry(PP5By.Name("Edit Role"));
            if (EditRoleDialog == null)
            {
                return false;
            }

            foreach (KeyValuePair<string, bool> authOpt in authOptions)
            {
                // "User", "Role", ...
                if (EditRoleDialog.GetCheckBoxElement(authOpt.Key).isElementChecked())
                {
                    if (!authOpt.Value)
                        EditRoleDialog.GetCheckBoxElement(authOpt.Key).LeftClick();
                }
                else
                {
                    if (authOpt.Value)
                        EditRoleDialog.GetCheckBoxElement(authOpt.Key).LeftClick();
                }
            }

            EditRoleDialog.GetBtnElement("OK").LeftClick();
            Executor.GetInstance().SwitchTo(SessionType.PP5IDE);
            return true;
        }
    }
}
