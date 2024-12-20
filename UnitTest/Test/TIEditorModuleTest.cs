﻿using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium;
using OpenQA.Selenium.Remote;

namespace PP5AutoUITests
{
    [TestClass]
    public class TIEditorModuleTest : TestBase
    {
        #region Combo Action Tests

        [TestInitialize]
        public void TIEditor_TestMethodSetup()
        {
            OpenNewTIEditorWindow();
        }

        [TestMethod]
        public void TestCopyAndPaste_SearchBoxFromTestCommandToTestEditCommand()
        {
            string testString = "testCopyAndPaste";

            IWebElement testCmdSearchTextBox = CurrentDriver.GetWebElementFromWebDriver(By.ClassName("CmdTreeView"))
                                                            .GetWebElementFromWebElement(MobileBy.AccessibilityId("searchText"));

            IWebElement testCmdEditSearchTextBox = CurrentDriver.GetWebElementFromWebDriver(By.ClassName("SettingAeraView"))
                                                                .GetWebElementFromWebElement(MobileBy.AccessibilityId("searchText"));

            // Clean up text in both search box
            testCmdSearchTextBox.ClearContent();
            testCmdEditSearchTextBox.ClearContent();

            // Input random text in testCmdSearchBox
            testCmdSearchTextBox.SendSingleKeys(testString);

            // Copy and paste text from command list searchbox to command edit searchbox
            testCmdSearchTextBox.CopyAndPaste(testCmdEditSearchTextBox);

            // Check input text same as expected
            testString.ShouldEqualTo(testCmdEditSearchTextBox.Text);
        }

        [TestMethod]
        public void TestCutAndPaste_SearchBoxFromTestCommandToTestEditCommand()
        {
            string testString = "testCutAndPaste";

            IWebElement testCmdSearchTextBox = CurrentDriver.GetWebElementFromWebDriver(By.ClassName("CmdTreeView"))
                                                            .GetWebElementFromWebElement(MobileBy.AccessibilityId("searchText"));

            IWebElement testCmdEditSearchTextBox = CurrentDriver.GetWebElementFromWebDriver(By.ClassName("SettingAeraView"))
                                                                .GetWebElementFromWebElement(MobileBy.AccessibilityId("searchText"));

            // Clean up text in both search box
            testCmdSearchTextBox.ClearContent();
            testCmdEditSearchTextBox.ClearContent();

            // Input random text in testCmdSearchBox
            testCmdSearchTextBox.SendSingleKeys(testString);

            // Cut and paste text from command list searchbox to command edit searchbox
            testCmdSearchTextBox.CutAndPaste(testCmdEditSearchTextBox);

            // Check input text same as expected
            testString.ShouldEqualTo(testCmdEditSearchTextBox.Text);
        }

        [TestMethod]
        [TestCategory("ByCommandName")]
        public void TestAddCommand_CommandVisibleInList_CheckSameCommandABS()
        {
            // Switch to test item context window
            TestItemTabNavi(TestItemTabType.TIContext);

            // Input Command name: "ABS", first command in Group: "Arithmetic"
            string CommandName = "ABS";
            TestCmdGroupType cgt = TestCmdGroupType.Arithmetic;

            // Add the command
            AddCommandBy(cgt, CommandName);

            // Check command name the same
            CommandName.ShouldEqualTo(GetCellValue("PGGrid", 1, "Test Command"));
        }

        [TestMethod]
        [TestCategory("ByCommandName")]
        public void TestAddCommand_CommandNotVisibleInList_CheckSameCommandXOR()
        {
            // Switch to test item context window
            TestItemTabNavi(TestItemTabType.TIContext);

            // Input Command name: "XOR", last command in Group: "Arithmetic"
            string CommandName = "XOR";
            TestCmdGroupType cgt = TestCmdGroupType.Arithmetic;

            // Add the command
            AddCommandBy(cgt, CommandName);

            // Check command name the same
            CommandName.ShouldEqualTo(GetCellValue("PGGrid", 1, "Test Command"));
        }

        [TestMethod]
        [TestCategory("ByCommandNumber")]
        [DataRow("ABS")]
        public void TestAddCommand_CommandVisibleInList_CheckSameCommandABSInNumber1(string commandName)
        {
            // Switch to test item context window
            TestItemTabNavi(TestItemTabType.TIContext);

            // Input Command name: "ABS", first command in Group: "Arithmetic"
            int CommandNumber = 1;
            TestCmdGroupType cgt = TestCmdGroupType.Arithmetic;

            // Add the command
            AddCommandBy(cgt, CommandNumber);

            // Check command name the same
            commandName.ShouldEqualTo(GetCellValue("PGGrid", 1, "Test Command"));
        }

        [TestMethod]
        [TestCategory("ByCommandNumber")]
        [DataRow("XOR")]
        public void TestAddCommand_CommandNotVisibleInList_CheckSameCommandXORInNumber25(string commandName)
        {
            // Switch to test item context window
            TestItemTabNavi(TestItemTabType.TIContext);

            // Input Command name: "XOR", last command in Group: "Arithmetic"
            int CommandNumber = 25;
            TestCmdGroupType cgt = TestCmdGroupType.Arithmetic;

            // Add the command
            AddCommandBy(cgt, CommandNumber);

            // Check command name the same
            commandName.ShouldEqualTo(GetCellValue("PGGrid", 1, "Test Command"));
        }

        [TestMethod]
        [TestCategory("ByCommandNumber")]
        public void TestAddCommand_GroupNotVisibleInList_CheckSameCommandABSInNumber1()
        {
            // Switch to test item context window
            TestItemTabNavi(TestItemTabType.TIContext);

            // Input Command name: "GetTDL_Information", first command in Group: "Data Logger"
            // Where group is not visible in command list
            int CommandNumber = 1;
            TestCmdGroupType cgt = TestCmdGroupType.Data_Logger;

            // Add the command
            AddCommandBy(cgt, CommandNumber);

            // Check command name the same
            "GetTDL_Information".ShouldEqualTo(GetCellValue("PGGrid", 1, "Test Command"));
        }

        [TestMethod]
        [TestCategory("ByCommandName")]
        public void TestAddCommand_NotInListByCommandName_ShouldReturnCommandNameNotExistedException()
        {
            //// Arrange
            // Input Command name that is not in Group: "Arithmetic"
            string CommandName = "XXXXXXXXXX";
            TestCmdGroupType cgt = TestCmdGroupType.Arithmetic;

            //// Act + Assert
            // Add the command, check exception message
            Exception exp = Assert.ThrowsException<CommandNameNotExistedException>(() => AddCommandBy(cgt, CommandName));
            CommandName.ShouldEqualTo(exp.Message);
        }

        [TestMethod]
        [TestCategory("ByCommandNumber")]
        public void TestAddCommand_NotInListByCommandNumber_ShouldReturnCommandNumberNotExistedException()
        {
            //// Arrange
            // Input Command number that is not in Group: "Arithmetic"
            int CommandNumber = 50;
            TestCmdGroupType cgt = TestCmdGroupType.Arithmetic;

            //// Act + Assert
            // Add the command, check exception message
            Exception exp = Assert.ThrowsException<CommandNumberOutOfRangeException>(() => AddCommandBy(cgt, CommandNumber));
            CommandNumber.ToString().ShouldEqualTo(exp.Message);
        }

        [TestMethod]
        [TestCategory("ByCommandName")]
        public void TestAddCommands_HasCommandNameNotInList_ShouldReturnCommandNameNotExistedException()
        {
            //// Arrange
            // Input Command name that is not in Group: "Arithmetic"
            string WrongCommandName1 = "FFFFFFFFFF";
            string WrongCommandName2 = "GGGGGGGGGG";
            TestCmdGroupType cgt = TestCmdGroupType.Arithmetic;

            //// Act + Assert
            // Add the commands, check exception message
            CommandNameNotExistedException exp = Assert.ThrowsException<CommandNameNotExistedException>(() => AddCommandsBy(cgt, WrongCommandName1, WrongCommandName2, "MOD"));
            WrongCommandName1.ShouldEqualTo(exp.CommandName);
        }

        [TestMethod]
        [TestCategory("ByCommandNumber")]
        public void TestAddCommands_WrongCommandNumber0_ShouldReturnCommandNumberNotExistedException()
        {
            //// Arrange
            // The group name
            TestCmdGroupType cgt = TestCmdGroupType.Arithmetic;

            //// Act + Assert
            // Add the commands with No: 0, 1, 10
            // Command number: 0 is not in Group: "Arithmetic"
            CommandNumberOutOfRangeException exp = Assert.ThrowsException<CommandNumberOutOfRangeException>(() => AddCommandsBy(cgt, 0, 1, 10));

            // check exception message
            0.ShouldEqualTo(exp.CommandNumber);
        }

        [TestMethod]
        [TestCategory("ByCommandNumber")]
        public void TestAddCommands_WrongCommandNumber26_ShouldReturnCommandNumberNotExistedException()
        {
            //// Arrange
            // The group name
            TestCmdGroupType cgt = TestCmdGroupType.Arithmetic;

            //// Act + Assert
            // Add the commands with No: 1, 10, 50
            // Command number: 50 is not in Group: "Arithmetic"
            CommandNumberOutOfRangeException exp = Assert.ThrowsException<CommandNumberOutOfRangeException>(() => AddCommandsBy(cgt, 1, 10, 50));

            // check exception message
            50.ShouldEqualTo(exp.CommandNumber);
        }

        //[TestMethod]
        //[TestCategory("ByCommandName")]
        //public void TestAddCommand_GroupNameNotInList_ShouldReturnGroupNameNotExistedException()
        //{
        //    //// Arrange
        //    // The group name not existed
        //    string GroupName = "xxxxxxxxxx";
        //    string CommandName = "ABS";

        //    //// Act + Assert
        //    // Add the command, check exception message
        //    GroupNameNotExistedException exp = Assert.ThrowsException<GroupNameNotExistedException>(() => AddCommandBy(GroupName, CommandName));
        //    GroupName.ShouldEqualTo(exp.GroupName);
        //}

        //[TestMethod]
        //[TestCategory("ByCommandNumber")]
        //public void TestAddCommand2_GroupNameNotInList_ShouldReturnGroupNameNotExistedException()
        //{
        //    //// Arrange
        //    // The group name not existed
        //    string GroupName = "xxxxxxxxxx";
        //    int CommandNumber = 26;

        //    //// Act + Assert
        //    // Add the command, check exception message
        //    GroupNameNotExistedException exp = Assert.ThrowsException<GroupNameNotExistedException>(() => AddCommandBy(GroupName, CommandNumber));
        //    GroupName.ShouldEqualTo(exp.GroupName);
        //}

        [TestMethod]
        [TestCategory("ByCommandName")]
        public void TestAddCommands_Add3CommandsByCommandName_ShouldAddToCommandEditPage()
        {
            // Switch to test item context window
            TestItemTabNavi(TestItemTabType.TIContext);

            // Command Group
            TestCmdGroupType cgt = TestCmdGroupType.AC_Source;

            // Add the command with command names: "ABS", "MOD", "XOR"
            // { "AC Source",  new[] { "ReadAC_Current", "ReadAC_Frequency" } },
            AddCommandsBy(cgt, "ReadAC_Current", "ReadAC_Frequency", "SetACDev_CPPHParameter2");

            // Check commands are correctly added
            CollectionAssert.AreEqual(new List<string> { "ReadAC_Current", "ReadAC_Frequency", "SetACDev_CPPHParameter2" }, GetSingleColumnValues(DataTableAutoIDType.PGGrid, 5/*Test Command*/));
        }

        [TestMethod]
        [TestCategory("ByCommandNumber")]
        public void TestAddCommands_Add3CommandsByCommandNumber_ShouldAddToCommandEditPage()
        {
            // Switch to test item context window
            TestItemTabNavi(TestItemTabType.TIContext);

            // Command Group
            TestCmdGroupType cgt = TestCmdGroupType.Arithmetic;

            // Add the commands with indexces: 1, 10, 25 (ABS, LOG, XOR)
            AddCommandsBy(cgt, 1, 10, 25);

            // Check commands are correctly added
            CollectionAssert.AreEqual(new List<string>{ "ABS", "LOG", "XOR" }, GetSingleColumnValues(DataTableAutoIDType.PGGrid, 5/*Test Command*/));
        }

        #endregion
    }
}
