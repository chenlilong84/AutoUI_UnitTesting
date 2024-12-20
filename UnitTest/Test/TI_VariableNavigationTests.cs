using System;
using Chroma.UnitTest.Common;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace PP5AutoUITests
{
    [TestClass]
    public class VariableNavigationTests : TestBase
    {
        [TestInitialize]
        public void TIEditor_TestMethodSetup()
        {
            AutoUIExecutor.SwitchTo(SessionType.PP5IDE);
        }

        // Condition Tab Column Navigation Tests
        [TestMethod("ConditionTab_NavigateFromNoToShowName")]
        [TestCategory("ConditionTab")]
        [DataRow(VariableDataType.String, VariableEditType.EditBox, VariableColumnType.EditType, VariableColumnType.No, typeof(object), DisplayName = "Navigate from EditType to No")]
        [DataRow(VariableDataType.Integer, VariableEditType.ComboBox, VariableColumnType.EditType, VariableColumnType.No, typeof(object), DisplayName = "Navigate from EditType to No with Integer")]
        public void ConditionTab_NavigateFromNoToShowName(
            VariableDataType varDataType,
            VariableEditType varEditType,
            VariableColumnType fromColumn,
            VariableColumnType toColumn,
            object dummy)
        {
            // Arrange
            var tabType = VariableTabType.Condition;
            //string callName = $"Test-{Guid.NewGuid()}";
            PP5DataGrid dataGrid = InitializeVariableDataGrid(tabType, "", "a", varDataType, varEditType);

            // Act & Assert
            try
            {
                VariableSelectionMoveTo(
                    tabType,
                    varDataType,
                    varEditType,
                    fromColumn,
                    toColumn
                );
                dataGrid.RefreshSelectedCell();
                toColumn.GetDescription().ShouldEqualTo(dataGrid.SelectedCellInfo.ColumnName);
            }
            catch (Exception ex)
            {
                Assert.Fail($"Navigation failed: {ex.Message}");
            }
        }

        [TestMethod("ConditionTab_NavigateFromShowNameToCallName")]
        [TestCategory("ConditionTab")]
        [DataRow(VariableDataType.Float, VariableEditType.EditBox, VariableColumnType.ShowName, VariableColumnType.CallName, DisplayName = "Navigate from ShowName to CallName")]
        [DataRow(VariableDataType.FloatPercentage, VariableEditType.ComboBox, VariableColumnType.ShowName, VariableColumnType.CallName, DisplayName = "Navigate from ShowName to CallName with FloatPercentage")]
        public void ConditionTab_NavigateFromShowNameToCallName(
            VariableDataType varDataType,
            VariableEditType varEditType,
            VariableColumnType fromColumn,
            VariableColumnType toColumn)
        {
            // Arrange
            var tabType = VariableTabType.Condition;
            string callName = $"Test-{Guid.NewGuid()}";

            // Act & Assert
            try
            {
                VariableSelectionMoveTo(
                    tabType,
                    varDataType,
                    varEditType,
                    fromColumn,
                    toColumn
                );
            }
            catch (Exception ex)
            {
                Assert.Fail($"Navigation failed: {ex.Message}");
            }
        }

        // Result Tab Column Navigation Tests
        [TestMethod("ResultTab_NavigateFromNoToMinimumSpec")]
        [TestCategory("ResultTab")]
        [DataRow(VariableDataType.Double, VariableEditType.EditBox, VariableColumnType.No, VariableColumnType.MinimumSpec, DisplayName = "Navigate from No to MinimumSpec")]
        [DataRow(VariableDataType.Integer, VariableEditType.ComboBox, VariableColumnType.No, VariableColumnType.MinimumSpec, DisplayName = "Navigate from No to MinimumSpec with Integer")]
        public void ResultTab_NavigateFromNoToMinimumSpec(
            VariableDataType varDataType,
            VariableEditType varEditType,
            VariableColumnType fromColumn,
            VariableColumnType toColumn)
        {
            // Arrange
            var tabType = VariableTabType.Result;
            string callName = $"Test-{Guid.NewGuid()}";

            // Act & Assert
            try
            {
                VariableSelectionMoveTo(
                    tabType,
                    varDataType,
                    varEditType,
                    fromColumn,
                    toColumn
                );
            }
            catch (Exception ex)
            {
                Assert.Fail($"Navigation failed: {ex.Message}");
            }
        }

        // Global Tab Column Navigation Tests
        [TestMethod("GlobalTab_NavigateLockToNo")]
        [TestCategory("GlobalTab")]
        [DataRow(VariableDataType.String, VariableEditType.EditBox, VariableColumnType.Lock, VariableColumnType.No, DisplayName = "Navigate from Lock to No")]
        [DataRow(VariableDataType.Long, VariableEditType.ComboBox, VariableColumnType.Lock, VariableColumnType.No, DisplayName = "Navigate from Lock to No with Long")]
        public void GlobalTab_NavigateLockToNo(
            VariableDataType varDataType,
            VariableEditType varEditType,
            VariableColumnType fromColumn,
            VariableColumnType toColumn)
        {
            // Arrange
            var tabType = VariableTabType.Global;
            string callName = $"Test-{Guid.NewGuid()}";

            // Act & Assert
            try
            {
                VariableSelectionMoveTo(
                    tabType,
                    varDataType,
                    varEditType,
                    fromColumn,
                    toColumn
                );
            }
            catch (Exception ex)
            {
                Assert.Fail($"Navigation failed: {ex.Message}");
            }
        }

        // Temp Tab Column Navigation Tests
        [TestMethod("TempTab_NavigateFromShowNameToDataType")]
        [TestCategory("TempTab")]
        [DataRow(VariableDataType.Float, VariableEditType.EditBox, VariableColumnType.ShowName, VariableColumnType.DataType, DisplayName = "Navigate from ShowName to DataType")]
        [DataRow(VariableDataType.Integer, VariableEditType.ComboBox, VariableColumnType.ShowName, VariableColumnType.DataType, DisplayName = "Navigate from ShowName to DataType with Integer")]
        public void TempTab_NavigateFromShowNameToDataType(
            VariableDataType varDataType,
            VariableEditType varEditType,
            VariableColumnType fromColumn,
            VariableColumnType toColumn)
        {
            // Arrange
            var tabType = VariableTabType.Temp;
            string callName = $"Test-{Guid.NewGuid()}";

            // Act & Assert
            try
            {
                VariableSelectionMoveTo(
                    tabType,
                    varDataType,
                    varEditType,
                    fromColumn,
                    toColumn
                );
            }
            catch (Exception ex)
            {
                Assert.Fail($"Navigation failed: {ex.Message}");
            }
        }

        // Error Scenario Tests
        [TestMethod("ErrorScenario_InvalidColumnNavigation")]
        [TestCategory("ErrorHandling")]
        [DataRow(VariableDataType.String, VariableEditType.EditBox, VariableColumnType.No, VariableColumnType.DefectCodeMin, DisplayName = "Navigate between incompatible columns")]
        public void ErrorScenario_InvalidColumnNavigation(
            VariableDataType varDataType,
            VariableEditType varEditType,
            VariableColumnType fromColumn,
            VariableColumnType toColumn)
        {
            // Arrange
            var tabType = VariableTabType.Condition;

            // Act & Assert
            Assert.ThrowsException<ArgumentException>(() =>
            {
                VariableSelectionMoveTo(
                    tabType,
                    varDataType,
                    varEditType,
                    fromColumn,
                    toColumn
                );
            }, "Should throw ArgumentException for invalid column navigation");
        }
    }
}
