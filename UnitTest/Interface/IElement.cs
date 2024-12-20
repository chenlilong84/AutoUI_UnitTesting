using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium.Windows;

namespace PP5AutoUITests
{
    public interface IElement : IWebElement
    {
        #region Element Identification Properties Signature
        string Name { get; }

        string AutomationId { get; }

        string ClassName { get; }

        string FrameworkId { get; }
        #endregion

        #region Element Type Properties Signature
        ElementControlType ControlType { get; }
        string LocalizedControlType { get; }
        #endregion

        #region Element Display Characteristics (顯示特徵) Signature
        int Width { get; }
        int Height { get; }
        Point PointAtTopLeft { get; }
        Point PointAtCenter { get; }
        Point PointAtTopCenter { get; }
        Point PointAtBottomCenter { get; }
        Point PointAtLeftCenter { get; }
        Point PointAtRightCenter { get; }

        bool IsWithinScreen { get; }
        string ToolTipText { get; }
        #endregion

        #region Element Patterns Signature
        public bool HasAnnotationPattern { get; }

        public bool HasDragPattern { get; }

        public bool CanDock { get; }

        public bool HasDropTargetPattern { get; }

        public bool CanExpandOrCollapse { get; }

        public bool HasGridItemPattern { get; }

        public bool HasGridPattern { get; }

        public bool CanInvoke { get; }

        public bool IsItemContainerPatternAvailable { get; }

        public bool HasLegacyIAccessiblePattern { get; }

        public bool IsMultipleViewPatternAvailable { get; }

        public bool IsObjectModelPatternAvailable { get; }

        public bool HasRangeValuePattern { get; }

        public bool HasScrollItemPattern { get; }

        public bool CanScroll { get; }

        //public bool IsSelectionItemPatternAvailable { get; }

        public bool CanSelect { get; }

        public bool IsSpreadsheetItemPatternAvailable { get; }

        public bool IsSpreadsheetPatternAvailable { get; }

        public bool IsStylesPatternAvailable { get; }

        public bool IsSynchronizedInputPatternAvailable { get; }

        public bool IsTableItemPatternAvailable { get; }

        public bool HasTablePattern { get; }

        public bool IsTextChildPatternAvailable { get; }

        public bool IsTextEditPatternAvailable { get; }

        public bool HasTextPattern { get; }

        public bool IsTextPattern2Available { get; }

        public bool CanToggle { get; }

        public bool HasTransformPattern { get; }

        public bool IsTransform2PatternAvailable { get; }

        public bool HasValuePattern { get; }

        public bool IsVirtualizedItemPatternAvailable { get; }

        public bool IsCustomNavigationPatternAvailable { get; }

        public bool IsSelectionPattern2Available { get; }
        #endregion

        public string ModuleName { get; }
        public bool CanScrollHorizontally { get; }
        public bool CanScrollVertically { get; }

        public bool IsWindow { get; }
        public bool IsTextBox { get; }
        bool IsCheckBox { get; }
        bool IsTextBlock { get; }
        bool IsButton { get; }
        bool IsRadioButton { get; }
        bool IsComboBox { get; }
        bool IsProgressBar { get; }
        bool IsTreeView { get; }
        bool IsImage { get; }
        bool IsGridCell { get; }
        

        public string GetText();

        public void SendText(string keysToSend);

        public void LeftClick();

        public void DoubleClick();

        public new IElement FindElement(By by);

        public IElement FindElement(ByAutomationIdOrName by);

        public new ReadOnlyCollection<IElement> FindElements(By by);
        bool IsSameElement(IElement elementToCompare);

        //public void RightClick();
        //public void DoubleClick();
    }
}
