using System.Collections;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
using Chroma.UnitTest.Common;
using Microsoft.VisualStudio.TestTools.UnitTesting.Logging;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Remote;
using PP5AutoUITests.Model;

namespace PP5AutoUITests
{
    public class PP5Element : WindowsElement, IElement
    {
        /// <summary>
        /// Initializes a new instance of the Element class.
        /// </summary>
        /// <param name="parent">Driver in use.</param>
        /// <param name="id">ID of the element.</param>
        public PP5Element(WindowsElement element)
            : base((RemoteWebDriver)element.WrappedDriver, element.Id)
        {

        }

        public PP5Element(AppiumWebElement element)
            : base((RemoteWebDriver)element.WrappedDriver, element.Id)
        {

        }

        //public new PP5Driver WrappedDriver => (PP5Driver)base.WrappedDriver;

        public event PropertyChangedEventHandler PropertyChanged;
        public virtual void OnPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (PropertyChanged != null)
                PropertyChanged(sender, e);
        }

        protected void NotifyPropertyChanged<T>(string propertyName, T oldvalue, T newvalue)
        {
            OnPropertyChanged(this, new PropertyChangedExtendedEventArgs<T>(propertyName, oldvalue, newvalue));
        }

        public string ModuleName
        {
            get
            {
                string[] titleParts = base.WrappedDriver.Title.Split('[');
                if (titleParts.Length == 1)
                    return WindowType.None.GetDescription();
                else
                   return titleParts.Last().Split('-')[0].Trim().TrimEnd(']');
            }
        }

        public string ShowName => ToString();

        #region Element Identification Properties
        public string Name => base.GetAttribute("Name");

        public string AutomationId => base.GetAttribute("AutomationId");

        public string ClassName => base.GetAttribute("ClassName");

        public string FrameworkId => base.GetAttribute("FrameworkId");
        #endregion

        #region Element Type Properties
        /// <summary>
        /// LocalizedControlTypeProperty（本地化控件类型）
        /// <para>本地化字符串所描述的控件類型，根據不同語言會有該語言的控件名稱，用於多語言自動化測試</para>
        /// </summary>
        public string LocalizedControlType => base.GetAttribute("LocalizedControlType");

        /// <summary>
        /// ControlTypeProperty（控件类型）
        /// <para>控件類型，以英文表示，不會根據不同語言而改變</para>
        /// </summary>
        //public string ControlType => base.GetAttribute("ControlType");
        public ElementControlType ControlType => TypeExtension.GetEnumByDescription<ElementControlType>(base.GetAttribute("ControlType").Replace("ControlType.", ""));
        #endregion

        #region Element Display Characteristics (顯示特徵)

        /// <summary>
        /// 當Element支援ValuePattern時，可使用此數值
        /// </summary>
        public string Value => HasValuePattern ? GetAttribute("Value.Value") : null;

        /// <summary>
        /// <para>The Width of an Element.</para>
        /// </summary>
        private int width;
        public int Width
        {
            get
            {
                width = this.Size.Width;
                return width;
            }
            set
            {
                int oldValue = width;
                width = value;
                NotifyPropertyChanged("Width", oldValue, width);
            }
        }

        /// <summary>
        /// <para>The Height of an Element.</para>
        /// </summary>
        private int height;
        public int Height
        {
            get
            {
                height = this.Size.Height;
                return height;
            }
            set
            {
                int oldValue = height;
                height = value;
                NotifyPropertyChanged("Height", oldValue, height);
            }
        }

        /// <summary>
        /// <para>The top-left point <b>(origin)</b> of an Element.</para>
        /// </summary>
        /// <value><strong>Point</strong> in ( <i>x</i>, <i>y</i> )</value>
        public Point PointAtTopLeft => new Point(this.Rect.Left, this.Rect.Top);

        /// <summary>
        /// <para>The center point of an Element.</para>
        /// </summary>
        /// <value><strong>Point</strong> in ( <i>x</i>, <i>y</i> )</value>
        public Point PointAtCenter => new Point((this.Rect.Left + this.Rect.Right) / 2, (this.Rect.Top + this.Rect.Bottom) / 2);

        /// <summary>
        /// <para>The top-center point of an Element.</para>
        /// </summary>
        /// <value><strong>Point</strong> in ( <i>x</i>, <i>y</i> )</value>
        public Point PointAtTopCenter => new Point((this.Rect.Left + this.Rect.Right) / 2, this.Rect.Top);

        /// <summary>
        /// <para>The bottom-center point of an Element.</para>
        /// </summary>
        /// <value><strong>Point</strong> in ( <i>x</i>, <i>y</i> )</value>
        public Point PointAtBottomCenter => new Point((this.Rect.Left + this.Rect.Right) / 2, this.Rect.Bottom);

        /// <summary>
        /// <para>The left-center point of an Element.</para>
        /// </summary>
        /// <value><strong>Point</strong> in ( <i>x</i>, <i>y</i> )</value>
        public Point PointAtLeftCenter => new Point(this.Rect.Left, (this.Rect.Top + this.Rect.Bottom) / 2);

        /// <summary>
        /// <para>The right-center point of an Element.</para>
        /// </summary>
        /// <value><strong>Point</strong> in ( <i>x</i>, <i>y</i> )</value>
        public Point PointAtRightCenter => new Point(this.Rect.Right, (this.Rect.Top + this.Rect.Bottom) / 2);

        /// <summary>
        /// <para>Get the ToolTip message of an Element.</para>
        /// </summary>
        public string ToolTipText => GetAttribute("HelpText");

        /// <summary>
        /// <para>Indicator for checking an Element is within screen.</para>
        /// </summary>
        /// <value><see langword="true"/> if element is in screen（元素在屏幕內）,otherwise <see langword="false"/></value>
        public bool IsWithinScreen => !bool.Parse(GetAttribute("IsOffscreen"));

        public override bool Displayed
        {
            get 
            {
                bool result;
                try
                {
                    result = base.Displayed;
                }
                catch (Exception)
                {
                    result = false;
                }
                // Log the Action
                return result;
            }
        }

        /// <summary>
        /// ScrollBar Vertical Position from (0 ~ 100%)
        /// </summary>
        private string _ScrollBarVerticalPosition => GetAttribute("Scroll.VerticalScrollPercent");

        public string ScrollBarVerticalPosition
        {
            get
            {
                return _ScrollBarVerticalPosition;
            }
            set
            {
                if (value != _ScrollBarVerticalPosition)
                {
                    //string oldValue = _ScrollBarPosition;
                    //_ScrollBarPosition = value;
                    NotifyPropertyChanged("ScrollBarVerticalPosition", value, _ScrollBarVerticalPosition);
                }
            }
        }

        /// <summary>
        /// ScrollBar Horizontal Position from (0 ~ 100%)
        /// </summary>
        private string _ScrollBarHorizontalPosition => GetAttribute("Scroll.HorizontalScrollPercent");

        public string ScrollBarHorizontalPosition
        {
            get
            {
                return _ScrollBarHorizontalPosition;
            }
            set
            {
                if (value != _ScrollBarHorizontalPosition)
                {
                    //string oldValue = _ScrollBarPosition;
                    //_ScrollBarPosition = value;
                    NotifyPropertyChanged("ScrollBarHorizontalPosition", value, _ScrollBarHorizontalPosition);
                }
            }
        }
        #endregion

        #region Element Patterns
        public bool HasAnnotationPattern => bool.Parse(GetAttribute("IsAnnotationPatternAvailable"));

        public bool HasDragPattern => bool.Parse(GetAttribute("IsDragPatternAvailable"));

        /// <summary>
        /// 是否支援 DockPattern（停靠模式）
        /// <para>說明：</para>
        /// <para>
        /// <list type="bullet">
        ///     <item>用途：用于可以在容器中停靠的控件，如工具栏、停靠面板。</item>
        ///     <item>
        ///         <description>属性：</description>
        ///         <list type="table">
        ///             <item>DockPosition：控件在容器中的停靠位置（顶部、底部、左侧、右侧、填充、无）。</item>
        ///         </list>
        ///     </item>
        ///     <item>
        ///         <description>方法：</description>
        ///         <list type="table">
        ///             <item>SetDockPosition(position)：设置控件的停靠位置。</item>
        ///         </list>
        ///     </item>
        /// </list>
        /// </para>
        /// <para>支持的控件：</para>
        /// <para>
        /// <list type="bullet">
        ///     <item>工具栏（ToolBar）</item>
        ///     <item>停靠面板中的控件</item>
        /// </list>
        /// </para>
        /// </summary>
        public bool CanDock => bool.Parse(GetAttribute("IsDockPatternAvailable"));

        public bool HasDropTargetPattern => bool.Parse(GetAttribute("IsDropTargetPatternAvailable"));

        /// <summary>
        /// 是否支援 ExpandCollapsePattern（展开折叠模式）
        /// <para>說明：</para>
        /// <para>
        /// <list type="bullet">
        ///     <item>用途：用于可以展开或折叠的控件，如树节点、组合框。</item>
        ///     <item>
        ///         <description>属性：</description>
        ///         <list type="table">
        ///             <item>ExpandCollapseState：当前的展开或折叠状态。</item>
        ///         </list>
        ///     </item>
        ///     <item>
        ///         <description>方法：</description>
        ///         <list type="table">
        ///             <item>Expand()：展开控件。</item>
        ///             <item>Collapse()：折叠控件。</item>
        ///         </list>
        ///     </item>
        /// </list>
        /// </para>
        /// <para>支持的控件：</para>
        /// <para>
        /// <list type="bullet">
        ///     <item>树节点（TreeItem）</item>
        ///     <item>组合框（ComboBox）</item>
        ///     <item>菜单项（MenuItem）</item>
        /// </list>
        /// </para>
        /// <para>Example:</para>
        /// <example>
        /// <code>
        /// // 展开组合框
        /// WindowsElement comboBox = driver.FindElementByAccessibilityId("ComboBoxAutomationId");
        /// driver.ExecuteScript("arguments[0].Expand();", comboBox);
        /// 
        /// // 检查展开状态
        /// string state = comboBox.GetAttribute("ExpandCollapse.ExpandCollapseState");
        /// </code>
        /// </example>
        /// </summary>
        public bool CanExpandOrCollapse => bool.Parse(GetAttribute("IsExpandCollapsePatternAvailable"));

        /// <summary>
        /// 是否支援 GridItemPattern（网格项模式）
        /// <para>說明：</para>
        /// <para>
        /// <list type="bullet">
        ///     <item>用途：用于网格中的单元格元素。</item>
        ///     <item>
        ///         <description>属性：</description>
        ///         <list type="table">
        ///             <item>Row：单元格所在行索引。</item>
        ///             <item>Column：单元格所在列索引。</item>
        ///             <item>RowSpan：跨越的行数。</item>
        ///             <item>ColumnSpan：跨越的列数。</item>
        ///             <item>ContainingGrid：包含该单元格的网格控件。</item>
        ///         </list>
        ///     </item>
        ///     <item>方法：無</item>
        /// </list>
        /// </para>
        /// <para>支持的控件：</para>
        /// <para>
        /// <list type="bullet">
        ///     <item>网格中的单元格（DataItem、Custom）</item>
        /// </list>
        /// </para>
        /// </summary>
        public bool HasGridItemPattern => bool.Parse(GetAttribute("IsGridItemPatternAvailable"));

        /// <summary>
        /// 是否支援 GridPattern（网格模式）
        /// <para>說明：</para>
        /// <para>
        /// <list type="bullet">
        ///     <item>用途：用于表示二维网格的控件，如数据表、网格。</item>
        ///     <item>
        ///         <description>属性：</description>
        ///         <list type="table">
        ///             <item>RowCount：行数。</item>
        ///             <item>ColumnCount：列数。</item>
        ///         </list>
        ///     </item>
        ///     <item>
        ///         <description>方法：</description>
        ///         <list type="table">
        ///             <item>GetItem(row, column)：获取指定单元格的元素。</item>
        ///         </list>
        ///     </item>
        /// </list>
        /// </para>
        /// <para>支持的控件：</para>
        /// <para>
        /// <list type="bullet">
        ///     <item>数据网格（DataGrid）</item>
        ///     <item>表格（Table）</item>
        /// </list>
        /// </para>
        /// </summary>
        public bool HasGridPattern => bool.Parse(GetAttribute("IsGridPatternAvailable"));

        /// <summary>
        /// 是否支援 InvokePattern（调用模式）
        /// <para>說明：</para>
        /// <para>
        /// <list type="bullet">
        ///     <item>用途：提供对控件的单一操作，如点击按钮、选择菜单项等。</item>
        ///     <item>属性：无特定属性。</item>
        ///     <item>
        ///         <description>方法：</description>
        ///         <list type="bullet">
        ///             <item>Invoke()：执行控件的调用操作。</item>
        ///         </list>
        ///     </item>
        /// </list>
        /// </para>
        /// <para>支持的控件：</para>
        /// <para>
        /// <list type="bullet">
        ///     <item>按钮（Button）</item>
        ///     <item>超链接（Hyperlink）</item>
        ///     <item>菜单项（MenuItem）</item>
        ///     <item>单选按钮（RadioButton）</item>
        ///     <item>图像（Image）（如果支持点击）</item>
        /// </list>
        /// </para>
        /// <para>Example:</para>
        /// <example>
        /// <code>
        /// element.Click();                                         // 假设 element 是一个支持 InvokePattern 的控件(使用 Click() 方法)
        /// driver.ExecuteScript("arguments[0].Invoke();", element); // 或者，如果需要直接调用 InvokePattern，可以使用 ExecuteScript     
        /// </code>
        /// </example>
        /// </summary>
        public bool CanInvoke => bool.Parse(GetAttribute("IsInvokePatternAvailable"));

        public bool IsItemContainerPatternAvailable => bool.Parse(GetAttribute("IsItemContainerPatternAvailable"));

        /// <summary>
        /// LegacyIAccessiblePattern（旧版辅助功能模式）
        /// <para>說明：</para>
        /// <para>
        /// <list type="bullet">
        ///     <item>用途：用于支持旧版的辅助功能接口，主要用于兼容性。</item>
        ///     <item>
        ///         <description>属性：</description>
        ///         <list type="table">
        ///             <item>ChildId：子元素ID</item>
        ///             <item>DefaultAction：默认操作描述。</item>
        ///             <item>Description：控件描述。</item>
        ///             <item>Name、Role、State等。</item>
        ///         </list>
        ///     </item>
        ///     <item>
        ///         <description>方法：</description>
        ///         <list type="table">
        ///             <item>DoDefaultAction()：执行默认操作。</item>
        ///             <item>Select()：选择控件。</item>
        ///             <item>GetSelection()：获取选定项。</item>
        ///         </list>
        ///     </item>
        /// </list>
        /// </para>
        /// <para>支持的控件：</para>
        /// <para>
        /// <list type="bullet">
        ///     <item>旧版控件或不完全支持UI自动化的控件</item>
        /// </list>
        /// </para>
        /// </summary>
        public bool HasLegacyIAccessiblePattern => bool.Parse(GetAttribute("IsLegacyIAccessiblePatternAvailable"));

        public bool IsMultipleViewPatternAvailable => bool.Parse(GetAttribute("IsMultipleViewPatternAvailable"));

        public bool IsObjectModelPatternAvailable => bool.Parse(GetAttribute("IsObjectModelPatternAvailable"));

        /// <summary>
        /// 是否支援 RangeValuePattern（范围值模式）
        /// <para>說明：</para>
        /// <para>
        /// <list type="bullet">
        ///     <item>用途：用于具有数值范围的控件，可设置或获取其数值位置。</item>
        ///     <item>
        ///         <description>属性：</description>
        ///         <list type="table">
        ///             <item>Value：当前值。</item>
        ///             <item>Minimum：最小值。</item>
        ///             <item>Maximum：最大值。</item>
        ///             <item>LargeChange：大步进值。</item>
        ///             <item>SmallChange：小步进值。</item>
        ///         </list>
        ///     </item>
        ///     <item>方法：無</item>
        /// </list>
        /// </para>
        /// <para>支持的控件：</para>
        /// <para>
        /// <list type="bullet">
        ///     <item>滑块（Slider）</item>
        ///     <item>进度条（ProgressBar）</item>
        ///     <item>滚动条（ScrollBar）</item>
        ///     <item>微调控件（Spinner）</item>
        /// </list>
        /// </para>
        /// <para>Example:</para>
        /// <example>
        /// <code>
        /// // 获取范围值属性
        /// string currentValue = element.GetAttribute("RangeValue.Value");
        /// string minimum = element.GetAttribute("RangeValue.Minimum");
        /// string maximum = element.GetAttribute("RangeValue.Maximum");
        /// 
        /// // 设置新值
        /// driver.ExecuteScript("arguments[0].SetValue(50);", element); // 将值设置为50 
        /// </code>
        /// </example>
        /// </summary>
        public bool HasRangeValuePattern => bool.Parse(GetAttribute("IsRangeValuePatternAvailable"));

        public bool HasScrollItemPattern => bool.Parse(GetAttribute("IsScrollItemPatternAvailable"));

        /// <summary>
        /// 是否支援 ScrollPattern（滚动模式）
        /// <para>說明：</para>
        /// <para>
        /// <list type="bullet">
        ///     <item>用途：ScrollPattern（滚动模式）</item>
        ///     <item>
        ///         <description>属性：</description>
        ///         <list type="table">
        ///             <item>HorizontallyScrollable：是否可水平滚动。</item>
        ///             <item>VerticalScrollPercent：当前垂直滚动位置百分比。</item>
        ///             <item>HorizontalViewSize：水平视图大小。</item>
        ///         </list>
        ///     </item>
        ///     <item>
        ///         <description>方法：</description>
        ///         <list type="table">
        ///             <item>Scroll()：滚动指定的量。</item>
        ///             <item>SetScrollPercent()：设置滚动位置百分比。</item>
        ///         </list>
        ///     </item>
        /// </list>
        /// </para>
        /// <para>支持的控件：</para>
        /// <para>
        /// <list type="bullet">
        ///     <item>滚动视图（ScrollViewer）</item>
        ///     <item>列表框（ListBox）</item>
        ///     <item>文本框（TextBox）（多行）</item>
        /// </list>
        /// </para>
        /// <para>Example:</para>
        /// <example>
        /// <code>
        /// // 滚动到指定位置
        /// driver.ExecuteScript("arguments[0].SetScrollPercent(50, 0);", element); // 水平50%，垂直0%
        /// 
        /// // 滚动指定量
        /// driver.ExecuteScript("arguments[0].Scroll('LargeIncrement', 'NoAmount');", element);
        /// </code>
        /// </example>
        /// </summary>
        public bool CanScroll => bool.Parse(GetAttribute("IsScrollPatternAvailable"));

        public bool CanScrollHorizontally => CanScroll && bool.Parse(GetAttribute("Scroll.HorizontallyScrollable"));

        public bool CanScrollVertically => CanScroll && bool.Parse(GetAttribute("Scroll.VerticallyScrollable"));

        /// <summary>
        /// SelectionItemPattern（选择项模式）
        /// <para>說明：</para>
        /// <para>
        /// <list type="bullet">
        ///     <item>用途：用于可被选中的子项，如列表项、树节点。</item>
        ///     <item>
        ///         <description>属性：</description>
        ///         <list type="table">
        ///             <item>IsSelected：是否被选中。</item>
        ///             <item>SelectionContainer：包含此项的控件。</item>
        ///         </list>
        ///     </item>
        ///     <item>
        ///         <description>方法：</description>
        ///         <list type="table">
        ///             <item>Select()：选择该项。</item>
        ///             <item>AddToSelection()：将该项添加到选择。</item>
        ///             <item>RemoveFromSelection()：从选择中移除该项。</item>
        ///         </list>
        ///     </item>
        /// </list>
        /// </para>
        /// <para>支持的控件：</para>
        /// <para>
        /// <list type="bullet">
        ///     <item>列表项（ListItem）</item>
        ///     <item>组合框项（ComboBoxItem）</item>
        ///     <item>树节点（TreeItem）</item>
        /// </list>
        /// </para>
        /// <para>Example:</para>
        /// <example>
        /// <code>
        /// // 选择列表中的某一项
        /// WindowsElement listItem = element.FindElementByName("项的名称");
        /// listItem.Click(); // 或者使用 Select()
        /// 
        /// // 检查项是否被选中
        /// string isSelected = listItem.GetAttribute("SelectionItem.IsSelected");
        /// </code>
        /// </example>
        /// </summary>
        public bool HasSelectionItemPattern => bool.Parse(GetAttribute("HasSelectionItemPattern"));

        /// <summary>
        /// 是否支援 SelectionPattern（选择模式）
        /// <para>說明：</para>
        /// <para>
        /// <list type="bullet">
        ///     <item>用途：用于包含可选择子项的控件，如列表框、组合框。</item>
        ///     <item>
        ///         <description>属性：</description>
        ///         <list type="table">
        ///             <item>CanSelectMultiple：是否可以多选。</item>
        ///             <item>IsSelectionRequired：是否必须选择一个子项。</item>
        ///             <item>Selection：当前选中的子项集合。</item>
        ///         </list>
        ///     </item>
        ///     <item>方法：無</item>
        /// </list>
        /// </para>
        /// <para>支持的控件：</para>
        /// <para>
        /// <list type="bullet">
        ///     <item>列表框（ListBox）</item>
        ///     <item>组合框（ComboBox）</item>
        ///     <item>树视图（TreeView）</item>
        /// </list>
        /// </para>
        /// <para>Example:</para>
        /// <example>
        /// <code>
        /// // 獲取當前選中的項
        /// IReadOnlyCollection&lt;WindowsElement&gt; selectedItems = element.FindElementsByXPath(".//ListItem[@IsSelected='True']");
        /// 
        /// // 檢查是否可以多選
        /// string canSelectMultiple = element.GetAttribute("Selection.CanSelectMultiple");
        /// </code>
        /// </example>
        /// </summary>
        public bool CanSelect => bool.Parse(GetAttribute("IsSelectionPatternAvailable"));

        public bool IsSpreadsheetItemPatternAvailable => bool.Parse(GetAttribute("IsSpreadsheetItemPatternAvailable"));

        public bool IsSpreadsheetPatternAvailable => bool.Parse(GetAttribute("IsSpreadsheetPatternAvailable"));

        public bool IsStylesPatternAvailable => bool.Parse(GetAttribute("IsStylesPatternAvailable"));

        public bool IsSynchronizedInputPatternAvailable => bool.Parse(GetAttribute("IsSynchronizedInputPatternAvailable"));

        /// <summary>
        /// TableItemPattern（表格项模式）
        /// <para>說明：</para>
        /// <para>
        /// <list type="bullet">
        ///     <item>用途：用于表格中的单元格，提供对其关联的行头和列头的访问。</item>
        ///     <item>
        ///         <description>属性：</description>
        ///         <list type="table">
        ///             <item>RowHeaderItems：与该单元格关联的行头项。</item>
        ///             <item>ColumnHeaderItems：与该单元格关联的列头项。</item>
        ///         </list>
        ///     </item>
        ///     <item>方法：無</item>
        /// </list>
        /// </para>
        /// <para>支持的控件：</para>
        /// <para>
        /// <list type="bullet">
        ///     <item>表格中的单元格</item>
        /// </list>
        /// </para>
        /// </summary>
        public bool IsTableItemPatternAvailable => bool.Parse(GetAttribute("IsTableItemPatternAvailable"));

        /// <summary>
        /// 是否支援 TablePattern（表格模式）
        /// <para>說明：</para>
        /// <para>
        /// <list type="bullet">
        ///     <item>用途：扩展GridPattern，提供关于表头信息的支持。</item>
        ///     <item>
        ///         <description>属性：</description>
        ///         <list type="table">
        ///             <item>RowHeaders：行头集合。</item>
        ///             <item>ColumnHeaders：列头集合。</item>
        ///             <item>RowOrColumnMajor：指定信息是按行还是按列组织的。</item>
        ///         </list>
        ///     </item>
        ///     <item>方法：無</item>
        /// </list>
        /// </para>
        /// <para>支持的控件：</para>
        /// <para>
        /// <list type="bullet">
        ///     <item>数据网格（DataGrid）</item>
        ///     <item>表格（Table）</item>
        /// </list>
        /// </para>
        /// </summary>
        public bool HasTablePattern => bool.Parse(GetAttribute("IsTablePatternAvailable"));

        public bool IsTextChildPatternAvailable => bool.Parse(GetAttribute("IsTextChildPatternAvailable"));

        public bool IsTextEditPatternAvailable => bool.Parse(GetAttribute("IsTextEditPatternAvailable"));

        /// <summary>
        /// 是否支援 TextPattern（文本模式）
        /// <para>說明：</para>
        /// <para>
        /// <list type="bullet">
        ///     <item>用途：用于包含文本内容的控件，提供对文本的访问和操作。</item>
        ///     <item>
        ///         <description>属性：</description>
        ///         <list type="table">
        ///             <item>DocumentRange：整个文档的文本范围。</item>
        ///             <item>SupportedTextSelection：支持的文本选择类型。</item>
        ///         </list>
        ///     </item>
        ///     <item>
        ///         <description>方法：</description>
        ///         <list type="table">
        ///             <item>GetSelection()：获取当前选定的文本范围。</item>
        ///             <item>GetVisibleRanges()：获取可见的文本范围。</item>
        ///         </list>
        ///     </item>
        /// </list>
        /// </para>
        /// <para>支持的控件：</para>
        /// <para>
        /// <list type="bullet">
        ///     <item>文本框（TextBox）</item>
        ///     <item>文本编辑器</item>
        /// </list>
        /// </para>
        /// </summary>
        public bool HasTextPattern => bool.Parse(GetAttribute("IsTextPatternAvailable"));

        public bool IsTextPattern2Available => bool.Parse(GetAttribute("IsTextPattern2Available"));

        /// <summary>
        /// 是否支援 TogglePattern（切换模式）
        /// <para>說明：</para>
        /// <para>
        /// <list type="bullet">
        ///     <item>用途：用于可以在多个状态之间切换的控件，如复选框、切换按钮。</item>
        ///     <item>
        ///         <description>属性：</description>
        ///         <list type="table">
        ///             <item>ToggleState：当前的切换状态（开、关、不确定）。</item>
        ///         </list>
        ///     </item>
        ///     <item>
        ///         <description>方法：無</description>
        ///         <list type="table">
        ///             <item>Toggle()：在状态之间切换。</item>
        ///         </list>
        ///     </item>
        /// </list>
        /// </para>
        /// <para>支持的控件：</para>
        /// <para>
        /// <list type="bullet">
        ///     <item>复选框（CheckBox）</item>
        ///     <item>切换按钮（ToggleButton）</item>
        /// </list>
        /// </para>
        /// </summary>
        public bool CanToggle => bool.Parse(GetAttribute("IsTogglePatternAvailable"));

        /// <summary>
        /// 是否支援 TransformPattern（变换模式）
        /// <para>說明：</para>
        /// <para>
        /// <list type="bullet">
        ///     <item>用途：用于可以移动、调整大小或旋转的控件，如窗口、浮动面板。</item>
        ///     <item>
        ///         <description>属性：</description>
        ///         <list type="table">
        ///             <item>CanMove：是否可移动。</item>
        ///             <item>CanResize：是否可调整大小。</item>
        ///             <item>CanRotate：是否可旋转。</item>
        ///         </list>
        ///     </item>
        ///     <item>
        ///         <description>方法：</description>
        ///         <list type="table">
        ///             <item>Move(x, y)：移动控件到指定位置。</item>
        ///             <item>Resize(width, height)：调整控件大小。</item>
        ///             <item>Rotate(degrees)：旋转控件。</item>
        ///         </list>
        ///     </item>
        /// </list>
        /// </para>
        /// <para>支持的控件：</para>
        /// <para>
        /// <list type="bullet">
        ///     <item>窗口（Window）</item>
        ///     <item>可调整大小的面板</item>
        /// </list>
        /// </para>
        /// </summary>
        public bool HasTransformPattern => bool.Parse(GetAttribute("IsTransformPatternAvailable"));

        public bool IsTransform2PatternAvailable => bool.Parse(GetAttribute("IsTransform2PatternAvailable"));

        /// <summary>
        /// 是否支援 ValuePattern（值模式）
        /// <para>說明：</para>
        /// <para>
        /// <list type="bullet">
        ///     <item>用途：用于具有可设置或检索值的控件，如文本框、进度条等。</item>
        ///     <item>
        ///         <description>属性：</description>
        ///         <list type="bullet">
        ///             <item>Value：获取或设置控件的当前值。</item>
        ///             <item>IsReadOnly：指示控件的值是否只读。</item>
        ///         </list>
        ///     </item>
        /// </list>
        /// </para>
        /// <para>支持的控件：</para>
        /// <para>
        /// <list type="bullet">
        ///     <item>文本框（Edit）</item>
        ///     <item>进度条（ProgressBar）</item>
        ///     <item>滑块（Slider）</item>
        ///     <item>微调控件（Spinner）</item>
        ///     <item>日期选择器（DatePicker）</item>
        /// </list>
        /// </para>
        /// <para>Example:</para>
        /// <example>
        /// <code>
        /// // 设置文本框的值
        /// WindowsElement textBox = driver.FindElementByAccessibilityId("TextBoxAutomationId");
        /// textBox.Clear();
        /// textBox.SendKeys("输入的文本");
        /// 
        /// // 获取控件的值
        /// string currentValue = textBox.GetAttribute("Value.Value");
        /// Console.WriteLine($"当前值为：{currentValue}");
        /// </code>
        /// </example>
        /// </summary>
        public bool HasValuePattern => bool.Parse(GetAttribute("IsValuePatternAvailable"));

        public bool IsVirtualizedItemPatternAvailable => bool.Parse(GetAttribute("IsVirtualizedItemPatternAvailable"));

        /// <summary>
        /// 是否支援 WindowPattern（窗口模式）
        /// <para>說明：</para>
        /// <para>
        /// <list type="bullet">
        ///     <item>用途：用于顶级窗口(Top Window)，提供窗口的基本操作。</item>
        ///     <item>
        ///         <description>属性：</description>
        ///         <list type="bullet">
        ///             <item>WindowVisualState：窗口的视觉状态（普通、最小化、最大化）。</item>
        ///             <item>WindowInteractionState：窗口的交互状态。</item>
        ///             <item>IsModal：是否为模态窗口。</item>
        ///             <item>IsTopmost：是否为最上层窗口。</item>
        ///         </list>
        ///     </item>
        ///     <item>
        ///         <description>方法：</description>
        ///         <list type="bullet">
        ///             <item>Close()：关闭窗口。</item>
        ///             <item>SetWindowVisualState(state)：设置窗口的视觉状态。</item>
        ///             <item>WaitForInputIdle(timeout)：等待窗口进入空闲状态。</item>
        ///         </list>
        ///     </item>
        /// </list>
        /// </para>
        /// <para>支持的控件：</para>
        /// <para>
        /// <list type="bullet">
        ///     <item>窗口（Window）</item>
        /// </list>
        /// </para>
        /// </summary>
        public bool IsWindow => bool.Parse(GetAttribute("IsWindowPatternAvailable"));

        public bool CanResize => IsWindow && bool.Parse(GetAttribute("Transform.CanResize"));
        public bool CanMove => IsWindow && bool.Parse(GetAttribute("Transform.CanMove"));
        public bool CanRotate => IsWindow && bool.Parse(GetAttribute("Transform.CanRotate"));

        public bool CanMaximize => IsWindow && bool.Parse(GetAttribute("Window.CanMaximize"));
        public bool CanMinimize => IsWindow && bool.Parse(GetAttribute("Window.CanMinimize"));

        public bool IsModal => IsWindow && bool.Parse(GetAttribute("Window.IsModal"));

        public bool IsTopmost => IsWindow && bool.Parse(GetAttribute("Window.IsTopmost"));

        public WindowInteractionState InteractionState => IsWindow ? (WindowInteractionState)Enum.Parse(typeof(WindowInteractionState), GetAttribute("Window.WindowInteractionState")) : WindowInteractionState.None;
        public WindowVisualState VisualState => IsWindow ? (WindowVisualState)Enum.Parse(typeof(WindowVisualState), GetAttribute("Window.WindowVisualState")) : WindowVisualState.None;

        //// PP5 IDE Window Pattern Attributes
        /// Normal Size (800x600):
        //Transform.CanMove:	true
        //Transform.CanResize:	true
        //Transform.CanRotate:	false
        //Window.CanMaximize:	true
        //Window.CanMinimize:	true
        //Window.IsModal:	false
        //Window.IsTopmost:	false
        //Window.WindowInteractionState:	ReadyForUserInteraction(2)
        //Window.WindowVisualState:	Normal(0)

        /// Maximized Size (1920x1040, according to current screen resolution):
        //Transform.CanMove:	false
        //Transform.CanResize:	false
        //Transform.CanRotate:	false
        //Window.CanMaximize:	true
        //Window.CanMinimize:	true
        //Window.IsModal:	false
        //Window.IsTopmost:	false
        //Window.WindowInteractionState:	ReadyForUserInteraction(2)
        //Window.WindowVisualState:	Maximized(1)

        //// TI Editor, "About" Dialog Pattern Attributes
        /// Normal Size (498x449):
        //Transform.CanMove:	true
        //Transform.CanResize:	false
        //Transform.CanRotate:	false
        //Window.CanMaximize:	false
        //Window.CanMinimize:	false
        //Window.IsModal:	true
        //Window.IsTopmost:	false
        //Window.WindowInteractionState:	ReadyForUserInteraction(2)
        //Window.WindowVisualState:	Normal(0)

        //// TI Editor, "OverAllCheck" error Dialog Pattern Attributes
        /// Normal Size (295x200):
        /// AutomationId:	"MessageBoxExDialog"
        //Transform.CanMove:	true
        //Transform.CanResize:	false
        //Transform.CanRotate:	false
        //Window.CanMaximize:	false
        //Window.CanMinimize:	false
        //Window.IsModal:	true
        //Window.IsTopmost:	true
        //Window.WindowInteractionState:	ReadyForUserInteraction(2)
        //Window.WindowVisualState:	Normal(0)

        //// TI Editor, Restore Default Docking Pattern Attributes
        /// Normal Size (378x200):
        // AutomationId:	"MessageBoxExDialog"
        //Transform.CanMove:	true
        //Transform.CanResize:	false
        //Transform.CanRotate:	false
        //Window.CanMaximize:	false
        //Window.CanMinimize:	false
        //Window.IsModal:	true
        //Window.IsTopmost:	true
        //Window.WindowInteractionState:	ReadyForUserInteraction(2)
        //Window.WindowVisualState:	Normal(0)

        public bool IsCustomNavigationPatternAvailable => bool.Parse(GetAttribute("IsCustomNavigationPatternAvailable"));

        public bool IsSelectionPattern2Available => bool.Parse(GetAttribute("IsSelectionPattern2Available"));

        public float? RangeValue
        {
            get
            {
                return HasRangeValuePattern ? float.Parse(GetAttribute("RangeValue.Value")) : null;
            }
        }

        public float? RangeValueMax
        {
            get
            {
                return HasRangeValuePattern ? float.Parse(GetAttribute("RangeValue.Maximum")) : null;
            }
        }

        public float? RangeValueMin
        {
            get
            {
                return HasRangeValuePattern ? float.Parse(GetAttribute("RangeValue.Minimum")) : null;
            }
        }

        #endregion

        #region Check Element control type
        public bool IsTextBox => this.ControlType == ElementControlType.TextBox;
        public bool IsCheckBox => this.ControlType == ElementControlType.CheckBox;
        public bool IsTextBlock => this.ControlType == ElementControlType.TextBlock;
        public bool IsButton => this.ControlType == ElementControlType.Button;
        public bool IsRadioButton => this.ControlType == ElementControlType.RadioButton;
        public bool IsComboBox => this.ControlType == ElementControlType.ComboBox;
        public bool IsProgressBar => this.ControlType == ElementControlType.ProgressBar;
        public bool IsTreeView => this.ControlType == ElementControlType.TreeView;
        public bool IsImage => this.ControlType == ElementControlType.Image;
        public bool IsGridCell => IsCustom && this.HasGridItemPattern;
        public bool IsCustom => this.ControlType == ElementControlType.Custom;

        #endregion

        #region Actions
        public new IElement FindElement(By by) => base.FindElement(by).ConvertToElement();

        public IElement FindElement(ByAutomationIdOrName by) => by.FindElement(this).ConvertToElement();

        public new ReadOnlyCollection<IElement> FindElements(By by)
            => base.FindElements(by).ConvertToElements();

        public string GetText()
        {
            if (this == null)
                throw new ArgumentNullException(this.GetType().ToString(), "The element is null!");

            if (IsTreeView)
                return this.AutomationId;

            if (this.GetFirstTextElement() != null)
                return this.GetFirstTextContent();
            else if (this.HasValuePattern)
                return this.Value;
            else if (this.HasTextPattern)
                return this.Text;
            else if (this.HasRangeValuePattern)
                return this.GetAttribute("RangeValue.Value");
            else if (!this.Name.IsNullOrEmpty())
                return this.Name;
            else if (!this.AutomationId.IsNullOrEmpty())
                return this.AutomationId;
            else if (!this.ClassName.IsNullOrEmpty())
                return this.ClassName;
            return this.GetType().ToString();
        }

        public void LeftClick()
        {
            if (!this.Enabled)
            {
                Logger.LogMessage($"Target is disabled, can not perform LeftClick action on {this.ControlType.ToString()} \"{this.GetText()}\"");
                return;
            }
            Logger.LogMessage($"LeftClick on {this.ControlType.ToString()} \"{this.GetText()}\"");
            new Actions(base.WrappedDriver).Click(this).Perform();
        }

        public void DoubleClick()
        {
            if (!this.Enabled)
            {
                Logger.LogMessage($"Target is disabled, can not perform LeftDoubleClick action on {this.ControlType.ToString()} \"{this.GetText()}\"");
                return;
            }
            Logger.LogMessage($"LeftDoubleClick on {this.ControlType.ToString()} \"{this.GetText()}\"");
            new Actions(base.WrappedDriver).DoubleClick(this).Perform();
        }

        public virtual void SendText(string keysToSend)
        {
            if (keysToSend.IsNullOrEmpty())
            {
                Logger.LogMessage("Empty keysToSend passed in, skip the SendText action on {0}!", this.ToString());
                return;
            }

            //if (IsTextBox && !this.Value.IsHexString())
                base.Clear();
            //else if (IsComboBox) { }
            //else
            //{
            //    this.SelectAllContent();
            //    AutoUIActionHelper.Press(Keys.Backspace);
            //}
            Logger.LogMessage($"Clear text in {this.ControlType.ToString()} \"{this.GetText()}\"");

            if (this.Value.IsHexString())
                this.SendKeys(keysToSend);
            else
                new Actions(base.WrappedDriver).SendKeys(this, keysToSend).Perform();
            Logger.LogMessage($"KeyboardInput \"{keysToSend}\" in {this.ControlType.ToString()} \"{this.GetText()}\"");
        }

        public void MoveToElement()
        {
            new Actions(base.WrappedDriver).MoveToElement(this).Perform();
        }

        public void MoveToElement(MoveToElementOffsetStartingPoint startingPoint, int subOffsetX = 0, int subOffsetY = 0)
        {
            int offsetX = 0, offsetY = 0;
            this.GetOffsetFromElementCenter(startingPoint, ref offsetX, ref offsetY);
            new Actions(base.WrappedDriver).MoveToElement(this, offsetX + subOffsetX, offsetY + subOffsetY, MoveToElementOffsetOrigin.Center)
                                           .Perform();
        }

        public void DragAndDropToOffset(MoveToElementOffsetStartingPoint startingPoint, int srcOffsetX, int srcOffsetY, int destOffsetX, int destOffsetY)
        {
            this.MoveToElement(startingPoint, srcOffsetX, srcOffsetY);
            new Actions(base.WrappedDriver).ClickAndHold()
                                           .MoveByOffset(destOffsetX, destOffsetY)
                                           .Release().Perform();
        }

        public void DragAndDropToOffset(DragAndDropInfo dragAndDropInfo)
        {
            this.MoveToElement(dragAndDropInfo.startingPointType, dragAndDropInfo.srcOffsetX, dragAndDropInfo.srcOffsetY);
            new Actions(base.WrappedDriver).ClickAndHold()
                                           .MoveByOffset(dragAndDropInfo.destOffsetX, dragAndDropInfo.destOffsetY)
                                           .Release().Perform();
        }
        #endregion

        #region Record Errors
        private Hashtable errorMessages = new Hashtable();
        public Hashtable ErrorMessages 
        { 
            get { return errorMessages; } 
            set { errorMessages = value; }
        }
        #endregion

        public override string ToString()
        {
            return string.Format(CultureInfo.InvariantCulture, "Element ({0}, ControlType:{1})", this.GetText(), this.LocalizedControlType);
        }

        public bool IsSameElement(IElement elementToCompare)
        {
            return this.Equals(elementToCompare);
        }
    }
}
