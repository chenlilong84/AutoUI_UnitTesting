using System.Collections;
using System.Collections.ObjectModel;
using Chroma.UnitTest.Common;
using Microsoft.VisualStudio.TestTools.UnitTesting.Logging;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Service;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Appium.Enums;
using OpenQA.Selenium.Appium.iOS;
using System.Drawing;
using System.Globalization;
using System.Windows.Input;
using System.ComponentModel;
using OpenQA.Selenium.Support.Events;
using System.Threading;
using System.Linq;
using System.Drawing.Printing;
using MyCursor;

namespace PP5AutoUITests
{
    public class PP5DataGrid : PP5Element
    {
        private readonly PP5Element Element;

        #region Constructor/Destructor/Init
        /// <summary>
        /// Initializes a new instance of the PP5DataGrid class.
        /// </summary>
        /// <param name="parent">Driver in use.</param>
        /// <param name="id">ID of the element.</param>
        public PP5DataGrid(WindowsElement element) : base(element) 
        {
            //LastRowNo = -1;
            //SelectedColumnNo = -1;
            RefreshSelectedCell();
        }
        public PP5DataGrid(PP5Element element) : base((WindowsElement)element) 
        {
            //LastRowNo = -1;
            //SelectedColumnNo = -1;
            Element = element;
            InitProperties();
            RefreshSelectedCell();
            Element.PropertyChanged += PP5AttributeChanged;
        }

        public void InitProperties()
        {
            _displayedColumnCount = null;
            _displayedRowCount = null;
            TiVariables = new VariableRowInfoSets();
        }

        public void Dispose()
        {
            //if (Element is INotifyPropertyChanged notifyPropertyChanged)
            //{
                Element.PropertyChanged -= PP5AttributeChanged;
            //}
        }
        #endregion

        #region Events
        public void PP5AttributeChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(Element.ScrollBarVerticalPosition) || e.PropertyName == nameof(Element.ScrollBarHorizontalPosition) || 
                e.PropertyName == nameof(Element.Height) || e.PropertyName == nameof(Element.Width))
            {
                // Update the DisplayedColumnCount or DisplayedRowCount
                UpdateDisplayedCounts();
            }
        }

        private void UpdateDisplayedCounts()
        {
            DisplayedColumnCount = this.GetDisplayedColumnCount();
            DisplayedRowCount = this.GetDisplayedRowCount();
        }
        #endregion

        #region Selected Cell Info
        private SelectedCellInfo _selectedCellInfo;

        public SelectedCellInfo SelectedCellInfo => _selectedCellInfo ??= new SelectedCellInfo(() => this.GetSelectedCell(), this);
        //{
        //    IElement selectedCell = null;
        //    //if (this.GetSelectedRow() != null)
        //        selectedCell =  this.GetSelectedRow()?.GetCellElementsOfRow().FirstOrDefault(c => c.Selected);
        //    //else
        //    //    selectedCell =  this.GetDataItems().FirstOrDefault().GetCellElementsOfRow().FirstOrDefault(c => c.Selected);

        //    //selectedCell ??= this.PerformGetElement("ByCell[1,1]");
        //    return selectedCell;
        //});

        private int DisplayedRowCountTemp;
        private int _selectedRowNo;
        public int GetSelectedRowNo()
        {
            if (_selectedCellInfo.RowNo == -1)
            {
                SelectedRowNo = -1;
                return SelectedRowNo;
            }
                
            if (DisplayedRowCountTemp != DisplayedRowCount)
                DisplayedRowCountTemp = DisplayedRowCount;
            _selectedRowNo = TotalRowCount > DisplayedRowCountTemp ? _selectedCellInfo.RowNo - (TotalRowCount - DisplayedRowCountTemp) : _selectedCellInfo.RowNo;
            SelectedRowNo = _selectedRowNo;
            return _selectedRowNo;
        }

        private int _selectedColumnNo;
        public int GetSelectedColumnNo()
        {
            if (_selectedCellInfo.ColumnNo == -1)
            {
                SelectedColumnNo = -1;
                return SelectedColumnNo;
            }
            SelectedColumnNo = _selectedColumnNo = _selectedCellInfo.ColumnNo;
            return _selectedColumnNo;
        }

        public void RefreshSelectedCell()
        {
            SelectedCellInfo.SelectedCell = null; // 重置，允許重新計算
            //GetSelectedRowNo();
            //GetSelectedColumnNo();
            SelectedCellInfo.SelectedCell = this.GetSelectedCell();
        }
        #endregion

        #region Row / Column properties
        private int lastRowNo;
        public int LastRowNo
        {
            get 
            {
                if (CanScrollVertically && !IsScrollPositionAtBottom)
                    ScrollToBottom();
                lastRowNo = CanScrollVertically ? DisplayedRowCount - 1 : TotalRowCount - 1;
                return lastRowNo;
            }
        }
        private int lastColumnNo;
        public int LastColumnNo
        {
            get
            {
                if (CanScrollHorizontally && IsScrollPositionAtRightMost)
                    ScrollToRight();
                lastColumnNo = CanScrollHorizontally ? DisplayedColumnCount : TotalColumnCount;
                return lastColumnNo;
            }
        }

        public string cellValueCache { get; set; }

        public int SelectedRowNo { get; set; }
        public int SelectedColumnNo { get; set; }

        public int TotalRowCount => this.GetRowCount();
        public int TotalColumnCount => this.GetColumnCount();

        private int? _displayedRowCount;
        public int DisplayedRowCount
        {
            get
            {
                if (_displayedRowCount == null)
                {
                    _displayedRowCount = this.GetDisplayedRowCount();
                }
                return _displayedRowCount.Value;
            }
            private set => _displayedRowCount = value;
        }
        private int? _displayedColumnCount;
        public int DisplayedColumnCount
        {
            get
            {
                if (_displayedColumnCount == null)
                {
                    _displayedColumnCount = this.GetDisplayedColumnCount();
                }
                return _displayedColumnCount.Value;
            }
            private set => _displayedColumnCount = value;
        }

        /// <summary>
        /// Scroll position at bottom
        /// </summary>
        public bool IsScrollPositionAtBottom
        {
            get
            {
                var scrollValue = Element.ScrollBarVerticalPosition;
                if (int.TryParse(scrollValue, out int scrollPercent))
                {
                    return scrollPercent == 100;
                }
                return false;
            }
        }

        /// <summary>
        /// Scroll position at right most
        /// </summary>
        public bool IsScrollPositionAtRightMost
        {
            get
            {
                var scrollValue = Element.ScrollBarHorizontalPosition;
                if (int.TryParse(scrollValue, out int scrollPercent))
                {
                    return scrollPercent == 100;
                }
                return false;
            }
        }
        #endregion

        #region For TI Editor
        public VariableRowInfoSets TiVariables { get; set; }
        #endregion

        #region DataGrid Actions
        public void ScrollToBottom()
        {
            //this.GetCellBy(1, 1).LeftClick();
            //while (!IsScrollPositionAtBottom)
            //{
            //    AutoUIActionHelper.Press(Keys.PageDown);
            //}
            int colHeaderHeight = this.PerformGetElement("/ByName[DataPanel]").PointAtTopCenter.Y - this.PointAtTopCenter.Y;
            int vertScrollBarAreaHeight = this.Height - 12 * 2 - colHeaderHeight;
            float vertScrollBarHeight = vertScrollBarAreaHeight * float.Parse(this.GetAttribute("Scroll.VerticalViewSize")) / 100;
            int vertScrollBarAvalSpace = Convert.ToInt32(vertScrollBarAreaHeight - vertScrollBarHeight);
            this.DragAndDropToOffset(MoveToElementOffsetStartingPoint.TopRight, -6, 12 + vertScrollBarAreaHeight / 2, 0, vertScrollBarAvalSpace);
            //Thread.Sleep(100);
            DisplayedRowCount = this.GetDisplayedRowCount();
        }

        public void ScrollToSpecificVerticalPosition(int rowNo)
        {
            int colHeaderHeight = this.PerformGetElement("/ByName[DataPanel]").PointAtTopCenter.Y - this.PointAtTopCenter.Y;
            int vertScrollBarAreaHeight = this.Height - 12 * 2 - colHeaderHeight;
            float vertScrollBarHeight = vertScrollBarAreaHeight * float.Parse(this.GetAttribute("Scroll.VerticalViewSize")) / 100;
            int vertScrollBarAvalSpace = Convert.ToInt32(vertScrollBarAreaHeight - vertScrollBarHeight);

            float rowHeight = vertScrollBarAreaHeight / TotalRowCount;

            this.DragAndDropToOffset(MoveToElementOffsetStartingPoint.TopRight, -6, 12 + vertScrollBarAreaHeight / 2, 0, (int)(rowHeight * (rowNo - 1)));
            DisplayedRowCount = this.GetDisplayedRowCount();
        }

        public void ScrollToRight()
        {
            int hrztScrollBarAreaWidth = GetHorizontalScrollBarAreaWidth();
            int hrztScrollBarAvalSpace = GetHorizontalScrollBarAvailableSpace();

            this.DragAndDropToOffset(MoveToElementOffsetStartingPoint.BottomLeft, 12 + hrztScrollBarAreaWidth / 2, -6, hrztScrollBarAvalSpace, 0);
            ScrollBarHorizontalPosition = this.GetAttribute("Scroll.HorizontalScrollPercent");
            //Thread.Sleep(100);
            DisplayedColumnCount = this.GetDisplayedColumnCount();
        }

        public void ScrollToSpecificColumn(string ColumnName)
        {
            var headerElements = this.GetDataTableHeaderElements();
            if (headerElements.Count() == 0 || !headerElements.Select(e => e.GetText()).Contains(ColumnName))
                throw new ArgumentException($"No headers found, or column name \"{ColumnName}\" not existed!");
            
            int SpecificColumnIndex = headerElements.Select(e => e.GetText()).ToList().IndexOf(ColumnName);

            int singleMovingStep = Convert.ToInt32(GetHorizontalScrollBarAvailableSpace() / 3);
            while (!headerElements.ElementAt(SpecificColumnIndex).IsWithinScreen)
            {
                int hrztScrollBarAreaWidth = GetHorizontalScrollBarAreaWidth();
                this.DragAndDropToOffset(MoveToElementOffsetStartingPoint.BottomLeft, 12 + hrztScrollBarAreaWidth / 2, -6, singleMovingStep, 0);
            }
        }

        private int GetHorizontalScrollBarAvailableSpace()
        {
            int hrztScrollBarAreaWidth = GetHorizontalScrollBarAreaWidth();
            float hrztScrollBarWidth = hrztScrollBarAreaWidth * float.Parse(this.GetAttribute("Scroll.HorizontalViewSize")) / 100;
            return Convert.ToInt32(hrztScrollBarAreaWidth - hrztScrollBarWidth);
        }

        private int GetHorizontalScrollBarAreaWidth()
        {
            return this.Width - 12 * 2;
        }
        #endregion

        public override string ToString()
        {
            return string.Format(CultureInfo.InvariantCulture, "DataGrid Element ({0}, size: {1} x {2})", this.GetText(), TotalRowCount, TotalColumnCount);
        }
    }

    public class SelectedCellInfo
    {
        private PP5Element _selectedCell;
        public PP5DataGrid pp5DataGrid;
        private List<string> headers = new List<string>();

        public SelectedCellInfo(Func<IElement> getSelectedCell) 
        {
            _selectedCell = (PP5Element)getSelectedCell?.Invoke();
        }

        public SelectedCellInfo(Func<IElement> getSelectedCell, PP5DataGrid _pp5DataGrid) : this(getSelectedCell)
        {
            this.pp5DataGrid = _pp5DataGrid;
            headers = pp5DataGrid?.GetDataGridHeaders();
        }

        public PP5Element SelectedCell 
        { 
            get { return _selectedCell; } 
            set { _selectedCell = value; } 
        }

        public int RowNo => _selectedCell?.GetRowIndexOfCellElement() + 1 ?? -1;

        public int ColumnNo => _selectedCell?.GetColumnIndexOfCellElement() + 1 ?? -1;

        public string ColumnName => headers?[ColumnNo - 1];

        public void SendText(string keysToSend)
        {
            _selectedCell.SendText(keysToSend);
            if (_selectedCell.IsGridCell)
            {
                if (DllHelper.GetCursor() == Cursor.IBeam)
                    _selectedCell.MoveToElementAndClick(0, _selectedCell.Height, MoveToElementOffsetOrigin.Center);
                else
                    AutoUIActionHelper.Press(Keys.Enter);
            }
        }
    }
}
