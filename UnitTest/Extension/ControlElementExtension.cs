﻿using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Chroma.UnitTest.Common;
using Microsoft.VisualStudio.TestTools.UnitTesting.Logging;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Support.UI;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Tab;
using static PP5AutoUITests.AutoUIActionHelper;
using static PP5AutoUITests.ElementFinder;
using static PP5AutoUITests.ThreadHelper;
using Keys = OpenQA.Selenium.Keys;

namespace PP5AutoUITests
{
    public static class ControlElementExtension
    {
        //static IElement toolbarItem = null;

        #region MenuList

        //static IWebDriver driver = Executor.GetInstance().GetCurrentDriver();

        //public static bool IsMenuItemEnabled(this IWebElement window, params string[] itemNames)
        //{
        //    return window.GetMenuItem(itemNames).Enabled;
        //}

        //public static IWebElement GetMenuItem(this IWebElement window, params string[] itemNames)
        //{
        //    IWebElement menu = window.GetElementFromWebElement(By.ClassName("Menu"));

        //    foreach (string itemName in itemNames)
        //    {
        //        Console.WriteLine($"LeftClick on Text \"{itemName}\"");
        //        var menuItemText = menu.GetElementFromWebElement(By.XPath($".//Text[@Name='{itemName}']"));
        //        var subMenu = CurrentDriver.GetParentElement(menuItemText);

        //        // Use attribute: 'IsExpandCollapsePatternAvailable' to check if menuitem can be expanded
        //        if (subMenu.GetAttribute("IsExpandCollapsePatternAvailable") == false.ToString())
        //        {
        //            return subMenu;
        //        }

        //        menuItemText.LeftClick();
        //        menu = subMenu;
        //    }
        //    return menu;
        //}

        public static bool MenuSelect(this IElement window, params string[] itemNames)
        {
            IElement menu = window.PerformGetElement("/ByClass[Menu]");
            IElement menuItem;
            bool isSuccess = true;

            foreach (string itemName in itemNames)
            {
                //Logger.LogMessage($"menuItems count: {menuItem.GetChildElementsCount()}");
                menuItem = menu.PerformClick($"/ByName[{itemName}]", ClickType.LeftClick);
                if (menuItem == null)
                    menuItem = window.PerformClick($"/ByClass[Popup]/ByName[{itemName}]", ClickType.LeftClick);
                if (!menuItem.Enabled)
                {
                    if (menuItem != null)
                        Press(Keys.Escape);
                    isSuccess = false;
                    break;
                }
                //isSuccess = menuItem.Enabled & isSuccess;
            }
            return isSuccess;
        }

        /// <summary>
        /// Menu selection helper function
        /// </summary>
        /// <param name="itemNames">name of the menuitem to be seleted in stirng</param>
        public static bool MenuSelect(params string[] itemNames)
        {
            var driver = Executor.GetInstance().GetCurrentDriver();
            //IElement menu = driver.GetElement<PP5Driver,IElement>(PP5By.ClassName("Menu"));
            IElement menu = driver.PerformGetElement("/ByClass[Menu]");
            IElement menuItem;
            bool isSuccess = true;
            
            foreach (string itemName in itemNames)
            {
                menuItem = menu.PerformClick($"/ByName[{itemName}]", ClickType.LeftClick);
                if (menuItem == null)
                    menuItem = driver.PerformClick($"/ByClass[Popup]/ByName[{itemName}]", ClickType.LeftClick);
                if (!menuItem.Enabled)
                {
                    if (menuItem != null)
                        Press(Keys.Escape);
                    isSuccess = false;
                    break;
                }
                //isSuccess = menuItem.Enabled & isSuccess;
            }
            return isSuccess;
        }

        public static IEnumerable<string> GetSubMenuListItemNames(string itemName)
        {
            //IElement menu = Executor.GetInstance().GetCurrentDriver().GetExtendedElement(PP5By.ClassName("Menu"));
            IElement menu = Executor.GetInstance().GetCurrentDriver().PerformGetElement("/ByClass[Menu]");
            IEnumerable<string> MenuListItemsNames;

            //Console.WriteLine($"LeftClick on Text \"{itemName}\"");
            //IWebElement subMenu = menu.GetElementFromWebElement(By.XPath($".//Text[@Name='{itemName}']/.."));
            //IElement subMenu = menu.GetExtendedElement(PP5By.Name(itemName));
            IElement subMenu = menu.PerformGetElement($"/ByName[{itemName}]");
            subMenu.LeftClick();
            MenuListItemsNames = subMenu.GetChildElements().Select(e => e.Text);
            subMenu.LeftClick();
            return MenuListItemsNames;
        }

        public static IEnumerable<string> GetMainMenuListItemNames()
        {
            //IElement menu = Executor.GetInstance().GetCurrentDriver().GetExtendedElement(PP5By.ClassName("Menu"));
            IElement menu = Executor.GetInstance().GetCurrentDriver().PerformGetElement("/ByClass[Menu]");
            return menu.GetMenuItems().Select(e => e.Text);
        }

        public static IEnumerable<string> GetSubMenuListItemNames(this IElement element, string itemName)
        {
            //IElement menu = Executor.GetInstance().GetCurrentDriver().GetExtendedElement(PP5By.ClassName("Menu"));
            //IElement menu = Executor.GetInstance().GetCurrentDriver().PerformGetElement("/ByClass[Menu]");
            IElement menu = element;
            IElement subMenu = null;
            IEnumerable<string> MenuListItemsNames;

            if (element.ControlType != ElementControlType.Menu)
                menu = element.PerformGetElement("/ByClass[Menu]");

            subMenu = menu.PerformGetElement($"/ByName[{itemName}]");

            subMenu.LeftClick();
            MenuListItemsNames = subMenu.GetMenuItems().Select(e => e.Text);
            subMenu.LeftClick();

            return MenuListItemsNames;
        }

        public static IEnumerable<string> GetMainMenuListItemNames(this IElement element)
        {
            //IElement menu = Executor.GetInstance().GetCurrentDriver().GetExtendedElement(PP5By.ClassName("Menu"));
            //IElement menu = Executor.GetInstance().GetCurrentDriver().PerformGetElement("/ByClass[Menu]");
            if (element.ControlType == ElementControlType.Menu)
                return element.GetMenuItems().Select(e => e.Text);
            else
                return element.PerformGetElement("/ByClass[Menu]").GetMenuItems().Select(e => e.Text);
        }

        #endregion

        #region ComboBox

        public static void GetComboBoxItems(this IElement comboBox, out ReadOnlyCollection<IElement> cmbItems)
        {
            //if (comboBox.TagName == ElementControlType.ComboBox.GetDescription())
            //    return CurrentDriver.GetElements(By.ClassName("ComboBoxItem"));
            //else if (comboBox.TagName == ElementControlType.ListBox.GetDescription())
            //    return CurrentDriver.GetElements(By.ClassName("ListBoxItem"));
            //else
            //    return null;
            
            //comboBox.LeftClick();
            //cmbItems = CurrentDriver.GetElementFromWebElement(By.ClassName("Popup"))
            //                        .GetElements(By.ClassName("ListBoxItem"));
            cmbItems = comboBox.PerformGetElements("/ByClass[ListBoxItem]").AsReadOnly();
        }

        public static string[] GetComboBoxItemNames(this IElement comboBox)
        {
            comboBox.GetComboBoxItems(out ReadOnlyCollection<IElement> cmbItems);
            return cmbItems.Select(e => e.Text).ToArray();
        }

        public static int GetComboBoxItemCount(this IElement comboBox)
        {
            //comboBox.LeftClick();
            //cmbItems = CurrentDriver.GetElementFromWebElement(By.ClassName("Popup"))
            //                        .GetElements(By.ClassName("ListBoxItem"));
            return comboBox.GetExtendedElements(PP5By.ClassName("ListBoxItem")).Count;
        }

        public static void SelectComboBoxItemByName(this IElement comboBox, string name, bool supportKeyInputSearch = true)
        {
            if (comboBox.CheckComboBoxHasItemByName(name, out IElement cmbItem))
            {
                if (supportKeyInputSearch)
                {
                    comboBox.SendSingleKeys(name);
                }
                else
                {
                    cmbItem.LeftClick();
                }
            }
        }

        public static void SelectComboBoxItemByName2(this IElement comboBox, string name)
        {
            comboBox.DoubleClickWithDelay(10);
            comboBox.SendText(name);
            //if (comboBox.GetAttribute("IsExpandCollapsePatternAvailable") == bool.FalseString)      // ListBox
            //{
            //    // Case 1: IsKeyboardFocusable=false, IsValuePatternAvailable=false
            //    if (comboBox.GetAttribute("IsKeyboardFocusable") == bool.FalseString)
            //}
            //else                                                                                    // ComboBox
            //{
            //    // Case 2: IsKeyboardFocusable=true, IsValuePatternAvailable=true
            //    // Case 3: IsKeyboardFocusable=true, IsValuePatternAvailable=true
            //}

            //if (supportKeyInputSearch)
            //{
            //    comboBox.SendSingleKeys(name);
            //}
            //else
            //{
            //    cmbItem.LeftClick();
            //}
        }

        public static void SelectComboBoxItemByIndex(this IElement comboBox, int index, bool supportKeyInputSearch = true)
        {
            comboBox.GetComboBoxItems(out ReadOnlyCollection<IElement> cmbItems);
            if (cmbItems.Count() >= index + 1)
            {
                string name = cmbItems[index].Text;
                if (supportKeyInputSearch)
                {
                    SendComboKeys(name, Keys.Enter);
                }
                else
                {
                    cmbItems[index].LeftClick();
                }
            }
        }

        public static IElement SelectComboBoxItemByIndex2(this IElement comboBox, int index)
        {
            comboBox.GetComboBoxItems(out ReadOnlyCollection<IElement> cmbItems);
            if (cmbItems.Count() >= index + 1)
            {
                comboBox.SelectComboBoxItemByName2(cmbItems[index].Text);
            }
            return cmbItems[index];
        }

        public static bool CheckComboBoxHasItemByName(this IElement comboBox, string name, out IElement cmbItem)
        {
            //comboBox.GetComboBoxItems(out ReadOnlyCollection<IWebElement> cmbItems);
            //cmbItem = cmbItems.FirstOrDefault(item => item.GetFirstTextContent() == name);
            // 20240820, Adam, directly get the ListBoxItem by name for improving speed
            cmbItem = comboBox.GetListBoxItemElement(name);
            return cmbItem != null;
        }

        #endregion

        #region DataGrid

        public static void SelectDataGridItemByRowIndex(this IElement element, int rowIdx)
        {
            //window.GetElements(By.XPath(".//DataItem"))[rowIdx].LeftClick();

            //ReadOnlyCollection<IWebElement> rows = window.GetChildElementsOfControlType(ElementControlType.DataItem);
            ReadOnlyCollection<IElement> rows = element.GetDataItems();

            if (rowIdx >= rows.Count || rowIdx < -1)
                throw new ArgumentException("wrong row index!");

            if (rowIdx == -1)
                rows[rows.Count - 1].LeftClick();
            else
                rows[rowIdx].LeftClick();
        }

        public static List<string> GetDataGridHeaders(this IElement dataGridElement)
        {
            return dataGridElement.GetDataTableHeaders().ToList();
        }

        public static List<string> GetSingleColumnValues(this IElement element, int colNo)
        {
            IEnumerable<IElement> column = element.GetColumnBy(colNo);
            List<string> columnValues = new List<string>();

            if (column == null)
                return null;
            else
            {
                columnValues = column.Select(c => c.GetText()).ToList();
                //columnValues.RemoveAll(s => s.IsNullOrEmpty());
                return columnValues;
            }
        }

        #endregion

        #region TreeView

        public static bool ExpandTreeView(this IElement treeviewElement)
        {
            // Use attribute: "ExpandCollapse.ExpandCollapseState" to check the expand/collapse state, where: Expanded (1), Collapsed (0)
            // if the window is the tree item leaf node: LeafNode (3), it's not expandable
            if (treeviewElement.isElementAtLeafNode())
                return false;
            else if (treeviewElement.isElementCollapsed())
                treeviewElement.GetFirstTextElement().DoubleClick();

            return WaitUntil(() => treeviewElement.isElementExpanded());
        }

        public static bool SelectTreeViewItem(this IElement treeviewElement, params string[] labels)
        {
            IElement tve = treeviewElement;
            IElement tvieTmp;
            bool atLeafNode = false;

            foreach (string label in labels)
            {
                //tvieTmp = tve.GetTreeViewItemElement(label);
                tvieTmp = tve.GetTreeViewItems().FirstOrDefault(tvi => tvi.GetText() == label);
                atLeafNode = !tvieTmp.ExpandTreeView();
                
                if (atLeafNode)
                {
                    // Select on the tree node
                    if (!tvieTmp.Selected)
                        tvieTmp.LeftClick();
                    break;
                }
                else
                    tve = tvieTmp;
            }
            return atLeafNode;
        }

        public static bool AddTreeViewItem(this IElement treeviewElement, params string[] labels)
        {
            IElement tve = treeviewElement;
            IElement tvieTmp;
            bool atLeafNode = false;

            foreach (string label in labels)
            {
                //Logger.LogMessage($"tve is: {tve}");
                //tvieTmp = tve.GetTreeViewItemElement(label);
                //var tvitems = tve.GetTreeViewItems();
                //Logger.LogMessage($"tv items count: {tvitems.Count}");
                //foreach (var tvitem in tvitems)
                //    Logger.LogMessage($"tv item: {tvitem}");
                tvieTmp = tve.GetTreeViewItems().FirstOrDefault(tvi => tvi.GetText() == label);

                if (tvieTmp == null)
                {
                    Logger.LogMessage("Label not found in tree view item!");
                    return false;
                }    

                //Logger.LogMessage($"tvieTmp is: {tvieTmp}");
                atLeafNode = !tvieTmp.ExpandTreeView();

                if (atLeafNode)
                {
                    // Select on the tree node
                    if (!tvieTmp.Selected)
                        tvieTmp.DoubleClick();
                    break;
                }
                else
                    tve = tvieTmp;
            }
            return atLeafNode;
        }

        public static IElement GetSelectedTreeViewItem(this IElement treeviewElement)
        {
            IElement tve = treeviewElement;
            IElement tvie;
            tvie = tve.GetFirstTreeViewElement();
            return tvie.GetTreeViewItems().FirstOrDefault(e => e.Selected);
        }

        public static bool CheckTreeViewItemSelected(this IElement treeviewElement, params string[] labels)
        {
            IElement tve = treeviewElement;
            IElement tvieTmp;
            bool atLeafNode = false;

            foreach (string label in labels)
            {
                tvieTmp = tve.GetTreeViewItems().FirstOrDefault(tvi => tvi.GetText() == label);

                if (tvieTmp == null)
                {
                    Logger.LogMessage("Label not found in tree view item!");
                    return false;
                }

                //Logger.LogMessage($"tvieTmp is: {tvieTmp}");
                atLeafNode = !tvieTmp.ExpandTreeView();

                if (atLeafNode)
                {
                    // Check if the leaf tree node is selected
                    return tvieTmp.Selected;
                }
                else
                    tve = tvieTmp;
            }
            return atLeafNode;
        }

        #endregion

        #region CheckBox

        public static void TickCheckBox(this IElement checkBox)
        {
            if (!checkBox.isElementChecked())
                checkBox.LeftClick();
        }

        public static void UnTickCheckBox(this IElement checkBox)
        {
            if (checkBox.isElementChecked())
                checkBox.LeftClick();
        }

        #endregion

        #region TabControl

        public static IElement TabSelect(this IElement tabControl, params object[] tabNamesOrIdxs)
        {
            if (tabNamesOrIdxs == null || tabNamesOrIdxs.Length == 0)
                return null;

            IElement tabControlTemp = tabControl;
            IElement tabItem = null;

            for (int i = 0; i < tabNamesOrIdxs.Length; i++)
            {
                tabItem = tabControlTemp.TabSelect(tabNamesOrIdxs[i]);
                if (i == tabNamesOrIdxs.Length - 1) break;
                //tabControlTemp = tabItem.GetFirstTabControlElement();
                //tabControlTemp = tabItem.GetExtendedElement(PP5By.ClassName("TabControl"));
                tabControlTemp = tabItem.PerformGetElement("/ByClass[TabControl]");
            }

            return tabItem;
        }

        public static IElement TabSelect(this IElement tabControl, params string[] tabNames)
        {
            if (tabNames == null || tabNames.Length == 0)
                return null;

            IElement tabControlTemp = tabControl;
            IElement tabItem = null;

            //tabControlTemp.TabSelect(tabNames.First());

            //if (tabNames.Length > 1)
            //{
            //    string[] tabNamesTmp = tabNames.Skip(1).ToArray();
            //    tabControlTemp = tabItem.GetFirstTabControlElement();
            //    tabItem = tabControlTemp.TabSelect(tabNamesTmp);
            //}

            for (int i = 0; i < tabNames.Length; i++)
            {
                tabItem = tabControlTemp.TabSelect(tabNames[i]);
                if (i == tabNames.Length - 1) break;
                //tabControlTemp = tabItem.GetFirstTabControlElement();
                //tabControlTemp = tabItem.GetExtendedElement(PP5By.ClassName("TabControl"));
                tabControlTemp = tabItem.PerformGetElement("/ByClass[TabControl]");
            }

            return tabItem;
        }

        public static IElement TabSelect(this IElement tabControl, params int[] indeces)
        {
            if (indeces == null || indeces.Length == 0)
                return null;

            IElement tabControlTemp = tabControl;
            IElement tabItem = null;

            /*
            //foreach (int idx in indeces)
            //{
            //    tabItem = tabControlTemp.GetElements(By.ClassName("TabItem"))[idx];
            //    if (!tabItem.Selected)
            //        tabItem.LeftClick();

            //    tabControlTemp = tabItem.GetElementFromWebElement(By.ClassName("TabControl"));
            //}

            //tabItem = tabControlTemp.GetElements(By.ClassName("TabItem"))[indeces.First()];
            //if (!tabItem.Selected && tabItem.isElementVisible())
            //    tabItem.LeftClick();
            */

            for (int i = 0; i < indeces.Length; i++)
            {
                tabItem = tabControlTemp.TabSelect(indeces[i]);
                if (i == indeces.Length - 1) break;
                //tabControlTemp = tabItem.GetExtendedElement(PP5By.ClassName("TabControl"));
                tabControlTemp = tabItem.PerformGetElement("/ByClass[TabControl]");
            }

            return tabItem;
        }

        public static IElement TabSelect(this IElement tabControl, int index)
        {
            //IElement tabItem = tabControl.GetExtendedElements(PP5By.ClassName("TabItem"))[index];
            Func<IElement> getTabItemFunc = () => tabControl.PerformGetElements("/ByClass[TabItem]")[index];
            IElement tabItem = getTabItemFunc();
            if (!tabItem.Selected && tabItem.isElementVisible())
                tabItem.LeftClick();
            return getTabItemFunc();
        }

        public static IElement TabSelect(this IElement tabControl, string tabName)
        {
            //IWebElement tabItem = tabControl.GetTabItemElement(tabName);
            //IElement tabItem = tabControl.GetExtendedElement(PP5By.Name(tabName));
            Func<IElement> getTabItemFunc = () => tabControl.PerformGetElement($"/ByName[{tabName}]");
            IElement tabItem = getTabItemFunc();
            //IElement tabItem = tabControl.PerformGetElement($"/ByName[{tabName}]");
            if (!tabItem.Selected && tabItem.isElementVisible())
                tabItem.LeftClick();
            return getTabItemFunc();
        }

        public static IElement TabSelect(this IElement tabControl, object tabNameOrIdx)
        {
            IElement tabItem = null;
            if (tabNameOrIdx.GetType() == typeof(string))
                tabItem = tabControl.TabSelect(tabNameOrIdx.ToString());
            else if (tabNameOrIdx.GetType() == typeof(int))
                tabItem = tabControl.TabSelect(int.Parse(tabNameOrIdx.ToString()));

            return tabItem;
        }

        #endregion

        #region Toolbar

        /// <summary>
        /// ToolBar selection helper function
        /// </summary>
        /// <param name="itemNames">Index of the ToolBar to be seleted in integer</param>
        public static IElement ToolBarSelect(this IElement window, int index)
        {
            IElement toolbarItem = window.GetToolbarElement((e) => e.isElementVisible()).GetRdoBtnElement(index);
            
            if (toolbarItem.Enabled)
                toolbarItem.LeftClick();

            return toolbarItem;
        }

        /// <summary>
        /// ToolBar selection helper function
        /// </summary>
        /// <param name="itemNames">Index of the ToolBar to be seleted in integer</param>
        public static bool ToolBarGetSelectionState(this IElement window, int index)
        {
            IElement toolbarItem = window.GetToolbarElement((e) => e.isElementVisible()).GetRdoBtnElement(index);
            return toolbarItem.Enabled;
        }

        #endregion

        #region ToolTip

        public static string GetToolTipMessage(this IElement element)
        {
            string tooltopMsg = Executor.GetInstance().SwitchTo(SessionType.Desktop).GetToolTipMessage((PP5Element)element);
            Executor.GetInstance().SwitchTo(SessionType.PP5IDE);
            return tooltopMsg;
        }

        #endregion

        #region Window
        /// <summary>
        /// 水平拖曳元素的方法
        /// </summary>
        /// <param name="window">要拖曳的元素</param>
        /// <param name="offsetX">水平拖曳的距離，正值向右拖曳，負值向左拖曳</param>
        public static void DragHorizontally(this IElement splitter, int offsetX)
        {
            try
            {
                // 1. 確認元素 ControlType 是否為 Thumb (ClassName: GridSplitter)
                if (splitter.ControlType != ElementControlType.Thumb)
                {
                    throw new InvalidOperationException("此元素不是分隔線，無法執行拖曳操作");
                }

                // 2. 使用 DragAndDropToOffset 執行水平拖曳
                // - 設定從分隔線中心開始拖曳
                // - Y方向偏移為0，保持垂直位置不變
                splitter.DragAndDropToOffset(
                        MoveToElementOffsetStartingPoint.Center,  // 從中心點開始拖曳
                        0,                                        // 起始X偏移
                        0,                                        // 起始Y偏移
                        offsetX,                                  // 水平拖曳距離
                        0                                         // 垂直拖曳距離(保持不變)
                );

                // 3. 等待UI更新
                Thread.Sleep(100);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"水平拖曳失敗: {ex.Message}");
            }
        }

        public static void DragVertically(this IElement splitter, int offsetY)
        {
            try
            {
                // 1. 確認元素 ControlType 是否為 Thumb (ClassName: GridSplitter)
                if (splitter.ControlType != ElementControlType.Thumb)
                {
                    throw new InvalidOperationException("此元素不是分隔線，無法執行拖曳操作");
                }

                // 2. 使用 DragAndDropToOffset 執行垂直拖曳
                // - 設定從分隔線中心開始拖曳
                // - X方向偏移為0，保持水平位置不變
                splitter.DragAndDropToOffset(
                        MoveToElementOffsetStartingPoint.Center,    // 從中心點開始拖曳
                        0,                                          // 起始X偏移
                        0,                                          // 起始Y偏移
                        0,                                          // 水平拖曳距離(保持不變)
                        offsetY                                     // 垂直拖曳距離
                );

                // 3. 等待UI更新
                Thread.Sleep(100);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"垂直拖曳失敗: {ex.Message}");
            }
        }

        public static void DragSplitter(this IElement splitter, System.Windows.Forms.Orientation orientation, int offset)
        {
            try
            {
                // 1. 確認元素 ControlType 是否為 Thumb (ClassName: GridSplitter)
                if (splitter.ControlType != ElementControlType.Thumb)
                {
                    throw new InvalidOperationException("此元素不是分隔線，無法執行拖曳操作");
                }

                int offsetX = 0, offsetY = 0;
                if (orientation == System.Windows.Forms.Orientation.Horizontal)     // 設定水平拖曳
                    offsetX = offset;
                else if (orientation == System.Windows.Forms.Orientation.Vertical)  // 設定垂直拖曳
                    offsetY = offset;

                // 2. 使用 DragAndDropToOffset 執行垂直/水平拖曳
                // - 設定從分隔線中心開始拖曳
                splitter.DragAndDropToOffset(
                        MoveToElementOffsetStartingPoint.Center,    // 從中心點開始拖曳
                        0,                                          // 起始X偏移
                        0,                                          // 起始Y偏移
                        offsetX,                                    // 水平拖曳距離
                        offsetY                                     // 垂直拖曳距離
                );

                // 3. 等待UI更新
                Thread.Sleep(100);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"垂直/水平拖曳失敗: {ex.Message}");
            }
        }

        /// <summary>
        /// 設定視窗大小
        /// </summary>
        /// <param name="width">視窗寬度</param>
        /// <param name="height">視窗高度</param>
        public static void SetWindowSize(this IElement window, int width, int height)
        {
            if (!window.IsWindow)
                throw new InvalidOperationException("此元素不是視窗類型，無法調整大小");
            if (window.IsModal && !window.CanResize)
                throw new InvalidOperationException("此視窗為modal且無法調整大小");

            // 先確認視窗是否為最大化狀態,如果是，要先解除最大化狀態(雙擊視窗最上面中間部分)
            // 接著判斷CanResize是否為true，如果為true，直接進行調整視窗大小
            window.RestoreWindow();
            if (!window.CanResize)
                throw new InvalidOperationException("此視窗無法調整大小");

            // 1. 取得目前視窗位置和大小
            var windowPosition = window.PointAtTopLeft;

            // 2. 使用 DragAndDropToOffset 方法調整視窗大小
            // 從視窗右下角拖曳來調整大小
            window.DragAndDropToOffset(
                MoveToElementOffsetStartingPoint.BottomRight,   // 從右下角開始拖曳
                0, 0,                                           // 起始偏移
                width - window.Width,                           // X方向拖曳距離 
                height - window.Height                          // Y方向拖曳距離
            );

            // 3. 確認視窗已調整至指定大小
            int newWidth = window.Width;
            int newHeight = window.Height;

            if (newWidth != width || newHeight != height)
                throw new InvalidOperationException($"視窗大小調整失敗。預期: {width}x{height}, 實際: {newWidth}x{newHeight}");
        }

        public static WindowVisualState RestoreWindow(this IElement winElement)
        {
            if (!winElement.IsWindow)
                throw new InvalidOperationException("此元素不是視窗類型，無法調整顯示狀態");
            if (winElement.VisualState == WindowVisualState.Normal)
                throw new InvalidOperationException("此視窗已經正常顯示");

            if (winElement.VisualState == WindowVisualState.Maximized)
            {
                //window.MoveToElement(MoveToElementOffsetStartingPoint.InnerCenterTop, 0, 0);
                //window.DoubleClick();
                winElement.PerformClick("ById[Restore]", ClickType.LeftClick);
            }
            return winElement.VisualState;
        }

        public static bool MinimizeWindow(this IElement winElement)
        {
            if (!winElement.IsWindow)
                throw new InvalidOperationException("此元素不是視窗類型，無法調整顯示狀態");
            if (winElement.VisualState == WindowVisualState.Minimized)
                throw new InvalidOperationException("此視窗已經最小化");

            return winElement.PerformClick($"/ById[Minimize]", ClickType.LeftClick) != null && winElement.VisualState == WindowVisualState.Minimized;
        }

        public static bool MaximizeWindow(this IElement winElement)
        {
            if (!winElement.IsWindow)
                throw new InvalidOperationException("此元素不是視窗類型，無法調整顯示狀態");
            if (winElement.VisualState == WindowVisualState.Maximized)
                throw new InvalidOperationException("此視窗已經最大化");

            return winElement.PerformClick($"/ById[Maximize]", ClickType.LeftClick) != null && winElement.VisualState == WindowVisualState.Maximized;
        }

        public static bool CloseWindow(this IElement winElement)
        {
            if (!winElement.IsWindow)
                throw new InvalidOperationException("此元素不是視窗類型，無法調整顯示狀態");

            return winElement.PerformClick($"/ById[Close]", ClickType.LeftClick) != null;
        }

        public static bool CloseWindow(this IElement ideElement, int windowNo)
        {
            if (!HasWindow(ideElement, windowNo, out var windowElement)) 
                return false;
            return windowElement.PerformClick($"/ById[Close]", ClickType.LeftClick) != null;
        }

        public static bool HasWindow(this IElement ideElement, int windowNo)
        {
            return ideElement.PerformGetElement($"/Window[{windowNo}]") != null;
        }

        public static bool HasWindow(this IElement ideElement, int windowNo, out IElement windowElement)
        {
            windowElement = ideElement.PerformGetElement($"/Window[{windowNo}]");
            return windowElement != null;
        }

        public static int GetWindowCount(this IElement ideElement)
        {
            var count = 0;
            IElement windowTemp = ideElement;
            while (HasWindow(windowTemp, 0, out IElement window))
            {
                windowTemp = window;
                count++;
            }
            return count;
        }
        #endregion
    }
}
