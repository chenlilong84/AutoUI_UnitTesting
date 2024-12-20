using System;
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

        //public static bool IsMenuItemEnabled(this IWebElement element, params string[] itemNames)
        //{
        //    return element.GetMenuItem(itemNames).Enabled;
        //}

        //public static IWebElement GetMenuItem(this IWebElement element, params string[] itemNames)
        //{
        //    IWebElement menu = element.GetElementFromWebElement(By.ClassName("Menu"));

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

        /// <summary>
        /// Menu selection helper function
        /// </summary>
        /// <param name="itemNames">name of the menuitem to be seleted in stirng</param>
        public static bool MenuSelect(params string[] itemNames)
        {
            var driver = Executor.GetInstance().GetCurrentDriver();
            IElement menu = driver.GetElement<PP5Driver,IElement>(PP5By.ClassName("Menu"));
            bool isSuccess = true;

            foreach (string itemName in itemNames)
            {
                //Console.WriteLine($"LeftClick on Text \"{itemName}\"");
                //menu.GetElementFromWebElement(By.XPath($".//Text[@Name='{itemName}']")).LeftClick();
                //menu.GetTextElement(itemName).LeftClick();
                //menu.GetTextElement(itemName).LeftClick();
                //if (menu.GetElement<IElement, IElement>(By.Name(itemName)).Enabled)

                IElement menuItem = menu.GetElement<IElement, IElement>(PP5By.Name(itemName));
                menuItem.LeftClick();
                if (!menuItem.Enabled)
                {
                    if (menuItem != null)
                        Press(Keys.Escape);
                        //driver.GetGrandParentElement(menu).LeftClick();
                    isSuccess = false;
                    break;
                }
                //isSuccess = menuItem.Enabled & isSuccess;
            }
            return isSuccess;
        }

        public static IEnumerable<string> GetSubMenuListItemNames(string itemName)
        {
            IElement menu = Executor.GetInstance().GetCurrentDriver().GetExtendedElement(PP5By.ClassName("Menu"));
            IEnumerable<string> MenuListItemsNames;

            //Console.WriteLine($"LeftClick on Text \"{itemName}\"");
            //IWebElement subMenu = menu.GetElementFromWebElement(By.XPath($".//Text[@Name='{itemName}']/.."));
            IElement subMenu = menu.GetExtendedElement(PP5By.Name(itemName));
            subMenu.LeftClick();
            MenuListItemsNames = subMenu.GetChildElements().Select(e => e.Text);
            subMenu.LeftClick();
            return MenuListItemsNames;
        }

        public static IEnumerable<string> GetMainMenuListItemNames()
        {
            IElement menu = Executor.GetInstance().GetCurrentDriver().GetExtendedElement(PP5By.ClassName("Menu"));
            return menu.GetMenuItems().Select(e => e.Text);
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
            cmbItems = comboBox.GetExtendedElements(PP5By.ClassName("ListBoxItem"));
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

        public static void SelectComboBoxItemByIndex2(this IElement comboBox, int index)
        {
            comboBox.GetComboBoxItems(out ReadOnlyCollection<IElement> cmbItems);
            if (cmbItems.Count() >= index + 1)
            {
                comboBox.SelectComboBoxItemByName2(cmbItems[index].Text);
            }
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
            //element.GetElements(By.XPath(".//DataItem"))[rowIdx].LeftClick();

            //ReadOnlyCollection<IWebElement> rows = element.GetChildElementsOfControlType(ElementControlType.DataItem);
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
            // if the element is the tree item leaf node: LeafNode (3), it's not expandable
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
                tabControlTemp = tabItem.GetExtendedElement(PP5By.ClassName("TabControl"));
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
                tabControlTemp = tabItem.GetExtendedElement(PP5By.ClassName("TabControl"));
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
                tabControlTemp = tabItem.GetExtendedElement(PP5By.ClassName("TabControl"));
            }

            return tabItem;
        }

        public static IElement TabSelect(this IElement tabControl, int index)
        {
            IElement tabItem = tabControl.GetExtendedElements(PP5By.ClassName("TabItem"))[index];
            if (!tabItem.Selected && tabItem.isElementVisible())
                tabItem.LeftClick();
            return tabItem;
        }

        public static IElement TabSelect(this IElement tabControl, string tabName)
        {
            //IWebElement tabItem = tabControl.GetTabItemElement(tabName);
            IElement tabItem = tabControl.GetExtendedElement(PP5By.Name(tabName));
            if (!tabItem.Selected && tabItem.isElementVisible())
                tabItem.LeftClick();
            return tabItem;
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
        public static bool CloseWindow(this IElement ideElement, int windowNo)
        {
            return ideElement.PerformClick($"/Window[{windowNo}]/ById[Close]", ClickType.LeftClick) != null;
        }

        public static bool HasWindow(this IElement ideElement, int windowNo)
        {
            return ideElement.PerformGetElement($"/Window[{windowNo}]") != null;
        }
        #endregion
    }
}
