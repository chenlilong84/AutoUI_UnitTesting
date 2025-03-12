using Chroma.UnitTest.Common;
using Microsoft.VisualStudio.TestTools.UnitTesting.Logging;
using OpenQA.Selenium;

namespace PP5AutoUITests
{
    public static class PerformAction
    {
        static PerformActionHelper AH = new PerformActionHelper();

        /// <summary>
        /// For elementPath format, please refer to <see cref="PerformGetElement"/> for detailed information
        /// <para>Example:</para>
        /// <example>
        /// </example>
        /// </summary>
        /// <param name="elementSrc"></param>
        /// <param name="elementPath"></param>
        /// <param name="clickType"></param>
        /// <returns></returns>
        public static IElement PerformClick(this IElement elementSrc, string elementPath/*node1(className1)/node2(Name2)/node3(AutoId3)*/, params ClickType[] clickTypes)
        {
            IElement elementFound = elementSrc.PerformGetElement(elementPath);
            elementFound.PerformClick(clickTypes);
            return elementFound;
        }

        public static IElement PerformClick(this IElement elementSrc, string elementPath, MoveToElementOffsetStartingPoint clickingPosition, params ClickType[] clickTypes)
        {
            IElement elementFound = elementSrc.PerformGetElement(elementPath);
            ((PP5Element)elementFound).MoveToElement(clickingPosition, 0, 0);
            elementFound.PerformClick(clickTypes);
            return elementFound;
        }

        /// <summary>
        /// For elementPath format, please refer to <see cref="PerformGetElement(PP5Driver, string)"/> for detailed information
        /// <para>Example:</para>
        /// <example>
        /// </example>
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="elementPath"></param>
        /// <param name="clickTypes"></param>
        /// <returns></returns>
        public static IElement PerformClick(this PP5Driver driver, string elementPath/*node1(className1)/node2(Name2)/node3(AutoId3)*/, params ClickType[] clickTypes)
        {
            IElement elementFound = driver.PerformGetElement(elementPath);
            elementFound.PerformClick(clickTypes);
            return elementFound;
        }

        public static IElement PerformClick(this PP5Driver driver, string elementPath, MoveToElementOffsetStartingPoint clickingPosition, params ClickType[] clickTypes)
        {
            IElement elementFound = driver.PerformGetElement(elementPath);
            ((PP5Element)elementFound).MoveToElement(clickingPosition, 0, 0);
            elementFound.PerformClick(clickTypes);
            return elementFound;
        }

        /// <summary>
        /// For elementPath format, please refer to <see cref="PerformGetElement(IElement, string)"/> for detailed information
        /// <para>Example:</para>
        /// <example>
        /// </example>
        /// </summary>
        /// <param name="elementSrc"></param>
        /// <param name="elementPath"></param>
        /// <param name="inputType"></param>
        /// <param name="inputParams"></param>
        /// <returns></returns>
        public static IElement PerformInput(this IElement elementSrc, string elementPath/*node1(className1)/node2(Name2)/node3(AutoId3)*/, InputType inputType, object inputParams)
        {
            IElement elementFound = elementSrc.PerformGetElement(elementPath);
            elementFound.PerformInput(inputType, inputParams);
            return elementFound;
        }

        /// <summary>
        /// For elementPath format, please refer to <see cref="PerformGetElement(PP5Driver, string)"/> for detailed information
        /// <para>Example:</para>
        /// <example>
        /// </example>
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="elementPath"></param>
        /// <param name="inputType"></param>
        /// <param name="inputParams"></param>
        /// <returns></returns>
        public static IElement PerformInput(this PP5Driver driver, string elementPath/*node1(className1)/node2(Name2)/node3(AutoId3)*/, InputType inputType, object inputParams)
        {
            IElement elementFound = driver.PerformGetElement(elementPath);
            elementFound.PerformInput(inputType, inputParams);
            return elementFound;
        }

        /// <summary>
        /// Supported elementPath types:
        /// <list type="number">
        /// <item>
        ///     <term>DataGrid related</term>
        ///     <description>
        ///         <list type="bullet">
        ///             <item>Find By DataGrid: <b>"/ByDataGrid@{DataGridId}[TBD]"</b></item>
        ///             <item>Find By Row: <b>"/ByRow[TBD]@{DataGridId}"</b></item>
        ///             <item>Find By Column: <b>"/ByColumn@{DataGridId}[TBD]"</b></item>
        ///             <item>Find By Cell: <b>"/ByCell@{DataGridId}[Refer to <see cref="PerformActionHelper.PerformCellMethod"/>]"</b></item>
        ///         </list>
        ///     </description>
        /// </item>
        /// <item>
        ///     <term>Using By Identifier</term>
        ///     <description>
        ///         <list type="bullet">
        ///             <item>Find By Name: <b>"/ByName[{Name}]"</b></item>
        ///             <item>Find By Id: <b>"/ById[{AutomationID}]"</b></item>
        ///             <item>Find By Id Or Name: <b>"/ByIdOrName[{AutomationID/Name}]"</b> or <b>"/ByNameOrId[{AutomationID/Name}]"</b></item>
        ///             <item>Find By ClassName: <b>"/ByClass[{ClassName}]"</b></item>
        ///         </list>
        ///     </description>
        /// </item>
        /// <item>
        ///     <term>Find by condition</term>
        ///     <description>
        ///         <list type="bullet">
        ///             <item>
        ///                 <description>ElementPath format: <b>/ByCondition[{ConditionExpression}]</b><br/>
        ///                 where {ConditionExpression} uses the format: <c>Func&lt;IElement, bool&gt;</c></description>
        ///             </item>
        ///         </list>
        ///     </description>
        /// </item>
        /// <item>
        ///     <term>Find by ControlType &amp; Name/Id</term>
        ///     <description>
        ///         <list type="bullet">
        ///             <item>
        ///                 ElementPath format: <b>/{ControlType}[{NameOrId}]</b><br/>
        ///                 where {ControlType} is Window, Button, Text, etc., and {NameOrId} is the identifier that uses Name or AutomationID
        ///             </item>
        ///         </list>
        ///     </description>
        /// </item>
        /// </list>
        /// <para>General elementPath format: <b>/ControlTypeName1[locatorName1]/ControlTypeName2[locatorName2]</b></para>
        /// <para>Example:</para>
        /// <example>
        ///     <list type="bullet">
        ///         <item>
        ///             <description>elementPath: <b>/Window[Notice]/Button[Yes]</b><br/>
        ///             Find the window with name: "Notice", then find the button with name "Yes".</description>
        ///         </item>
        ///         <item>
        ///             <description>elementPath: <b>/Window[Notice]/Button[0]</b><br/>
        ///             Find the window with name: "Notice", then find the first (0) button.</description>
        ///         </item>
        ///     </list>
        /// </example>
        /// </summary>
        /// <param name="elementSrc">The source element from which to begin searching.</param>
        /// <param name="elementPath">The path used to locate the target element.</param>
        /// <returns>Returns the located element or null if not found.</returns>
        public static IElement PerformGetElement(this IElement elementSrc, string elementPath)
        {
            // Parsing the element path
            IElement elementFound = elementSrc;
            var elementPaths = elementPath.Split('/').ToList();
            elementPaths.RemoveAt(0);

            // 2024/12/06, return self-element if elementPath: "/" is given
            if (elementPaths.Count == 1 && elementPaths[0] == null)
                return elementFound;

            foreach (var elePath in elementPaths)
                elementFound = AH.HandleAllCases(elementSrc, elePath);

            /* Legacy method using If-else
                //if (elementControlType.StartsWith("ByDataGrid", StringComparison.InvariantCultureIgnoreCase) ||
                //    elementControlType.StartsWith("ByCell", StringComparison.InvariantCultureIgnoreCase) ||
                //    elementControlType.StartsWith("ByRow", StringComparison.InvariantCultureIgnoreCase) ||
                //    elementControlType.StartsWith("ByColumn", StringComparison.InvariantCultureIgnoreCase))
                //{
                //    if (elementControlType.Contains('@'))
                //    {
                //        // Use AutomationID
                //        var dataGridId = elementControlType.Split('@')[1];
                //        var findType = elementControlType.Split('@')[0];
                //        elementFound = elementSrc.GetDataGridElement(dataGridId).PerformDataGridMethod(findType, elementLocator);
                //    }
                //    else
                //        elementFound = elementSrc.PerformDataGridMethod(elementControlType, elementLocator);
                //}
                //else if (elementControlType.Equals("ByIdOrName", StringComparison.InvariantCultureIgnoreCase) ||
                //         elementControlType.Equals("ByNameOrId", StringComparison.InvariantCultureIgnoreCase))
                //{
                //    elementFound = elementSrc.GetElementFromWebElement(new ByAutomationIdOrName(elementLocator));
                //}
                //else if (elementControlType.Equals("ByName", StringComparison.InvariantCultureIgnoreCase))
                //{
                //    elementFound = elementSrc.GetElementFromWebElement(By.Name(elementLocator));
                //}
                //else if (elementControlType.Equals("ById", StringComparison.InvariantCultureIgnoreCase))
                //{
                //    elementFound = elementSrc.GetElementFromWebElement(MobileBy.AccessibilityId(elementLocator));
                //}
                //else if (elementControlType.Equals("ByClass", StringComparison.InvariantCultureIgnoreCase))
                //{
                //    elementFound = elementSrc.GetElementFromWebElement(By.ClassName(elementLocator));
                //}
                //else if (elementControlType.StartsWith("ByCondition", StringComparison.InvariantCultureIgnoreCase))
                //{
                //    if (elementControlType.Contains('#'))
                //    {
                //        //elementFound = elementSrc.GetSpecificChildOfControlTypeByBFS(By.ClassName(elementLocator));
                //        Func<IWebElement, bool> condition = null;
                //        if (elementLocator == "Visible")
                //            condition = (e) => e.isElementVisible();

                //        var controlType = elementControlType.Split('#')[1];
                //        elementFound = elementSrc.GetSpecificChildOfControlType(TypeExtension.GetEnumByDescription<ElementControlType>(controlType), condition);
                //    }
                //}
                //else
                //{
                //    ElementControlType eleCtrlType = TypeExtension.GetEnumByDescription<ElementControlType>(elementControlType);

                //    if (int.TryParse(elementLocator, out int elementLocatorIndex))
                //    {
                //        elementFound = elementSrc.GetSpecificChildOfControlTypeByBFS(eleCtrlType, elementLocatorIndex);
                //    }
                //    else
                //        elementFound = elementSrc.GetSpecificChildOfControlTypeByBFS(eleCtrlType, elementLocator);
                //}
                */
            return elementFound;
        }

        /// <summary>
        /// Supported elementPath types:
        /// <list type="number">
        /// <item>
        ///     <term>DataGrid related</term>
        ///     <description>
        ///         <list type="bullet">
        ///             <item>Find By DataGrid: <b>"/ByDataGrid@{DataGridId}[TBD]"</b></item>
        ///             <item>Find By Row: <b>"/ByRow[TBD]"</b></item>
        ///             <item>Find By Column: <b>"/ByColumn[TBD]"</b></item>
        ///             <item>Find By Cell: <b>"/ByCell[Refer to PerformCellMethod]"</b></item>
        ///         </list>
        ///     </description>
        /// </item>
        /// <item>
        ///     <term>Using By Identifier</term>
        ///     <description>
        ///         <list type="bullet">
        ///             <item>Find By Name: <b>"/ByName[{Name}]"</b></item>
        ///             <item>Find By Id: <b>"/ById[{AutomationID}]"</b></item>
        ///             <item>Find By Id Or Name: <b>"/ByIdOrName[{AutomationID/Name}]"</b> or <b>"/ByNameOrId[{AutomationID/Name}]"</b></item>
        ///             <item>Find By ClassName: <b>"/ByClass[{ClassName}]"</b></item>
        ///         </list>
        ///     </description>
        /// </item>
        /// <item>
        ///     <term>Find by condition</term>
        ///     <description>
        ///         <list type="bullet">
        ///             <item>
        ///                 <description>ElementPath format: <b>/ByCondition[{ConditionExpression}]</b><br/>
        ///                 where {ConditionExpression} uses the format: <c>Func&lt;IElement, bool&gt;</c></description>
        ///             </item>
        ///         </list>
        ///     </description>
        /// </item>
        /// <item>
        ///     <term>Find by ControlType &amp; Name/Id</term>
        ///     <description>
        ///         <list type="bullet">
        ///             <item>
        ///                 ElementPath format: <b>/{ControlType}[{NameOrId}]</b><br/>
        ///                 where {ControlType} is Window, Button, Text, etc., and {NameOrId} is the identifier that uses Name or AutomationID
        ///             </item>
        ///         </list>
        ///     </description>
        /// </item>
        /// </list>
        /// <para>General elementPath format: <b>/ControlTypeName1[locatorName1]/ControlTypeName2[locatorName2]</b></para>
        /// <para>Example:</para>
        /// <example>
        ///     <list type="bullet">
        ///         <item>
        ///             <description>elementPath: <b>/Window[Notice]/Button[Yes]</b><br/>
        ///             Find the window with name: "Notice", then find the button with name "Yes".</description>
        ///         </item>
        ///         <item>
        ///             <description>elementPath: <b>/Window[Notice]/Button[0]</b><br/>
        ///             Find the window with name: "Notice", then find the first (0) button.</description>
        ///         </item>
        ///     </list>
        /// </example>
        /// </summary>
        /// <param name="driver">The source driver from which to begin searching.</param>
        /// <param name="elementPath">The path used to locate the target element.</param>
        /// <returns>Returns the located element or null if not found.</returns>
        public static IElement PerformGetElement(this PP5Driver driver, string elementPath)
        {
            // Parsing the element path
            IElement elementFound = null;
            var elementPaths = elementPath.Split('/').ToList();
            elementPaths.RemoveAt(0);

            foreach (var elePath in elementPaths)
            {
                var parts = elePath.Split('[');
                string elementControlType = parts[0];
                string elementLocator = parts[1].TrimEnd(']');

                switch (AH.GetControlTypeCategory(elementControlType))
                {
                    case PerformTypeCategory.DataGrid:
                        elementFound = AH.HandleDataGridCases(driver, elementControlType, elementLocator);
                        break;

                    case PerformTypeCategory.ByIdOrName:
                        elementFound = AH.HandleByIdOrNameCase(driver, elementControlType, elementLocator);
                        break;

                    case PerformTypeCategory.ByName:
                        elementFound = AH.HandleByNameCase(driver, elementControlType, elementLocator);
                        break;

                    case PerformTypeCategory.ById:
                        elementFound = AH.HandleByIdCase(driver, elementControlType, elementLocator);
                        break;

                    case PerformTypeCategory.ByClass:
                        elementFound = AH.HandleByClassCase(driver, elementControlType, elementLocator);
                        break;
                }

                /*
                if (elementControlType.StartsWith("ByDataGrid", StringComparison.InvariantCultureIgnoreCase) ||
                    elementControlType.StartsWith("ByCell", StringComparison.InvariantCultureIgnoreCase) ||
                    elementControlType.StartsWith("ByRow", StringComparison.InvariantCultureIgnoreCase) ||
                    elementControlType.StartsWith("ByColumn", StringComparison.InvariantCultureIgnoreCase))
                {
                    if (elementControlType.Contains('@'))
                    {
                        // Use AutomationID
                        var dataGridId = elementControlType.Split('@')[1];
                        var findType = elementControlType.Split('@')[0];
                        elementFound = driver.GetElementFromWebElement(MobileBy.AccessibilityId(dataGridId)).PerformDataGridMethod(findType, elementLocator);
                    }
                }
                else if (elementControlType.Equals("ByIdOrName", StringComparison.InvariantCultureIgnoreCase) ||
                         elementControlType.Equals("ByNameOrId", StringComparison.InvariantCultureIgnoreCase))
                {
                    elementFound = driver.GetElementFromWebElement(new ByAutomationIdOrName(elementLocator));
                }
                else if (elementControlType.Equals("ByName", StringComparison.InvariantCultureIgnoreCase))
                {
                    elementFound = driver.GetElementFromWebElement(By.Name(elementLocator));
                }
                else if (elementControlType.Equals("ById", StringComparison.InvariantCultureIgnoreCase))
                {
                    elementFound = driver.GetElementFromWebElement(MobileBy.AccessibilityId(elementLocator));
                }
                else if (elementControlType.Equals("ByClass", StringComparison.InvariantCultureIgnoreCase))
                {
                    elementFound = driver.GetElementFromWebElement(By.ClassName(elementLocator));
                }
                */
            }
            return elementFound;
        }

        public static List<IElement> PerformGetElements(this IElement elementSrc, string elementPath)
        {
            // Parsing the element path
            IElement elementTmp = elementSrc;
            List<IElement> elementsFound = new List<IElement>();
            var elementPaths = elementPath.Split('/').ToList();
            elementPaths.RemoveAt(0);

            // 2024/12/06, return self-element if elementPath: "/" is given
            if (elementPaths.Count == 1 && elementPaths[0] == null)
                return new List<IElement>() { elementSrc };

            for (int i = 0; i < elementPaths.Count; i++)
            {
                if (i == elementPaths.Count - 1)
                    elementsFound = AH.HandleAllCases_multiElements(elementTmp, elementPaths[i]);
                else
                    elementTmp = AH.HandleAllCases(elementTmp, elementPaths[i]);
            }
            return elementsFound;
        }

        private static void PerformClick(this IElement element, ClickType[] clickTypes)
        {
            foreach (var clickType in clickTypes)
            {
                switch (clickType)
                {
                    case ClickType.None:
                        break;
                    case ClickType.LeftClick:
                        element?.LeftClick();
                        break;
                    case ClickType.LeftDoubleClick:
                        element?.DoubleClick();
                        break;
                    case ClickType.RightClick:
                        element?.RightClick();
                        break;
                    case ClickType.LeftDoubleClickDelay:
                        element?.LeftClickWithDelay(50);
                        break;
                    case ClickType.TickCheckBox:
                        element?.TickCheckBox();
                        break;
                    case ClickType.UnTickCheckBox:
                        element?.UnTickCheckBox();
                        break;
                }
            }
        }

        private static void PerformInput(this IElement element, InputType inputType, object inputParams)
        {
            switch (inputType)
            {
                case InputType.None:
                case InputType.SelectAllContent:
                case InputType.ClearContent:
                case InputType.PasteContent:
                case InputType.CopyContent:
                case InputType.CutContent:
                    element?.PerformInputAction(inputType);
                    break;

                case InputType.SendSingleKeys:
                case InputType.SendComboKeys:
                case InputType.SendContent:
                case InputType.CopyAndPaste:
                case InputType.CutAndPaste:
                    element?.PerformInputAction(inputType, inputParams);
                    break;
            }
        }

        private static void PerformInputAction(this IElement element, InputType inputType)
        {
            switch (inputType)
            {
                case InputType.None:
                    break;
                case InputType.SelectAllContent:
                    element.SelectAllContent();
                    break;
                case InputType.ClearContent:
                    element.ClearContent();
                    break;
                case InputType.PasteContent:
                    element.PasteContent();
                    break;
                case InputType.CopyContent:
                    element.CopyContent();
                    break;
                case InputType.CutContent:
                    element.CutContent();
                    break;
            }
        }

        private static void PerformInputAction(this IElement element, InputType inputType, object inputParams)
        {
            switch (inputType)
            {
                case InputType.SendSingleKeys:
                    string inputString = inputParams as string;
                    element?.SendSingleKeys(inputString);
                    break;
                case InputType.SendComboKeys:
                    string[] inputStringList = inputParams as string[];
                    element?.SendComboKeys(inputStringList);
                    break;
                case InputType.SendContent:
                    string keysToSend = inputParams as string;
                    element?.SendText(keysToSend);
                    break;
                case InputType.CopyAndPaste:
                case InputType.CutAndPaste:
                    IElement ToElement = inputParams as IElement;
                    element?.PerformInputMixedAction(inputType, ToElement);
                    break;
            }
        }

        private static void PerformInputMixedAction(this IElement element, InputType inputType, IElement ToElement)
        {
            switch (inputType)
            {
                case InputType.CopyAndPaste:
                    element.CopyAndPaste(ToElement);
                    break;
                case InputType.CutAndPaste:
                    element.CutAndPaste(ToElement);
                    break;
            }
        }
    }

    internal class PerformActionHelper
    {
        //static PerformActionHelper ActionHelper;
        //PerformActionHelper()
        //{

        //}

        //internal PerformActionHelper GetInstance()
        //{
        //    if (ActionHelper == null)
        //        return new PerformActionHelper();
        //    return ActionHelper;
        //}

        private IElement PerformDataGridMethod(IElement eleDataGrid, string findType, string elementLocator)
        {
            // cell methods:        ByCell (elementLocator)
            // column methods:      ByColumn (elementLocator)
            // row methods:         ByRow(elementLocator)
            // DataGrid methods:    ByDataGrid(elementLocator)
            //var elementLocatorPaths = elementLocator.Split('[').ToList();
            //var elementType = elementLocatorPaths[0];       // Cell, Col, Row
            //var locatorParams = elementLocatorPaths[1].TrimEnd(']');
            IElement elementResult = null;

            try
            {
                switch (findType)
                {
                    case "ByCell":
                        elementResult = PerformCellMethod(eleDataGrid,elementLocator);
                        break;
                    case "ByColumn":
                        elementResult = PerformColumnMethod(eleDataGrid, elementLocator);
                        break;
                    case "ByRow":
                        elementResult = PerformRowMethod(eleDataGrid, elementLocator);
                        break;
                    case "ByDataGrid":
                        elementResult = PerformDataGridMethod(eleDataGrid, elementLocator);
                        break;
                }
            }
            catch (Exception ex)
            {
                Logger.LogMessage(ex.Message);
            }
            return elementResult;
        }

        /// <summary>
        /// Perform Cell related methods based on the elementLocator format.
        /// <list type="bullet">
        ///     <item>
        ///         <description>
        ///             Case 1: elementLocator = "colNo" <br/>
        ///             Params: <b>(int colNo)</b>
        ///         </description>
        ///     </item>
        ///     <item>
        ///         <description>
        ///             Case 2: elementLocator = "colNo,#cellName" <br/>
        ///             Params: <b>(int colNo, string cellName)</b>
        ///         </description>
        ///     </item>
        ///     <item>
        ///         <description>
        ///             Case 3: elementLocator = "rowNo,colNo" <br/>
        ///             Params: <b>(int rowNo, int colNo)</b>
        ///         </description>
        ///     </item>
        ///     <item>
        ///         <description>
        ///             Case 4: elementLocator = "rowNo,@colName" <br/>
        ///             Params: <b>(int rowNo, string colName)</b>
        ///         </description>
        ///     </item>
        /// </list>
        /// </summary>
        /// <param name="eleDataGrid">The DataGrid element containing the cells.</param>
        /// <param name="elementLocator">The input params as a string for the cell method.</param>
        /// <returns>The located IElement based on the specified elementLocator.</returns>
        /// <exception cref="Exception">Thrown when the locator format is incorrect.</exception>

        private IElement PerformCellMethod(IElement eleDataGrid, string elementLocator)
        {
            // Case 1: elementLocator = "colNo",                params = (int colNo)
            // Case 2: elementLocator = "colNo,#cellName",      params = (int colNo, string cellName)
            // Case 3: elementLocator = "rowNo,colNo",          params = (int rowNo, int colNo)
            // Case 4: elementLocator = "rowNo,@colName",        params = (int rowNo, string colName)
            IElement elementResult = null;
            if (!CheckLocatorFormat(elementLocator, out int formatType, out object formatParams))
                throw new Exception("Locator Format incorrect!");

            switch (formatType)
            {
                case 1:     // Case 1: GetCellBy(int colNo)
                    elementResult = eleDataGrid.GetSelectedRow().GetCellBy((int)formatParams);
                    break;
                case 2:     // Case 2: GetCellByName(int colNo, string cellName)
                    elementResult = eleDataGrid.GetCellByName((int)((List<object>)formatParams)[0], ((List<object>)formatParams)[1].ToString());
                    break;
                case 3:     // Case 3: GetCellBy(int rowNo, int colNo))
                    elementResult = eleDataGrid.GetCellBy((int)((List<object>)formatParams)[0], (int)((List<object>)formatParams)[1]);
                    break;
                case 4:     // Case 4: GetCellBy(int rowNo, string colName)
                    elementResult = eleDataGrid.GetCellBy((int)((List<object>)formatParams)[0], ((List<object>)formatParams)[1].ToString());
                    break;
            }

            return elementResult;
        }

        private IElement PerformColumnMethod(IElement eleDataGrid, string locatorParams)
        {
            return null;
        }

        private IElement PerformRowMethod(IElement eleDataGrid, string locatorParams)
        {
            return null;
        }

        private IElement PerformDataGridMethod(IElement eleDataGrid, string locatorParams)
        {
            return null;
        }

        private bool CheckLocatorFormat(string locator, out int formatType, out object formatParams)
        {
            formatType = -1;
            formatParams = null;
            bool isLegalFormat = true;

            // Case 1: elementLocator = "colNo",            
            // Case 2: elementLocator = "colNo,#cellName",      
            // Case 3: elementLocator = "rowNo,colNo",          
            // Case 4: elementLocator = "rowNo,@colName",
            if (!locator.Contains(','))
                // Case 1: GetCellBy(this IWebElement rowElement, int colNo)
                if (int.TryParse(locator, out int locatorInt))
                {
                    if (!(locatorInt > 0 || locatorInt == -1))
                        isLegalFormat = false;
                    else
                    {
                        formatType = 1;
                        formatParams = locatorInt;
                    }
                }
                else
                    isLegalFormat = false;
            else
            {
                string[] indecesOrNames = locator.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                var firstParam = indecesOrNames[0];
                var secondParam = indecesOrNames[1];

                List<object> Params = new List<object>();

                if (int.TryParse(firstParam, out int locatorInt))
                {
                    if (!(locatorInt > 0 || locatorInt == -1))
                        isLegalFormat = false;
                    else
                    {
                        Params.Add(locatorInt);
                    }
                }
                else
                    isLegalFormat = false;

                if (int.TryParse(secondParam, out locatorInt))
                {
                    if (!(locatorInt > 0 || locatorInt == -1))
                        isLegalFormat = false;
                    else
                    {
                        formatType = 3;
                        Params.Add(locatorInt);
                    }
                }
                else
                {
                    if (secondParam.StartsWith("#"))
                    {
                        formatType = 2;
                        Params.Add(secondParam.Remove(0, 1));
                    }
                    else if (secondParam.StartsWith("@"))
                    {
                        formatType = 4;
                        Params.Add(secondParam.Remove(0, 1));
                    }
                    else
                        isLegalFormat = false;
                }
                formatParams = Params;
            }

            return isLegalFormat;
        }

        internal IElement HandleAllCases(IElement elementSrc, string singleElePath)
        {
            var parts = singleElePath.Split('[');
            string elementControlType = parts[0];
            string elementLocator = parts[1].TrimEnd(']');
            IElement elementFound = null;

            switch (GetControlTypeCategory(elementControlType))
            {
                case PerformTypeCategory.DataGrid:
                    elementFound = HandleDataGridCases(elementSrc, elementControlType, elementLocator);
                    break;

                case PerformTypeCategory.ByIdOrName:
                    elementFound = HandleByIdOrNameCase(elementSrc, elementControlType, elementLocator);
                    break;

                case PerformTypeCategory.ByName:
                    elementFound = HandleByNameCase(elementSrc, elementControlType, elementLocator);
                    break;

                case PerformTypeCategory.ById:
                    elementFound = HandleByIdCase(elementSrc, elementControlType, elementLocator);
                    break;

                case PerformTypeCategory.ByClass:
                    elementFound = HandleByClassCase(elementSrc, elementControlType, elementLocator);
                    break;

                case PerformTypeCategory.ByCondition:
                    elementFound = HandleByConditionCase(elementSrc, elementControlType, elementLocator);
                    break;

                default:
                    elementFound = HandleDefaultCase(elementSrc, elementControlType, elementLocator);
                    break;
            }

            return elementFound;
        }

        internal List<IElement> HandleAllCases_multiElements(IElement elementSrc, string singleElePath)
        {
            List<IElement> elementsFound = new List<IElement>();
            var parts = singleElePath.Split('[');
            string elementControlType = parts[0];
            string elementLocator = null;
            if (singleElePath.Contains("[") || singleElePath.Contains("]"))
                elementLocator = parts[1].TrimEnd(']');

            switch (GetControlTypeCategory(elementControlType))
            {
                case PerformTypeCategory.ByClass:
                    elementsFound = HandleByClassCase_multiElements(elementSrc, elementControlType, elementLocator);
                    break;

                case PerformTypeCategory.ByCondition:
                    elementsFound = HandleByConditionCase_multiElements(elementSrc, elementControlType, elementLocator);
                    break;

                default:
                    if (elementLocator == null)
                        elementsFound = HandleDefaultCase_multiElements(elementSrc, elementControlType);
                    else
                        elementsFound.Add(HandleDefaultCase(elementSrc, elementControlType, elementLocator));
                    break;
            }
            return elementsFound;
        }

        internal IElement HandleDataGridCases(IElement elementSrc, string elementControlType, string elementLocator)
        {
            IElement elementFound;
            if (elementControlType.Contains('@'))
            {
                // Use AutomationID
                var dataGridId = elementControlType.Split('@')[1];
                var findType = elementControlType.Split('@')[0];
                //elementFound = elementSrc.GetDataGridElement(dataGridId).PerformDataGridMethod(findType, elementLocator);
                elementFound = PerformDataGridMethod(elementSrc.GetElementFromPP5ElementRetry(PP5By.Id(dataGridId)), findType, elementLocator);
            }
            else
            {
                elementFound = PerformDataGridMethod(elementSrc, elementControlType, elementLocator);
            }
            return elementFound;
        }

        internal IElement HandleByIdOrNameCase(IElement elementSrc, string elementControlType, string elementLocator)
        {
            //return elementSrc.GetElementFromWebElement(new ByAutomationIdOrName(elementLocator));
            //for (int i = 0; i < elementLocatorParts.Length; i++)
            //{
            //    if (elementControlType.Contains("#Retry"))
            //        elementFound = elementFound.GetElementWithRetry<IWebElement, IWebElement>(PP5By.IdOrName(elementLocatorParts[i])).ConvertToElement();
            //    else
            //        elementFound = elementFound.GetElement<IWebElement, IWebElement>(PP5By.IdOrName(elementLocatorParts[i])).ConvertToElement();
            //}
            //return elementFound;

            IElement elementFound = elementSrc;
            var elementLocatorParts = elementLocator.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            ByAutomationIdOrName[] IdOrNameBys = new ByAutomationIdOrName[elementLocatorParts.Length];
            for (int i = 0; i < elementLocatorParts.Length; i++)
            {
                IdOrNameBys[i] = PP5By.IdOrName(elementLocatorParts[i].Trim());
            }

            if (elementControlType.Contains("#Retry"))
                elementFound = elementFound.GetElementFromPP5ElementRetry(SharedSetting.NORMAL_TIMEOUT, IdOrNameBys);
            else
                elementFound = elementFound.GetElementFromPP5Element(SharedSetting.NORMAL_TIMEOUT, IdOrNameBys);
            return elementFound;
        }

        internal IElement HandleByNameCase(IElement elementSrc, string elementControlType, string elementLocator)
        {
            //return elementSrc.GetElementFromWebElement(By.Name(elementLocator));
            IElement elementFound = elementSrc;
            var elementLocatorParts = elementLocator.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            By[] NameBys = new By[elementLocatorParts.Length];
            for (int i = 0; i < elementLocatorParts.Length; i++)
            {
                NameBys[i] = PP5By.Name(elementLocatorParts[i].Trim());
            }

            if (elementControlType.Contains("#Retry"))
                elementFound = elementFound.GetElementWithRetry<IElement, IElement>(SharedSetting.NORMAL_TIMEOUT, NameBys);
            else
                elementFound = elementFound.GetElement<IElement, IElement>(SharedSetting.NORMAL_TIMEOUT, NameBys);
            return elementFound;

            //for (int i = 0; i < elementLocatorParts.Length; i++)
            //{
            //    if (elementControlType.Contains("#Retry"))
            //        elementFound = elementFound.GetElementWithRetry<IElement, IElement>(PP5By.Name(elementLocatorParts[i]));
            //    else
            //        elementFound = elementFound.GetElement<IElement, IElement>(PP5By.Name(elementLocatorParts[i]));
            //}
            //return elementFound;
        }

        internal IElement HandleByIdCase(IElement elementSrc, string elementControlType, string elementLocator)
        {
            //return elementSrc.GetElementFromWebElement(MobileBy.AccessibilityId(elementLocator));
            IElement elementFound = elementSrc;
            var elementLocatorParts = elementLocator.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            By[] IdBys = new By[elementLocatorParts.Length];
            for (int i = 0; i < elementLocatorParts.Length; i++)
            {
                IdBys[i] = PP5By.Id(elementLocatorParts[i].Trim());
            }

            if (elementControlType.Contains("#Retry"))
                elementFound = elementFound.GetElementWithRetry<IElement, IElement>(SharedSetting.NORMAL_TIMEOUT, IdBys);
            else
                elementFound = elementFound.GetElement<IElement, IElement>(SharedSetting.NORMAL_TIMEOUT, IdBys);
            return elementFound;

            //for (int i = 0; i < elementLocatorParts.Length; i++)
            //{
            //    if (elementControlType.Contains("#Retry"))
            //        elementFound = elementFound.GetElementWithRetry<IElement, IElement>(PP5By.Id(elementLocatorParts[i]));
            //    else
            //        elementFound = elementFound.GetElement<IElement, IElement>(PP5By.Id(elementLocatorParts[i]));
            //}
            //return elementFound;
        }

        internal IElement HandleByClassCase(IElement elementSrc, string elementControlType, string elementLocator)
        {
            IElement elementFound = elementSrc;
            var elementLocatorParts = elementLocator.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            By[] ClassNameBys = new By[elementLocatorParts.Length];
            for (int i = 0; i < elementLocatorParts.Length; i++)
            {
                ClassNameBys[i] = PP5By.ClassName(elementLocatorParts[i].Trim());
            }

            if (elementControlType.Contains("#Retry"))
                elementFound = elementFound.GetElementWithRetry<IElement, IElement>(SharedSetting.NORMAL_TIMEOUT, ClassNameBys);
            else
                elementFound = elementFound.GetElement<IElement, IElement>(SharedSetting.NORMAL_TIMEOUT, ClassNameBys);
            return elementFound;

            /*
            //for (int i = 0; i < elementLocatorParts.Length; i++)
            //{
            //    if (elementControlType.Contains("#Retry"))
            //        elementFound = elementFound.GetElementWithRetry<IElement, IElement>(PP5By.ClassName(elementLocatorParts[i]));
            //    else
            //        elementFound = elementFound.GetElement<IElement, IElement>(PP5By.ClassName(elementLocatorParts[i]));
            //}
            //return elementFound;
            */
        }

        internal IElement HandleDataGridCases(PP5Driver driver, string elementControlType, string elementLocator)
        {
            IElement elementFound = null;
            if (elementControlType.Contains('@'))
            {
                // Use AutomationID
                var dataGridId = elementControlType.Split('@')[1];
                var findType = elementControlType.Split('@')[0];
                elementFound = PerformDataGridMethod(driver.GetDataTableElement(dataGridId), findType, elementLocator);
            }
            return elementFound;
        }

        internal IElement HandleByIdOrNameCase(PP5Driver driver, string elementControlType, string elementLocator)
        {
            //if (elementControlType.Contains("#Retry"))
            //    return driver.GetExtendedElementBySingleWithRetry(new ByAutomationIdOrName(elementLocator));
            //else
            //    return driver.GetExtendedElementBySingle(new ByAutomationIdOrName(elementLocator));

            IElement elementFound;
            var elementLocatorParts = elementLocator.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            ByAutomationIdOrName[] IdOrNameBys = new ByAutomationIdOrName[elementLocatorParts.Length];
            for (int i = 0; i < elementLocatorParts.Length; i++)
            {
                IdOrNameBys[i] = PP5By.IdOrName(elementLocatorParts[i].Trim());
            }

            if (elementControlType.Contains("#Retry"))
                elementFound = driver.GetElementWithRetry<PP5Driver, IElement>(SharedSetting.NORMAL_TIMEOUT, IdOrNameBys);
            else
                elementFound = driver.GetElement<PP5Driver, IElement>(SharedSetting.NORMAL_TIMEOUT, IdOrNameBys);
            return elementFound;

            //for (int i = 0; i < elementLocatorParts.Length; i++)
            //{
            //    if (elementControlType.Contains("#Retry"))
            //        elementFound = driver.GetElementWithRetry<PP5Driver, IWebElement>(PP5By.IdOrName(elementLocatorParts[i])).ConvertToElement();
            //    else
            //        elementFound = driver.GetElement<PP5Driver, IWebElement>(PP5By.IdOrName(elementLocatorParts[i])).ConvertToElement();
            //}
            //return elementFound;
        }

        internal IElement HandleByNameCase(PP5Driver driver, string elementControlType, string elementLocator)
        {
            IElement elementFound;
            var elementLocatorParts = elementLocator.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            By[] NameBys = new By[elementLocatorParts.Length];
            for (int i = 0; i < elementLocatorParts.Length; i++)
            {
                NameBys[i] = PP5By.Name(elementLocatorParts[i].Trim());
            }

            if (elementControlType.Contains("#Retry"))
                elementFound = driver.GetElementWithRetry<PP5Driver, IElement>(SharedSetting.NORMAL_TIMEOUT, NameBys);
            else
                elementFound = driver.GetElement<PP5Driver, IElement>(SharedSetting.NORMAL_TIMEOUT, NameBys);
            return elementFound;

            //if (elementControlType.Contains("#Retry"))
            //    return driver.GetExtendedElementBySingleWithRetry(PP5By.Name(elementLocator));
            //else
            //    return driver.GetExtendedElementBySingle(PP5By.Name(elementLocator));
        }

        internal IElement HandleByIdCase(PP5Driver driver, string elementControlType, string elementLocator)
        {
            IElement elementFound;
            var elementLocatorParts = elementLocator.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            By[] IdBys = new By[elementLocatorParts.Length];
            for (int i = 0; i < elementLocatorParts.Length; i++)
            {
                IdBys[i] = PP5By.Id(elementLocatorParts[i].Trim());
            }

            if (elementControlType.Contains("#Retry"))
                elementFound = driver.GetElementWithRetry<PP5Driver, IElement>(SharedSetting.NORMAL_TIMEOUT, IdBys);
            else
                elementFound = driver.GetElement<PP5Driver, IElement>(SharedSetting.NORMAL_TIMEOUT, IdBys);
            return elementFound;

            //if (elementControlType.Contains("#Retry"))
            //    return driver.GetExtendedElementBySingleWithRetry(PP5By.Id(elementLocator));
            //else
            //    return driver.GetExtendedElementBySingle(PP5By.Id(elementLocator));
        }

        internal IElement HandleByClassCase(PP5Driver driver, string elementControlType, string elementLocator)
        {
            IElement elementFound;
            var elementLocatorParts = elementLocator.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            By[] ClassNameBys = new By[elementLocatorParts.Length];
            for (int i = 0; i < elementLocatorParts.Length; i++)
            {
                ClassNameBys[i] = PP5By.ClassName(elementLocatorParts[i].Trim());
            }

            if (elementControlType.Contains("#Retry"))
                elementFound = driver.GetElementWithRetry<PP5Driver, IElement>(SharedSetting.NORMAL_TIMEOUT, ClassNameBys);
            else
                elementFound = driver.GetElement<PP5Driver, IElement>(SharedSetting.NORMAL_TIMEOUT, ClassNameBys);
            return elementFound;

            //if (elementControlType.Contains("#Retry"))
            //    return driver.GetExtendedElementBySingleWithRetry(PP5By.ClassName(elementLocator));
            //else
            //    return driver.GetExtendedElementBySingle(PP5By.ClassName(elementLocator));
        }

        internal IElement HandleByConditionCase(IElement elementSrc, string elementControlType, string elementLocator)
        {
            IElement elementFound = null;
            if (elementControlType.Contains('#'))
            {
                Func<IElement, bool> condition = null;
                if (elementLocator == "Visible")
                {
                    condition = (e) => e.isElementVisible();
                }

                var controlType = elementControlType.Split('#')[1];
                elementFound = elementSrc.GetSpecificChildOfControlType(TypeExtension.GetEnumByDescription<ElementControlType>(controlType), condition);
            }
            return elementFound;
        }

        internal IElement HandleDefaultCase(IElement elementSrc, string elementControlType, string elementLocator)
        {
            ElementControlType eleCtrlType = TypeExtension.GetEnumByDescription<ElementControlType>(elementControlType);
            IElement elementFound;

            if (int.TryParse(elementLocator, out int elementLocatorIndex))
            {
                elementFound = elementSrc.GetSpecificChildOfControlTypeByBFS(eleCtrlType, elementLocatorIndex);
            }
            else
            {
                elementFound = elementSrc.GetSpecificChildOfControlTypeByBFS(eleCtrlType, elementLocator);
            }
            return elementFound;
        }

        internal List<IElement> HandleByClassCase_multiElements(IElement elementSrc, string elementControlType, string elementLocator)
        {
            List<IElement> elementsFound = new List<IElement>();
            var elementLocatorParts = elementLocator.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            By[] ClassNameBys = new By[elementLocatorParts.Length];
            for (int i = 0; i < elementLocatorParts.Length; i++)
            {
                ClassNameBys[i] = PP5By.ClassName(elementLocatorParts[i].Trim());
            }

            if (elementControlType.Contains("#Retry"))
                elementsFound = elementSrc.GetElementsWithRetry<IElement, IElement>(SharedSetting.NORMAL_TIMEOUT, ClassNameBys).ToList();
            else
                elementsFound = elementSrc.GetElements<IElement, IElement>(SharedSetting.NORMAL_TIMEOUT, ClassNameBys).ToList();
            return elementsFound;
        }

        internal List<IElement> HandleByConditionCase_multiElements(IElement elementSrc, string elementControlType, string elementLocator)
        {
            List<IElement> elementsFound = new List<IElement>();
            if (elementControlType.Contains('#'))
            {
                Func<IElement, bool> condition = null;
                if (elementLocator == "Visible")
                {
                    condition = (e) => e.isElementVisible();
                }

                var controlType = elementControlType.Split('#')[1];
                elementsFound = elementSrc.GetSpecificChildrenOfControlType(TypeExtension.GetEnumByDescription<ElementControlType>(controlType), condition).ToList();
            }
            return elementsFound;
        }

        internal List<IElement> HandleDefaultCase_multiElements(IElement elementSrc, string elementControlType)
        {
            ElementControlType eleCtrlType = TypeExtension.GetEnumByDescription<ElementControlType>(elementControlType);
            return elementSrc.GetSpecificChildrenOfControlTypeBFS(eleCtrlType).ToList();
        }

        internal PerformTypeCategory GetControlTypeCategory(string elementControlType)
        {
            if (IsDataGridType(elementControlType))
            {
                return PerformTypeCategory.DataGrid;
            }

            if (IsByIdOrNameType(elementControlType))
            {
                return PerformTypeCategory.ByIdOrName;
            }

            if (IsByNameType(elementControlType))
            {
                return PerformTypeCategory.ByName;
            }

            if (IsByIdType(elementControlType))
            {
                return PerformTypeCategory.ById;
            }

            if (IsByClassType(elementControlType))
            {
                return PerformTypeCategory.ByClass;
            }

            if (IsByConditionType(elementControlType))
            {
                return PerformTypeCategory.ByCondition;
            }

            return PerformTypeCategory.Default;
        }

        private bool IsDataGridType(string elementControlType) =>
            elementControlType.StartsWith("ByDataGrid", StringComparison.InvariantCultureIgnoreCase) ||
            elementControlType.StartsWith("ByCell", StringComparison.InvariantCultureIgnoreCase) ||
            elementControlType.StartsWith("ByRow", StringComparison.InvariantCultureIgnoreCase) ||
            elementControlType.StartsWith("ByColumn", StringComparison.InvariantCultureIgnoreCase);

        private bool IsByIdOrNameType(string elementControlType) =>
            elementControlType.StartsWith("ByIdOrName", StringComparison.InvariantCultureIgnoreCase) ||
            elementControlType.StartsWith("ByNameOrId", StringComparison.InvariantCultureIgnoreCase);

        private bool IsByNameType(string elementControlType) =>
            elementControlType.StartsWith("ByName", StringComparison.InvariantCultureIgnoreCase);

        private bool IsByIdType(string elementControlType) =>
            elementControlType.StartsWith("ById", StringComparison.InvariantCultureIgnoreCase);

        private bool IsByClassType(string elementControlType) =>
            elementControlType.StartsWith("ByClass", StringComparison.InvariantCultureIgnoreCase);

        private bool IsByConditionType(string elementControlType) =>
            elementControlType.StartsWith("ByCondition", StringComparison.InvariantCultureIgnoreCase);
    }

    internal enum PerformTypeCategory
    {
        DataGrid,
        ByIdOrName,
        ByName,
        ById,
        ByClass,
        ByCondition,
        Default
    }
}
