using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PP5AutoUITests
{
    /// <summary>
    /// 定義 Auto UI 動作種類
    /// </summary>
    public enum ActionType : int
    {
        /// <summary>
        /// No action, only get element
        /// </summary>
        None,
        /// <summary>
        /// 滑鼠左鍵單擊 (left mouse button single click)
        /// </summary>
        LeftClick = 1,
        /// <summary>
        /// 滑鼠左鍵雙擊 (left mouse button double click)
        /// </summary>
        LeftDoubleClick = 2,
        /// <summary>
        /// 鍵盤輸入
        /// </summary>
        SendKeys = 3,
    }

    public enum ClickType : int
    {
        /// <summary>
        /// No action, only get element
        /// </summary>
        None,
        /// <summary>
        /// 滑鼠左鍵單擊 (left mouse button single click)
        /// </summary>
        LeftClick = 1,
        /// <summary>
        /// 滑鼠左鍵雙擊 (left mouse button double click)
        /// </summary>
        LeftDoubleClick = 2,
        /// <summary>
        /// 滑鼠右鍵單擊 (right mouse button single click)
        /// </summary>
        RightClick = 3,
        /// <summary>
        /// 滑鼠右鍵雙擊 (right mouse button double click)
        /// </summary>
        RightDoubleClick = 4,
        /// <summary>
        /// 滑鼠左鍵延遲雙擊 (left mouse button double click with delay)
        /// </summary>
        LeftDoubleClickDelay = 5,
        /// <summary>
        /// Tick Checkbox  (using left mouse button)
        /// </summary>
        TickCheckBox = 6,
        /// <summary>
        /// UnTick Checkbox (using left mouse button)
        /// </summary>
        UnTickCheckBox = 7,
    }

    public enum InputType : int
    {
        /// <summary>
        /// No action, only get element
        /// </summary>
        None,
        /// <summary>
        /// 單一鍵盤輸入
        /// </summary>
        SendSingleKeys = 1,
        /// <summary>
        /// 多個鍵盤輸入
        /// </summary>
        SendComboKeys = 2,
        /// <summary>
        /// 清空欄位後再做鍵盤輸入
        /// </summary>
        SendContent = 3,
        /// <summary>
        /// 欄位內全選 (操作同 Ctrl + A)
        /// </summary>
        SelectAllContent = 4,
        /// <summary>
        /// 欄位內全選後清空欄位
        /// </summary>
        ClearContent = 5,
        /// <summary>
        /// 將剪貼簿的內容貼到欄位內 (操作同 Ctrl + V)
        /// </summary>
        PasteContent = 6,
        /// <summary>
        /// 將欄位內全選後複製到剪貼簿 (SelectAllContent 與 Ctrl + C 混和動作)
        /// </summary>
        CopyContent = 7,
        /// <summary>
        /// 將欄位內全選後剪下 (SelectAllContent 與 Ctrl + X 混和動作)
        /// </summary>
        CutContent = 8,
        /// <summary>
        /// 來源欄位複製後，貼到目標欄位 (CopyContent + PasteContent 混和動作)
        /// </summary>
        CopyAndPaste = 9,
        /// <summary>
        /// 來源欄位剪下後，貼到目標欄位 (CutContent + PasteContent 混和動作)
        /// </summary>
        CutAndPaste = 10,
    }

    /// <summary>
    /// 定義Test Program 動作種類
    /// </summary>
    public enum TPAction : int
    {
        /// <summary>
        /// Switch to page: Pre Test
        /// </summary>
        SwitchToPreTestPage = 0,
        /// <summary>
        /// Switch to page: UUT Test
        /// </summary>
        SwitchToUUTTestPage,
        /// <summary>
        /// Switch to page: Post Test
        /// </summary>
        SwitchToPostTestPage,
        /// <summary>
        /// Switch to page: System TI
        /// </summary>
        SwitchToSystemTIPage,
        /// <summary>
        /// Switch to page: User-Defined TI
        /// </summary>
        SwitchToUserDefinedTIPage,
        /// <summary>
        /// Switch to page: Test item > Parameter (condition)
        /// </summary>
        SwitchToTestConditionVariablePage,
        /// <summary>
        /// Switch to page: Test item > Vector
        /// </summary>
        SwitchToVectorVariablePage,
        /// <summary>
        /// Switch to page: Test item > Global
        /// </summary>
        SwitchToGlobalVariablePage,
        /// <summary>
        /// Switch to page: Test item > Result
        /// </summary>
        SwitchToResultVariablePage,
        /// <summary>
        /// Switch to page: TPInfo
        /// </summary>
        SwitchToTPInfoPage,
        /// <summary>
        /// Switch to page: Report Format > ByTI
        /// </summary>
        SwitchToReportFormatByTIPage,
        /// <summary>
        /// Switch to page: Report Format > ByTP
        /// </summary>
        SwitchToReportFormatByTPPage,
    }

    /// <summary>
    /// 定義Test Item 動作種類
    /// </summary>
    public enum TIAction : int
    {
        /// <summary>
        /// Switch to Management and set TI active
        /// </summary>
        SetTIActive = 0,
        /// <summary>
        /// Switch to Management and set TI inactive
        /// </summary>
        SetTIInactive = 1,
        ///// <summary>
        ///// Switch to Management and modify edit: Remark
        ///// </summary>
        //TIModifyRemark = 2,
    }
}
