using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PP5AutoUITests
{
    /// <summary>
    /// 定義 TI Editor Toolbar 的功能按鈕種類
    /// </summary>
    public enum TIToolBarButton : int
    {
        TIEditor = 0,
        Execution = 1,
        Save = 3,
        SaveAs = 4,
        Undo = 6, 
        Redo = 7,
        Cut = 8,
        Insert = 9,
        Copy = 10,
        Paste = 11,
        Delete = 12,
    }

    /// <summary>
    /// 定義 Hardware Configuration Toolbar 的功能按鈕種類
    /// </summary>
    public enum HWConfToolBarButton : int
    {
        LoadConnection = 0,
        SetAsSysDefault = 1,
        Save = 2,
        SaveAs = 3,
        Undo = 5,
        Redo = 6,
        Delete = 7,
        Test = 8,
    }

    /// <summary>
    /// 定義 Management Toolbar 的功能按鈕種類
    /// </summary>
    public enum MngtToolBarButton : int
    {
        Security = 0,
        TPTI,
        SystemSetup,
        MISC,
        ExFuncion
    }
}
