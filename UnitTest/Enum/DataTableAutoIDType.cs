using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PP5AutoUITests
{
    /// <summary>
    /// Combined enum class for all tabs
    /// </summary>
    public enum TIDataTableAutoIDType
    {
        CndGrid = VariableTabType.Condition,
        RstGrid = VariableTabType.Result,
        TmpGrid = VariableTabType.Temp,
        GlbGrid = VariableTabType.Global,
        DefectCodeGrid = VariableTabType.DefectCode,
        PGGrid = TestItemTabType.TIContext,
        ParameterGrid = 6,
        LoginGrid = 7,
    }

    public enum MngtDataTableAutoIDType
    {
        [System.ComponentModel.Description("Test Item")]
        TestItem_DataGrid = 0,
        [System.ComponentModel.Description("Test Program")]
        TestProgram_DataGrid = 1,
        [System.ComponentModel.Description("DLL")]
        DllDataGrid = 2,
        [System.ComponentModel.Description("Python")]
        PythonDataGrid = 3,
        [System.ComponentModel.Description("Log Data")]
        logDataGrid = 4,
    }
}
