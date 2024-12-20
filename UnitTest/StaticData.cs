using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PP5AutoUITests
{
    public static class StaticData
    {
        public static PP5DataGrid currentDataGrid { get; set; }
        public static PP5Element currentElement { get; set; }
        //public static VariableRowInfoSet conditionVariables { get; set; }
        //public static VariableRowInfoSet resultVariables { get; set; }
        //public static VariableRowInfoSet tempVariables { get; set; }
        //public static VariableRowInfoSet globalVariables { get; set; }
        public static VariableRowInfoSets TiVariables { get; set; }
        public static bool isHexString { get; set; }
        public static bool isExternalSignal { get; set; }
    }
}
