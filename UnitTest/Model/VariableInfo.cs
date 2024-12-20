using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms.VisualStyles;
using static PP5AutoUITests.TIEditorHelper;

namespace PP5AutoUITests
{
    // Define a class that encapsulates the variable information.
    public class VariableInfo
    {
        public VariableTabType TabType { get; set; }
        public string CallName { get; set; }
        public VariableDataType DataType { get; set; }
        public VariableEditType EditType { get; set; }
        public int ArrSize1 { get; set; }
        public int ArrSize2 { get; set; }

        // Constructor to easily create a VariableInfo object
        public VariableInfo(VariableTabType tabType, string callName, VariableDataType dataType = VariableDataType.Empty,
                            VariableEditType editType = VariableEditType.Empty, int arrSize1 = 0, int arrSize2 = 0)
        {
            TabType = tabType;
            CallName = callName;
            DataType = dataType;
            EditType = editType;
            ArrSize1 = arrSize1;
            ArrSize2 = arrSize2;
        }
    }

    public class VariableRowInfoSets
    {
        public VariableRowInfoSets() 
        {
            variableRowInfoSets = new List<VariableRowInfoSet>();
        }

        public List<VariableRowInfoSet> variableRowInfoSets { get; set; }

        public bool Remove(VariableTabType tabType, string callName)
        {
            bool removeSuccess = false;
            if (variableRowInfoSets.Count > 0)
                removeSuccess = variableRowInfoSets.Remove(variableRowInfoSets.Find(vris => vris.tabType == tabType && vris.variableRowInfos.ContainsKey(callName)));

            return removeSuccess;
        }

        public VariableRowInfo GetVariableRowInfoSet(VariableTabType tabType, string callName)
        {
            var VRISs = variableRowInfoSets.FirstOrDefault(vris => vris.tabType == tabType);

            if (VRISs != null && VRISs.variableRowInfos.ContainsKey(callName))
                return VRISs.variableRowInfos[callName];

            return null;
        }
    }

    public class VariableRowInfoSet
    {
        public VariableRowInfoSet(VariableTabType _tabType) 
        {
            variableRowInfos = new Dictionary<string, VariableRowInfo>();
            tabType = _tabType;
        }
        public Dictionary<string, VariableRowInfo> variableRowInfos { get; set; }
        public VariableTabType tabType { get; set; }

        public VariableRowInfoSet Clone()
        {
            VariableRowInfoSet _rowInfoSet = new VariableRowInfoSet(tabType);
            Dictionary<string, VariableRowInfo> _variableRowInfos = new Dictionary<string, VariableRowInfo>();
            foreach (KeyValuePair<string, VariableRowInfo> vri in variableRowInfos)
            {
                _variableRowInfos.Add(vri.Key, vri.Value.Clone());
            }
            _rowInfoSet.variableRowInfos = _variableRowInfos;
            return _rowInfoSet;
        }
    }

    public class VariableRowInfo
    {
        public VariableRowInfo(DataTypeInfo _dataTypeInfo, VariableEditType _editType)
        {
            dataTypeInfo = _dataTypeInfo;
            editType = _editType;
        }
        public DataTypeInfo dataTypeInfo { get; set; }
        public VariableEditType editType { get; set; }
        
        // can set HexString by HexString Editor window
        public bool CanSetHexString => IsVariableHexStringDataType(dataTypeInfo.dataType) && editType == VariableEditType.EditBox;

        // can set External Signal by External Signal popup window
        public bool CanSetExternalSignal => IsVariableStringDataType(dataTypeInfo.dataType) && editType == VariableEditType.External_Signal;

        public VariableRowInfo Clone()
        {
            return new VariableRowInfo(dataTypeInfo.Clone(), editType);
        }
    }

    public class DataTypeInfo
    {
        public DataTypeInfo(VariableDataType _dataType, int _arrSize1, int _arrSize2)
        {
            dataType = _dataType;
            arrSize1 = _arrSize1;
            arrSize2 = _arrSize2;
            if (IsVariableArrayDataType(dataType))
            {
                arrSize1 = 0;
                arrSize2 = 0;
            }
        }

        public DataTypeInfo(VariableDataType _dataType)
        {
            dataType = _dataType;
            if (IsVariableArrayDataType(dataType))
            {
                arrSize1 = 0;
                arrSize2 = 0;
            }
        }

        public VariableDataType dataType { get; set; }
        public int arrSize1 { get; set; }
        public int arrSize2 { get; set; }

        public DataTypeInfo Clone()
        {
            return new DataTypeInfo(dataType, arrSize1, arrSize2);
        }
    }
}
