using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PP5AutoUITests
{
    public static class TIEditorHelper
    {
        public static bool IsVariableUUTArrayDataType(VariableDataType dataType)
        {
            VariableDataType DataTypeUUTArray = VariableDataType.FloatArrayOfSize640 | VariableDataType.IntegerArrayOfSize640 |
                                                VariableDataType.FloatPercentageArrayOfSize640 | VariableDataType.DoubleArrayOfUUTMaxSize |
                                                VariableDataType.StringArrayOfUUTMaxSize;
            return dataType == (dataType & DataTypeUUTArray);
        }

        public static bool IsVariableArrayDataType(VariableDataType dataType)
        {
            return IsVariableUUTArrayDataType(dataType) || IsVariable1DArrayDataType(dataType) || IsVariable2DArrayDataType(dataType);
        }

        public static bool IsVariable1DArrayDataType(VariableDataType dataType)
        {
            VariableDataType DataType1DArray = VariableDataType.FloatArray | VariableDataType.IntegerArray |
                                   VariableDataType.DoubleArray | VariableDataType.StringArray |
                                   VariableDataType.ByteArray | VariableDataType.HexStringArray | VariableDataType.LongArray;
            return dataType == (dataType & DataType1DArray);
        }

        public static bool IsVariable2DArrayDataType(VariableDataType dataType)
        {
            VariableDataType DataType2DArray = VariableDataType.Float2DArray | VariableDataType.Integer2DArray |
                                   VariableDataType.Double2DArray | VariableDataType.String2DArray |
                                   VariableDataType.Byte2DArray | VariableDataType.HexString2DArray | VariableDataType.Long2DArray;
            return dataType == (dataType & DataType2DArray);
        }

        public static bool IsVariableVectorDataType(VariableDataType dataType)
        {
            VariableDataType DataTypeVector = VariableDataType.LineInVector | VariableDataType.SpecVector |
                                               VariableDataType.ACSpecVector | VariableDataType.ExtMeasVector |
                                               VariableDataType.LoadVector | VariableDataType.ACLoadVector | VariableDataType.ConstantVector;
            return dataType == (dataType & DataTypeVector);
        }

        public static bool IsVariableStringDataType(VariableDataType dataType)
        {
            VariableDataType DataTypeString = VariableDataType.String | VariableDataType.StringArray | VariableDataType.String2DArray;
            return dataType == (dataType & DataTypeString);
        }

        public static bool IsVariableHexStringDataType(VariableDataType dataType)
        {
            VariableDataType DataTypeHexString = VariableDataType.HexString | VariableDataType.HexStringArray | VariableDataType.HexString2DArray;
            return dataType == (dataType & DataTypeHexString);
        }

        public static bool IsVariableByteArrayDataType(VariableDataType dataType)
        {
            VariableDataType DataTypeByteArray = VariableDataType.ByteArray | VariableDataType.Byte2DArray;
            return dataType == (dataType & DataTypeByteArray);
        }

        public static bool IsVariableFloatType(VariableDataType dataType)
        {
            VariableDataType DataTypeFloat = VariableDataType.Float | VariableDataType.FloatPercentage | VariableDataType.Double |
                                             VariableDataType.FloatPercentageArrayOfSize640 | VariableDataType.DoubleArrayOfUUTMaxSize | VariableDataType.FloatArrayOfSize640;
            return dataType == (dataType & DataTypeFloat);
        }
    }
}
