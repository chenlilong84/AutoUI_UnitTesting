using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PP5AutoUITests
{
    ///// <summary>
    ///// 定義 Test Item Variable Data type
    ///// </summary>
    //public enum VariableDataType : int
    //{
    //    [System.ComponentModel.Description("Float")]
    //    Float = 1,

    //    [System.ComponentModel.Description("Integer")]
    //    Integer = 2,

    //    [System.ComponentModel.Description("Float(%)")]
    //    FloatPercentage = 3,

    //    [System.ComponentModel.Description("Short")]
    //    Short = 4,

    //    [System.ComponentModel.Description("String")]
    //    String = 5,

    //    [System.ComponentModel.Description("Byte")]
    //    Byte = 6,

    //    [System.ComponentModel.Description("Float[L]")]
    //    FloatArrayOfSize640 = 7,

    //    [System.ComponentModel.Description("Integer[L]")]
    //    IntegerArrayOfSize640 = 8,

    //    [System.ComponentModel.Description("Float(%)[L]")]
    //    FloatPercentageArrayOfSize640 = 9,

    //    [System.ComponentModel.Description("HexString")]
    //    HexString = 10,

    //    [System.ComponentModel.Description("Float[]")]
    //    FloatArray = 11,

    //    [System.ComponentModel.Description("Integer[]")]
    //    IntegerArray = 12,

    //    [System.ComponentModel.Description("LineInVector")]
    //    LineInVector = 13,

    //    [System.ComponentModel.Description("LoadVector")]
    //    LoadVector = 14,

    //    [System.ComponentModel.Description("SpecVector")]
    //    SpecVector = 15,

    //    [System.ComponentModel.Description("ExtMeasVector")]
    //    ExtMeasVector = 16,

    //    [System.ComponentModel.Description("ACLoadVector")]
    //    ACLoadVector = 17,

    //    [System.ComponentModel.Description("ACSpecVector")]
    //    ACSpecVector = 18,

    //    [System.ComponentModel.Description("ConstantVector")]
    //    ConstantVector = 19,

    //    [System.ComponentModel.Description("Double")]
    //    Double = 20,

    //    [System.ComponentModel.Description("Double[]")]
    //    DoubleArray = 21,

    //    [System.ComponentModel.Description("Double[L]")]
    //    DoubleArrayOfUUTMaxSize = 22,

    //    [System.ComponentModel.Description("String[]")]
    //    StringArray = 23,

    //    [System.ComponentModel.Description("Byte[]")]
    //    ByteArray = 24,

    //    [System.ComponentModel.Description("Double[,]")]
    //    Double2DArray = 25,

    //    [System.ComponentModel.Description("Float[,]")]
    //    Float2DArray = 26,

    //    [System.ComponentModel.Description("Integer[,]")]
    //    Integer2DArray = 27,

    //    [System.ComponentModel.Description("Byte[,]")]
    //    Byte2DArray = 28,

    //    [System.ComponentModel.Description("String[,]")]
    //    String2DArray = 29,

    //    [System.ComponentModel.Description("String[L]")]
    //    StringArrayOfUUTMaxSize = 30,

    //    [System.ComponentModel.Description("Long")]
    //    Long = 31,

    //    [System.ComponentModel.Description("Long[]")]
    //    LongArray = 32,

    //    [System.ComponentModel.Description("HexString[]")]
    //    HexStringArray = 33,

    //    [System.ComponentModel.Description("HexString[,]")]
    //    HexString2DArray = 34,

    //    [System.ComponentModel.Description("Long[,]")]
    //    Long2DArray = 35,

    //    [System.ComponentModel.Description("Chart")]
    //    Chart = 36,

    //    [System.ComponentModel.Description("Picture")]
    //    Picture = 37,
    //}

    public enum VariableColumnType : int 
    {
        [System.ComponentModel.Description("No")]
        No = 1,
        [System.ComponentModel.Description("Show Name")]
        ShowName,
        [System.ComponentModel.Description("Call Name")]
        CallName,
        [System.ComponentModel.Description("Data Type")]
        DataType,
        [System.ComponentModel.Description("Edit Type")]
        EditType,
        [System.ComponentModel.Description("Min.")]
        Min,
        [System.ComponentModel.Description("Max.")]
        Max,
        [System.ComponentModel.Description("Default")]
        Default,
        [System.ComponentModel.Description("Format")]
        Format,
        [System.ComponentModel.Description("Enum Item")]
        EnumItem,
        [System.ComponentModel.Description("MinimumSpec")]
        MinimumSpec,
        [System.ComponentModel.Description("Defect Code(1)")]
        DefectCodeMin,
        [System.ComponentModel.Description("MaximumSpec")]
        MaximumSpec,
        [System.ComponentModel.Description("Defect Code(2)")]
        DefectCodeMax,
        [System.ComponentModel.Description("")]
        Lock,
    }

    [Flags]
    public enum VariableDataType : long
    {
        Empty = 0L,
        [System.ComponentModel.Description("Float")]
        Float = 1L,
        [System.ComponentModel.Description("Integer")]
        Integer = 2L,
        [System.ComponentModel.Description("Float(%)")]
        FloatPercentage = 4L,
        [System.ComponentModel.Description("Short")]
        Short = 8L,
        [System.ComponentModel.Description("String")]
        String = 16L,
        [System.ComponentModel.Description("Byte")]
        Byte = 32L,
        [System.ComponentModel.Description("Float[L]")]
        FloatArrayOfSize640 = 64L,
        [System.ComponentModel.Description("Integer[L]")]
        IntegerArrayOfSize640 = 128L,
        [System.ComponentModel.Description("Float(%)[L]")]
        FloatPercentageArrayOfSize640 = 256L,
        [System.ComponentModel.Description("HexString")]
        HexString = 512L,
        [System.ComponentModel.Description("Float[]")]
        FloatArray = 1024L,
        [System.ComponentModel.Description("Integer[]")]
        IntegerArray = 2048L,
        [System.ComponentModel.Description("LineInVector")]
        LineInVector = 4096L,
        [System.ComponentModel.Description("LoadVector")]
        LoadVector = 8192L,
        [System.ComponentModel.Description("SpecVector")]
        SpecVector = 16384L,
        [System.ComponentModel.Description("ExtMeasVector")]
        ExtMeasVector = 32768L,
        [System.ComponentModel.Description("ACLoadVector")]
        ACLoadVector = 65536L,
        [System.ComponentModel.Description("ACSpecVector")]
        ACSpecVector = 131072L,
        [System.ComponentModel.Description("ConstantVector")]
        ConstantVector = 262144L,
        [System.ComponentModel.Description("Double")]
        Double = 524288L,
        [System.ComponentModel.Description("Double[]")]
        DoubleArray = 1048576L,
        [System.ComponentModel.Description("Double[L]")]
        DoubleArrayOfUUTMaxSize = 2097152L,
        [System.ComponentModel.Description("String[]")]
        StringArray = 4194304L,
        [System.ComponentModel.Description("Byte[]")]
        ByteArray = 8388608L,
        [System.ComponentModel.Description("Double[,]")]
        Double2DArray = 16777216L,
        [System.ComponentModel.Description("Float[,]")]
        Float2DArray = 33554432L,
        [System.ComponentModel.Description("Integer[,]")]
        Integer2DArray = 67108864L,
        [System.ComponentModel.Description("Byte[,]")]
        Byte2DArray = 134217728L,
        [System.ComponentModel.Description("String[,]")]
        String2DArray = 268435456L,
        [System.ComponentModel.Description("String[L]")]
        StringArrayOfUUTMaxSize = 536870912L,
        [System.ComponentModel.Description("Long")]
        Long = 1073741824L,
        [System.ComponentModel.Description("Long[]")]
        LongArray = 2147483648L,
        [System.ComponentModel.Description("HexString[]")]
        HexStringArray = 4294967296L,
        [System.ComponentModel.Description("HexString[,]")]
        HexString2DArray = 8589934592L,
        [System.ComponentModel.Description("Long[,]")]
        Long2DArray = 17179869184L,
        [System.ComponentModel.Description("Chart")]
        Chart = 34359738368L,
        [System.ComponentModel.Description("Picture")]
        Picture = 68719476736L,
        [System.ComponentModel.Description("StringCategory")]
        StringCtgy = VariableDataType.String | VariableDataType.StringArrayOfUUTMaxSize | VariableDataType.StringArray | VariableDataType.String2DArray,
        [System.ComponentModel.Description("HexStringCategory")]
        HexStringCtgy = VariableDataType.HexString | VariableDataType.HexStringArray | VariableDataType.HexString2DArray,
    }

    /// <summary>
    /// 定義 Test Item Variable Edit type
    /// </summary>
    public enum VariableEditType
    {
        Empty = 0,
        EditBox = 1,
        ComboBox = 2,
        External_Signal = 3,
    }

    /// <summary>
    /// 定義 Variable Format 種類
    /// </summary>
    public enum VariableDecimalPlaces
    {
        Empty = 0,
        [System.ComponentModel.Description(".0")]
        Zero = 1,
        [System.ComponentModel.Description(".1")]
        One = 2,
        [System.ComponentModel.Description(".2")]
        Two = 3,
        [System.ComponentModel.Description(".3")]
        Three = 4,
        [System.ComponentModel.Description(".4")]
        Four = 5,
        [System.ComponentModel.Description(".5")]
        Five = 6,
        [System.ComponentModel.Description(".6")]
        Six = 7
    }
}
