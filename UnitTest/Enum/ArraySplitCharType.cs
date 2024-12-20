using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PP5AutoUITests
{
    public enum Rank1SplitChar
    {
        [System.ComponentModel.Description("Tab")]
        Tab,
        [System.ComponentModel.Description("Comma(,)")]
        Comma,
        [System.ComponentModel.Description("Semicolon(;")]
        Semicolon,
        [System.ComponentModel.Description("Space")]
        Space,
        [System.ComponentModel.Description("NewLine")]
        NewLine,
        [System.ComponentModel.Description("Custom")]
        Custom
    }

    public enum Rank2SplitChar
    {
        [System.ComponentModel.Description("Semicolon(;")]
        Semicolon,
        [System.ComponentModel.Description("NewLine")]
        NewLine,
        [System.ComponentModel.Description("Custom")]
        Custom
    }
}
