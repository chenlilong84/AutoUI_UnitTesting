using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PP5AutoUITests
{
    //
    // 摘要:
    //     Provides values that indicate from where element offsets for MoveToElement are
    //     calculated.
    public enum MoveToElementOffsetStartingPoint
    {
        //
        // 摘要:
        //     Offsets are calculated from the top-left corner of the element.
        TopLeft = 0,
        //
        // 摘要:
        //     Offsets are calcuated from the center-top of the element.
        CenterTop,
        //
        // 摘要:
        //     Offsets are calcuated from the top-right corner of the element.
        TopRight,
        //
        // 摘要:
        //     Offsets are calcuated from the center-left of the element.
        CenterLeft,
        //
        // 摘要:
        //     Offsets are calcuated from the center of the element.
        Center,
        //
        // 摘要:
        //     Offsets are calcuated from the center-right of the element.
        CenterRight,
        //
        // 摘要:
        //     Offsets are calcuated from the bottom-left corner of the element.
        BottomLeft,
        //
        // 摘要:
        //     Offsets are calcuated from the center-bottom of the element.
        CenterBottom,
        //
        // 摘要:
        //     Offsets are calcuated from the bottom-right corner of the element.
        BottomRight,
        //
        // 摘要:
        //     Offsets are calculated from the inner top-left corner of the element.
        InnerTopLeft,
        //
        // 摘要:
        //     Offsets are calcuated from the inner center-top position of the element.
        InnerCenterTop,
        //
        // 摘要:
        //     Offsets are calcuated from the inner top-right corner of the element.
        InnerTopRight,
        //
        // 摘要:
        //     Offsets are calcuated from the inner center-left position of the element.
        InnerCenterLeft,
        //
        // 摘要:
        //     Offsets are calcuated from the inner center-right position of the element.
        InnerCenterRight,
        //
        // 摘要:
        //     Offsets are calcuated from the inner bottom-left corner of the element.
        InnerBottomLeft,
        //
        // 摘要:
        //     Offsets are calcuated from the inner center-bottom position of the element.
        InnerCenterBottom,
        //
        // 摘要:
        //     Offsets are calcuated from the inner bottom-right corner of the element.
        InnerBottomRight,
        //
        // 摘要:
        //     Offsets are calculated from the outer top-left corner of the element.
        OuterTopLeft,
        //
        // 摘要:
        //     Offsets are calcuated from the outer center-top position of the element.
        OuterCenterTop,
        //
        // 摘要:
        //     Offsets are calcuated from the outer top-right corner of the element.
        OuterTopRight,
        //
        // 摘要:
        //     Offsets are calcuated from the outer center-left position of the element.
        OuterCenterLeft,
        //
        // 摘要:
        //     Offsets are calcuated from the outer center-right position of the element.
        OuterCenterRight,
        //
        // 摘要:
        //     Offsets are calcuated from the outer bottom-left corner of the element.
        OuterBottomLeft,
        //
        // 摘要:
        //     Offsets are calcuated from the outer center-bottom position of the element.
        OuterCenterBottom,
        //
        // 摘要:
        //     Offsets are calcuated from the outer bottom-right corner of the element.
        OuterBottomRight
    }
}
