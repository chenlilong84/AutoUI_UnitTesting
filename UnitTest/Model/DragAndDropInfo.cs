using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PP5AutoUITests.Model
{
    public class DragAndDropInfo
    {
        public DragAndDropInfo(MoveToElementOffsetStartingPoint _startingPointType, int _srcOffsetX, int _srcOffsetY, int _destOffsetX, int _destOffsetY) 
        {
            startingPointType = _startingPointType;
            srcOffsetX = _srcOffsetX;
            srcOffsetY = _srcOffsetY;
            destOffsetX = _destOffsetX;
            destOffsetY = _destOffsetY;
        }
        public MoveToElementOffsetStartingPoint startingPointType { get; set; }
        public int srcOffsetX { get ; set; }
        public int srcOffsetY { get ; set; }
        public int destOffsetX { get ; set; }
        public int destOffsetY { get ; set; }
    }
}
