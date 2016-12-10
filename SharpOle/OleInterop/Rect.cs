using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace SharpOle.OleInterop
{
    [StructLayout(LayoutKind.Sequential, Pack = 4)]
    public struct Rect
    {
        public int left;
        public int top;
        public int right;
        public int bottom;
    }

}
