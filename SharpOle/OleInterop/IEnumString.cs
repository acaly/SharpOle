using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace SharpOle.OleInterop
{
    [Guid("00000101-0000-0000-C000-000000000046"), InterfaceType(1)]
    [ComImport]
    public interface IEnumString
    {
        [PreserveSig]
        int Next(
            [In] uint celt,
            [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr, SizeParamIndex = 0)] [Out] string[] rgelt,
            [MarshalAs(UnmanagedType.LPArray)] [Out] uint[] pceltFetched);
        [PreserveSig]
        int Skip(
            [In] uint celt);
        [PreserveSig]
        int Reset();
        [PreserveSig]
        int Clone(
            [MarshalAs(UnmanagedType.Interface)] out IEnumString ppEnum);
    }

}
