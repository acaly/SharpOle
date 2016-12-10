using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace SharpOle.OleInterop
{
    [Guid("0000000E-0000-0000-C000-000000000046"), InterfaceType(1)]
    [ComImport]
    public interface IBindCtx
    {
        void RegisterObjectBound(
            [MarshalAs(UnmanagedType.IUnknown)] [In] object punk);
        void RevokeObjectBound(
            [MarshalAs(UnmanagedType.IUnknown)] [In] object punk);
        void ReleaseBoundObjects();
        void SetBindOptions(
            [In] IntPtr pbindopts);
        void GetBindOptions(
            [In] [Out] IntPtr pbindopts);
        void GetRunningObjectTable(
            [MarshalAs(UnmanagedType.Interface)] out IRunningObjectTable pprot);
        void RegisterObjectParam(
            [MarshalAs(UnmanagedType.LPWStr)] [In] string pszKey,
            [MarshalAs(UnmanagedType.IUnknown)] [In] object punk);
        void GetObjectParam(
            [MarshalAs(UnmanagedType.LPWStr)] [In] string pszKey,
            [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);
        void EnumObjectParam(
            [MarshalAs(UnmanagedType.Interface)] out IEnumString ppEnum);
        void RevokeObjectParam(
            [MarshalAs(UnmanagedType.LPWStr)] [In] string pszKey);
    }
}
