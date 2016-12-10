using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace SharpOle.OleInterop
{
    [Guid("00000010-0000-0000-C000-000000000046"), InterfaceType(1)]
    [ComImport]
    public interface IRunningObjectTable
    {
        void Register(
            [In] uint grfFlags,
            [MarshalAs(UnmanagedType.IUnknown)] [In] object punkObject,
            [MarshalAs(UnmanagedType.Interface)] [In] IMoniker pmkObjectName,
            out uint pdwRegister);
        void Revoke(
            [In] uint dwRegister);
        [PreserveSig]
        int IsRunning(
            [MarshalAs(UnmanagedType.Interface)] [In] IMoniker pmkObjectName);
        void GetObject(
            [MarshalAs(UnmanagedType.Interface)] [In] IMoniker pmkObjectName,
            [MarshalAs(UnmanagedType.IUnknown)] out object ppunkObject);
        void NoteChangeTime(
            [In] uint dwRegister,
            [MarshalAs(UnmanagedType.LPArray)] [In] FileTime[] pFileTime);
        void GetTimeOfLastChange(
            [MarshalAs(UnmanagedType.Interface)] [In] IMoniker pmkObjectName,
            [MarshalAs(UnmanagedType.LPArray)] [Out] FileTime[] pFileTime);
        void EnumRunning(
            [MarshalAs(UnmanagedType.Interface)] out IEnumMoniker ppenumMoniker);
    }
}
