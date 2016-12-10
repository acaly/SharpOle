using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace SharpOle.OleInterop
{
    [Guid("0000010F-0000-0000-C000-000000000046"), InterfaceType(1)]
    [ComImport]
    public interface IAdviseSink
    {
        void OnDataChange(
            [MarshalAs(UnmanagedType.LPArray)] [In] FormatEtc[] pFormatetc,
            [MarshalAs(UnmanagedType.LPArray)] [In] StgMedium[] pStgmed);
        void OnViewChange(
            [In] uint dwAspect,
            [In] int lindex);
        void OnRename(
            [MarshalAs(UnmanagedType.Interface)] [In] IMoniker pmk);
        void OnSave();
        void OnClose();
    }
}
