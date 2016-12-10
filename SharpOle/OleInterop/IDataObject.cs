using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace SharpOle.OleInterop
{
    [Guid("0000010E-0000-0000-C000-000000000046"), InterfaceType(1)]
    [ComImport]
    public interface IDataObject
    {
        [PreserveSig]
        int GetData(
            [MarshalAs(UnmanagedType.LPArray)] [In] FormatEtc[] pformatetcIn,
            [MarshalAs(UnmanagedType.LPArray)] [Out] StgMedium[] pRemoteMedium);
        [PreserveSig]
        int GetDataHere(
            [MarshalAs(UnmanagedType.LPArray)] [In] FormatEtc[] pFormatetc,
            [MarshalAs(UnmanagedType.LPArray)] [In] [Out] StgMedium[] pRemoteMedium);
        [PreserveSig]
        int QueryGetData(
            [MarshalAs(UnmanagedType.LPArray)] [In] FormatEtc[] pFormatetc);
        [PreserveSig]
        int GetCanonicalFormatEtc(
            [MarshalAs(UnmanagedType.LPArray)] [In] FormatEtc[] pformatectIn,
            [MarshalAs(UnmanagedType.LPArray)] [Out] FormatEtc[] pformatetcOut);
        [PreserveSig]
        int SetData(
            [MarshalAs(UnmanagedType.LPArray)] [In] FormatEtc[] pFormatetc,
            [MarshalAs(UnmanagedType.LPArray)] [In] StgMedium[] pmedium,
            [In] int fRelease);
        [PreserveSig]
        int EnumFormatEtc(
            [In] uint dwDirection,
            [MarshalAs(UnmanagedType.Interface)] out IEnumFormatEtc ppenumFormatEtc);
        [PreserveSig]
        int DAdvise(
            [MarshalAs(UnmanagedType.LPArray)] [In] FormatEtc[] pFormatetc,
            [In] uint ADVF,
            [MarshalAs(UnmanagedType.Interface)] [In] IAdviseSink pAdvSink,
            out uint pdwConnection);
        [PreserveSig]
        int DUnadvise(
            [In] uint dwConnection);
        [PreserveSig]
        int EnumDAdvise(
            [MarshalAs(UnmanagedType.Interface)] out IEnumStatData ppenumAdvise);
    }
}
