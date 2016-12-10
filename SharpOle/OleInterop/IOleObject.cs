using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace SharpOle.OleInterop
{
    [ComConversionLoss, Guid("00000112-0000-0000-C000-000000000046"), InterfaceType(1)]
    [ComImport]
    public interface IOleObject
    {
        void SetClientSite(
            [MarshalAs(UnmanagedType.Interface)] [In] IOleClientSite pClientSite);
        void GetClientSite(
            [MarshalAs(UnmanagedType.Interface)] out IOleClientSite ppClientSite);
        void SetHostNames(
            [MarshalAs(UnmanagedType.LPWStr)] [In] string szContainerApp,
            [MarshalAs(UnmanagedType.LPWStr)] [In] string szContainerObj);
        void Close(
            [In] uint dwSaveOption);
        void SetMoniker(
            [In] uint dwWhichMoniker,
            [MarshalAs(UnmanagedType.Interface)] [In] IMoniker pmk);
        void GetMoniker(
            [In] uint dwAssign,
            [In] uint dwWhichMoniker,
            [MarshalAs(UnmanagedType.Interface)] out IMoniker ppmk);
        [PreserveSig]
        int InitFromData(
            [MarshalAs(UnmanagedType.Interface)] [In] IDataObject pDataObject,
            [In] int fCreation,
            [In] uint dwReserved);
        void GetClipboardData(
            [In] uint dwReserved,
            [MarshalAs(UnmanagedType.Interface)] out IDataObject ppDataObject);
        [PreserveSig]
        int DoVerb(
            [In] int iVerb,
            [MarshalAs(UnmanagedType.LPArray)] [In] Msg[] lpmsg,
            [MarshalAs(UnmanagedType.Interface)] [In] IOleClientSite pActiveSite,
            [In] int lindex,
            [In] IntPtr hWndParent,
            [MarshalAs(UnmanagedType.LPArray)] [In] Rect[] lprcPosRect);
        [PreserveSig]
        int EnumVerbs(
            [MarshalAs(UnmanagedType.Interface)] out IEnumOleVerb ppEnumOleVerb);
        [PreserveSig]
        int Update();
        [PreserveSig]
        int IsUpToDate();
        void GetUserClassID(
            out Guid pClsid);
        [PreserveSig]
        int GetUserType(
            [In] uint dwFormOfType,
            [In] IntPtr pszUserType);
        void SetExtent(
            [In] uint dwDrawAspect,
            [MarshalAs(UnmanagedType.LPArray)] [In] SizeL[] pSizel);
        void GetExtent(
            [In] uint dwDrawAspect,
            [MarshalAs(UnmanagedType.LPArray)] [Out] SizeL[] pSizel);
        void Advise(
            [MarshalAs(UnmanagedType.Interface)] [In] IAdviseSink pAdvSink,
            out uint pdwConnection);
        void Unadvise(
            [In] uint dwConnection);
        void EnumAdvise(
            [MarshalAs(UnmanagedType.Interface)] out IEnumStatData ppenumAdvise);
        [PreserveSig]
        int GetMiscStatus(
            [In] uint dwAspect,
            out uint pdwStatus);
        void SetColorScheme(
            [MarshalAs(UnmanagedType.LPArray)] [In] LogPalette[] pLogpal);
    }

}
