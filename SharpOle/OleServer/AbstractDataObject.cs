using SharpOle.OleInterop;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace SharpOle.OleServer
{
    public abstract class AbstractDataObject : IDataObject
    {
        public AbstractDataObject() { }

        public abstract Guid GetGuid();

        internal AbstractOleObject OleObject { get; set; }

        internal void SetClipboard()
        {
            Natives.OleSetClipboard(this);
        }

        int IDataObject.DAdvise(FormatEtc[] pFormatetc, uint ADVF, IAdviseSink pAdvSink, out uint pdwConnection)
        {
            pdwConnection = 0;
            return Natives.E_NOTIMPL;
        }

        int IDataObject.DUnadvise(uint dwConnection)
        {
            return Natives.E_NOTIMPL;
        }

        int IDataObject.EnumDAdvise(out IEnumStatData ppenumAdvise)
        {
            ppenumAdvise = null;
            return Natives.E_NOTIMPL;
        }

        int IDataObject.EnumFormatEtc(uint dwDirection, out IEnumFormatEtc ppenumFormatEtc)
        {
            if (dwDirection == 1/*DATADIR_GET*/)
            {
                ppenumFormatEtc = new SimpleEnumFormatEtc();
                return 0;
            }
            ppenumFormatEtc = null;
            return Natives.E_NOTIMPL;
        }

        int IDataObject.GetCanonicalFormatEtc(FormatEtc[] pformatectIn, FormatEtc[] pformatetcOut)
        {
            pformatetcOut[0].ptd = IntPtr.Zero;
            return Natives.E_NOTIMPL;
        }

        int IDataObject.GetData(FormatEtc[] pformatetcIn, StgMedium[] pRemoteMedium)
        {
            switch (pformatetcIn[0].cfFormat)
            {
                case 0xC00E: //des
                    if (pformatetcIn[0].tymed != 1)
                    {
                        return Natives.DV_E_FORMATETC;
                    }
                    pRemoteMedium[0].pUnkForRelease = IntPtr.Zero;
                    pRemoteMedium[0].tymed = 1;
                    pRemoteMedium[0].unionmember = GetObjectDescriptor();
                    break;
                case 0xC00B: //embed
                    if (pformatetcIn[0].tymed != 8)
                    {
                        return Natives.DV_E_FORMATETC;
                    }
                    pRemoteMedium[0].pUnkForRelease = IntPtr.Zero;
                    pRemoteMedium[0].tymed = 8;
                    pRemoteMedium[0].unionmember = GetSourceStoragePtr();
                    break;
                case 3: //CF_METAFILEPICT
                    if (pformatetcIn[0].tymed != 32)
                    {
                        return Natives.DV_E_FORMATETC;
                    }
                    pRemoteMedium[0].pUnkForRelease = IntPtr.Zero;
                    pRemoteMedium[0].tymed = 32;
                    pRemoteMedium[0].unionmember = GetImage();
                    break;
                default:
                    return Natives.DV_E_FORMATETC;
            }
            return 0;
        }

        int IDataObject.GetDataHere(FormatEtc[] pFormatetc, StgMedium[] pRemoteMedium)
        {
            switch (pFormatetc[0].cfFormat)
            {
                case 0xC00E: //des
                    return Natives.DV_E_FORMATETC;
                case 0xC00B: //embed
                    if (pFormatetc[0].tymed != 8 || pRemoteMedium[0].tymed != 8)
                    {
                        return Natives.DV_E_FORMATETC;
                    }
                    var dest = pRemoteMedium[0].unionmember;
                    var destStg = (IStorage)Marshal.GetObjectForIUnknown(dest);
                    GetSourceStorage().CopyTo(0, null, IntPtr.Zero, destStg);
                    break;
                default:
                    return Natives.DV_E_FORMATETC;
            }
            return 0;
        }

        int IDataObject.QueryGetData(FormatEtc[] pFormatetc)
        {
            switch (pFormatetc[0].cfFormat)
            {
                case 0xC00E: //des
                case 0xC00B: //embed
                case 3:
                    return 0;
                default:
                    return Natives.DV_E_FORMATETC;
            }
        }

        int IDataObject.SetData(FormatEtc[] pFormatetc, StgMedium[] pmedium, int fRelease)
        {
            return Natives.E_NOTIMPL;
        }

        private IntPtr GetObjectDescriptor()
        {
            return GetObjectDescriptor("Test Object", "Test Application hahaha");
        }

        private IntPtr GetObjectDescriptor(string name, string src)
        {
            var s1 = Encoding.Unicode.GetBytes(name);
            var s2 = Encoding.Unicode.GetBytes(src);
            var l1 = Marshal.SizeOf(typeof(OBJECTDESCRIPTOR));
            var l2 = l1 + s1.Length + 2;
            var len = l2 + s2.Length + 2;
            var ret = Marshal.AllocHGlobal(len);
            OBJECTDESCRIPTOR d = new OBJECTDESCRIPTOR
            {
                cbSize = len,
                clsid = GetGuid(),
                dwDrawAspect = 1,
                sizel_x = 300, //TODO
                sizel_y = 200,
                dwFullUserTypeName = l1,
                dwSrcOfCopy = l2,
            };
            Marshal.StructureToPtr(d, ret, false);
            Marshal.Copy(s1, 0, IntPtr.Add(ret, l1), s1.Length);
            Marshal.WriteInt16(ret, l2 - 2, 0);
            Marshal.Copy(s2, 0, IntPtr.Add(ret, l2), s2.Length);
            Marshal.WriteInt16(ret, len - 2, 0);
            return ret;
        }

        private IStorage GetSourceStorage()
        {
            IStorage stg;
            Natives.StgCreateDocfile(null, 0x4011012, 0, out stg);
            Guid guid = GetGuid();
            Natives.WriteClassStg(stg, ref guid);
            return stg;
        }

        private IntPtr GetSourceStoragePtr()
        {
            IStorage stg = GetSourceStorage();
            return Marshal.GetIUnknownForObject(stg);
        }

        private IntPtr GetImage()
        {
            if (OleObject == null)
            {
                return IntPtr.Zero;
            }

            IntPtr hDC = Natives.CreateDC("DISPLAY", null, null, IntPtr.Zero);
            Metafile mf = new Metafile(hDC, EmfType.EmfOnly);
            using (var g = Graphics.FromImage(mf))
            {
                OleObject.Draw(g);
            }
            IntPtr _hEmf = mf.GetHenhmetafile();
            uint _bufferSize = Natives.GdipEmfToWmfBits(_hEmf, 0, null, Natives.MM_ANISOTROPIC,
                Natives.EmfToWmfBitsFlags.EmfToWmfBitsFlagsDefault);
            byte[] _buffer = new byte[_bufferSize];
            Natives.GdipEmfToWmfBits(_hEmf, _bufferSize, _buffer, Natives.MM_ANISOTROPIC,
                    Natives.EmfToWmfBitsFlags.EmfToWmfBitsFlagsDefault);
            IntPtr hwmf = Natives.SetMetaFileBitsEx(_bufferSize, _buffer);

            var pic = Marshal.AllocHGlobal(20);
            Marshal.WriteInt32(pic, 0, 8/*MM_ANISOTROPIC*/);
            Marshal.WriteInt32(pic, 4, 20000);
            Marshal.WriteInt32(pic, 8, 20000);
            Marshal.WriteIntPtr(pic, 12, hwmf);
            return pic;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct OBJECTDESCRIPTOR
        {
            public int cbSize;
            public Guid clsid;
            public int dwDrawAspect;
            public int sizel_x;
            public int sizel_y;
            public int pointl_x;
            public int pointl_y;
            public int dwStatus;
            public int dwFullUserTypeName;
            public int dwSrcOfCopy;
        }
    }
}
