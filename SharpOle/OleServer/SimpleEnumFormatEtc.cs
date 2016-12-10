using SharpOle.OleInterop;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace SharpOle.OleServer
{
    [ComVisible(true)]
    [Guid("0BE18992-E4D9-4C09-BFB4-367B36B67694")]
    [ProgId("SharpOleLib.SimpleEnumFormatEtc")]
    [ClassInterface(ClassInterfaceType.None)]
    public class SimpleEnumFormatEtc : IEnumFormatEtc
    {
        private uint _Index = 0;
        private const int _Count = 3;

        public SimpleEnumFormatEtc()
        {
        }

        public int Clone(out IEnumFormatEtc ppEnum)
        {
            ppEnum = new SimpleEnumFormatEtc();
            return 0;
        }

        public int Next(uint celt, FormatEtc[] rgelt, uint[] pceltFetched)
        {
            for (uint i = 0; i < celt; ++i)
            {
                if (_Index >= _Count)
                {
                    if (pceltFetched != null)
                    {
                        pceltFetched[0] = i;
                    }
                    return 1;
                }
                Fill(rgelt, i);

                _Index += 1;
            }
            if (pceltFetched != null)
            {
                pceltFetched[0] = celt;
            }
            return 0;
        }

        public int Reset()
        {
            _Index = 0;
            return 0;
        }

        public int Skip(uint celt)
        {
            _Index += celt;
            return _Index >= _Count ? 1 : 0;
        }

        private void Fill(FormatEtc[] rgelt, uint i)
        {
            if (_Index == 0)
            {
                rgelt[i].cfFormat = 0xC00E; //object descriptor
                rgelt[i].dwAspect = 1;
                rgelt[i].lindex = -1;
                rgelt[i].ptd = IntPtr.Zero;
                rgelt[i].tymed = 1; //hglobal
            }
            else if (_Index == 1)
            {
                rgelt[i].cfFormat = 0xC00B; //embed source
                rgelt[i].dwAspect = 1;
                rgelt[i].lindex = -1;
                rgelt[i].ptd = IntPtr.Zero;
                rgelt[i].tymed = 8; //istorage
            }
            else if (_Index == 2)
            {
                rgelt[i].cfFormat = 3; //CF_METAFILEPICT
                rgelt[i].dwAspect = 1;
                rgelt[i].lindex = -1;
                rgelt[i].ptd = IntPtr.Zero;
                rgelt[i].tymed = 32; //
            }
        }
    }
}
