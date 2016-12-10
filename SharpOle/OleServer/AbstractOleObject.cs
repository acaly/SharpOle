using SharpOle.OleInterop;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace SharpOle.OleServer
{
    public abstract class AbstractOleObject : AbstractDataObject, IOleObject, IPersistStorage
    {
        private IOleAdviseHolder _AdviseHolder;
        private IOleClientSite _ClientSite;
        private SizeL[] _Size;
        private string _HostAppName, _HostObjName;

        private readonly Guid _Guid;

        public AbstractOleObject()
        {
            //attach this to DataObject
            OleObject = this;

            _Size = new[] {
                new SizeL
                {
                    cx = 300,
                    cy = 200,
                }
            };

            _Guid = new Guid(this.GetType().GetCustomAttribute<GuidAttribute>().Value);
        }

        public abstract void Draw(Graphics g);

        public override Guid GetGuid()
        {
            return _Guid;
        }

        public void CopyToClipboard()
        {
            var data = new SimpleDataObject(this);
            data.SetClipboard();
        }

        void IOleObject.Advise(IAdviseSink pAdvSink, out uint pdwConnection)
        {
            if (_AdviseHolder == null)
            {
                Natives.CreateOleAdviseHolder(out _AdviseHolder);
            }
            _AdviseHolder.Advise(pAdvSink, out pdwConnection);
        }

        void IOleObject.Close(uint dwSaveOption)
        {
        }

        int IOleObject.DoVerb(int iVerb, Msg[] lpmsg, IOleClientSite pActiveSite, int lindex, IntPtr hWndParent, Rect[] lprcPosRect)
        {
            //TODO
            return 0;
        }

        void IOleObject.EnumAdvise(out IEnumStatData ppenumAdvise)
        {
            if (_AdviseHolder == null)
            {
                Natives.CreateOleAdviseHolder(out _AdviseHolder);
            }
            _AdviseHolder.EnumAdvise(out ppenumAdvise);
        }

        int IOleObject.EnumVerbs(out IEnumOleVerb ppEnumOleVerb)
        {
            //this exception is recongized by mapper
            throw new NotImplementedException();
        }

        void IOleObject.GetClientSite(out IOleClientSite ppClientSite)
        {
            ppClientSite = _ClientSite;
        }

        void IOleObject.GetClipboardData(uint dwReserved, out IDataObject ppDataObject)
        {
            throw new NotImplementedException();
        }

        void IOleObject.GetExtent(uint dwDrawAspect, SizeL[] pSizel)
        {
            pSizel[0] = _Size[0];
        }

        int IOleObject.GetMiscStatus(uint dwAspect, out uint pdwStatus)
        {
            pdwStatus = 0;
            return 0; //return value?
        }

        void IOleObject.GetMoniker(uint dwAssign, uint dwWhichMoniker, out IMoniker ppmk)
        {
            throw new NotImplementedException();
        }

        void IOleObject.GetUserClassID(out Guid pClsid)
        {
            pClsid = GetGuid();
        }

        int IOleObject.GetUserType(uint dwFormOfType, IntPtr pszUserType)
        {
            return 0x00040000;
        }

        int IOleObject.InitFromData(IDataObject pDataObject, int fCreation, uint dwReserved)
        {
            return 0;
        }

        int IOleObject.IsUpToDate()
        {
            return 0;
        }

        void IOleObject.SetClientSite(IOleClientSite pClientSite)
        {
            _ClientSite = pClientSite;
        }

        void IOleObject.SetColorScheme(LogPalette[] pLogpal)
        {
            //TODO we can get foreground color and background color here.
        }

        void IOleObject.SetExtent(uint dwDrawAspect, SizeL[] pSizel)
        {
            _Size[0] = pSizel[0];
        }

        void IOleObject.SetHostNames(string szContainerApp, string szContainerObj)
        {
            _HostAppName = szContainerApp;
            _HostObjName = szContainerObj;
        }

        void IOleObject.SetMoniker(uint dwWhichMoniker, IMoniker pmk)
        {
            throw new NotImplementedException();
        }

        void IOleObject.Unadvise(uint dwConnection)
        {
            _AdviseHolder.Unadvise(dwConnection);
        }

        int IOleObject.Update()
        {
            return 0;
        }

        #region Storage

        void IPersistStorage.GetClassID(out Guid pClassID)
        {
            pClassID = GetGuid();
        }

        void IPersistStorage.HandsOffStorage()
        {
        }

        void IPersistStorage.InitNew(IStorage pstg)
        {
        }

        int IPersistStorage.IsDirty()
        {
            return 1; //not dirty
        }

        void IPersistStorage.Load(IStorage pstg)
        {
        }

        void IPersistStorage.Save(IStorage pStgSave, int fSameAsLoad)
        {
        }

        void IPersistStorage.SaveCompleted(IStorage pStgNew)
        {
        }

        #endregion
    }
}
