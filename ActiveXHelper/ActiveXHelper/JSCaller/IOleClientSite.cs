using System;
using System.Runtime.InteropServices;

namespace lpp.ActiveXHelper.JSCaller
{
    /// <summary>
    /// 用于ActiveX调用js函数
    /// </summary>
    [ComImport,
    Guid("00000118-0000-0000-C000-000000000046"),
    InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IOleClientSite
    {
        void SaveObject();
        void GetMoniker(uint dwAssign, uint dwWhichMoniker, object ppmk);
        void GetContainer(out IOleContainer ppContainer);
        void ShowObject();
        void OnShowWindow(bool fShow);
        void RequestNewObjectLayout();
    }
}
