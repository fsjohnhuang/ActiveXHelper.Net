using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Text;
using System.Reflection;
using lpp.ActiveXHelper.JSCaller;
using mshtml;

namespace lpp.ActiveXHelper
{
    public class BaseActiveX : UserControl, IObjectSafety
    {
        #region IObjectSafety 成员

        private const string _IID_IDispatch = "{00020400-0000-0000-C000-000000000046}";
        private const string _IID_IDispatchEx = "{a6ef9860-c720-11d0-9337-00a0c90dcaa9}";
        private const string _IID_IPersistStorage = "{0000010A-0000-0000-C000-000000000046}";
        private const string _IID_IPersistStream = "{00000109-0000-0000-C000-000000000046}";
        private const string _IID_IPersistPropertyBag = "{37D84F60-42CB-11CE-8135-00AA004BB851}";

        private const int INTERFACESAFE_FOR_UNTRUSTED_CALLER = 0x00000001;
        private const int INTERFACESAFE_FOR_UNTRUSTED_DATA = 0x00000002;
        private const int S_OK = 0;
        private const int E_FAIL = unchecked((int)0x80004005);
        private const int E_NOINTERFACE = unchecked((int)0x80004002);

        private bool _fSafeForScripting = true;
        private bool _fSafeForInitializing = true;


        public int GetInterfaceSafetyOptions(ref Guid riid, ref int pdwSupportedOptions, ref int pdwEnabledOptions)
        {
            int Rslt = E_FAIL;

            string strGUID = riid.ToString("B");
            pdwSupportedOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER | INTERFACESAFE_FOR_UNTRUSTED_DATA;
            switch (strGUID)
            {
                case _IID_IDispatch:
                case _IID_IDispatchEx:
                    Rslt = S_OK;
                    pdwEnabledOptions = 0;
                    if (_fSafeForScripting == true)
                        pdwEnabledOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER;
                    break;
                case _IID_IPersistStorage:
                case _IID_IPersistStream:
                case _IID_IPersistPropertyBag:
                    Rslt = S_OK;
                    pdwEnabledOptions = 0;
                    if (_fSafeForInitializing == true)
                        pdwEnabledOptions = INTERFACESAFE_FOR_UNTRUSTED_DATA;
                    break;
                default:
                    Rslt = E_NOINTERFACE;
                    break;
            }

            return Rslt;
        }

        public int SetInterfaceSafetyOptions(ref Guid riid, int dwOptionSetMask, int dwEnabledOptions)
        {
            int Rslt = E_FAIL;

            string strGUID = riid.ToString("B");
            switch (strGUID)
            {
                case _IID_IDispatch:
                case _IID_IDispatchEx:
                    if (((dwEnabledOptions & dwOptionSetMask) == INTERFACESAFE_FOR_UNTRUSTED_CALLER) &&
                            (_fSafeForScripting == true))
                        Rslt = S_OK;
                    break;
                case _IID_IPersistStorage:
                case _IID_IPersistStream:
                case _IID_IPersistPropertyBag:
                    if (((dwEnabledOptions & dwOptionSetMask) == INTERFACESAFE_FOR_UNTRUSTED_DATA) &&
                            (_fSafeForInitializing == true))
                        Rslt = S_OK;
                    break;
                default:
                    Rslt = E_NOINTERFACE;
                    break;
            }

            return Rslt;
        }

        #endregion

        #region 调用js函数
        private Type typeIOleObject = null;
        private IOleClientSite oleClientSite = null;
        private IOleContainer pObj = null;

        /// <summary>
        /// 调用JS函数
        /// </summary>
        /// <param name="fnName">js函数名</param>
        /// <param name="args">入参</param>
        protected void CallJS(string fnName, params object[] args)
        {
            if (typeIOleObject == null)
            {
                typeIOleObject = this.GetType().GetInterface("IOleObject", true);
                object tmpOldClientSite = typeIOleObject.InvokeMember("GetClientSite",
                 BindingFlags.Instance | BindingFlags.InvokeMethod | BindingFlags.Public,
                null,
                this,
                null);

                oleClientSite = tmpOldClientSite as IOleClientSite;
                oleClientSite.GetContainer(out pObj);
            }

            //获取页面的Script集合
            IHTMLDocument pDoc2 = (IHTMLDocument)pObj;
            object script = pDoc2.Script;

            try
            {
                //调用JavaScript方法OnScaned并传递参数，因为此方法可能并没有在页面中实现，所以要进行异常处理
                script.GetType().InvokeMember(fnName,
                BindingFlags.Instance | BindingFlags.InvokeMethod | BindingFlags.Public,
               null,
                script,
                args);
            }
            catch { }
        }
        #endregion
    }
}
