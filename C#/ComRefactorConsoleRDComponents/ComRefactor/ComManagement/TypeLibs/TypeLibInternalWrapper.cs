using Rubberduck.VBEditor.ComManagement.TypeLibs.Unmanaged;
using ComTypes = System.Runtime.InteropServices.ComTypes;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Utility;
using System;
using ComRefactor.ComManagement.TypeLibs.Abstract;
using ComRefactor.ComReflection.TypeLibs.Abstract;
using ComRefactor.ComManagement.TypeLibs;
using System.Runtime.InteropServices.ComTypes;
using ComRefactor.ComManagement.TypeLibs.Unmanaged;

namespace ComRefactorr.ComManagement.TypeLibs
{
    internal class TypeLibInternalWrapper :  TypeLibInternalSelfMarshalForwarderBase, ITypeLibInternalWrapper
    {
        private DisposableList<ITypeInfoInternalWrapper> _cachedTypeInfos;

        private ITypeLib _ITypeLib;
        private ITypeLibInternal _target_ITypeLib => (ITypeLibInternal)_ITypeLib;

        //private ComPointer<ITypeLibInternal> _typeLibPointer;
        //private ITypeLibInternal _target_ITypeLib => _typeLibPointer.Interface;

        private bool _isDisposed;

        public ITypeInfoInternalWrapperCollection TypeInfos { get; set; }

        // helpers
        public string Name => CachedTextFields._name;
        public string DocString => CachedTextFields._docString;
        public int HelpContext => CachedTextFields._helpContext;
        public string HelpFile => CachedTextFields._helpFile;
        public int TypesCount => _target_ITypeLib.GetTypeInfoCount();

        private struct TypeLibTextFields
        {
            public string _name;
            public string _docString;
            public int _helpContext;
            public string _helpFile;
        }

        private TypeLibTextFields? _cachedTextFields;
        private TypeLibTextFields CachedTextFields
        {
            get
            {
                if (_cachedTextFields.HasValue)
                {
                    return _cachedTextFields.Value;
                }

                var cache = new TypeLibTextFields();
                // as a C# caller, it's easier to work with ComTypes.ITypeLib
               ((ComTypes.ITypeLib)_target_ITypeLib).GetDocumentation((int)KnownDispatchMemberIDs.MEMBERID_NIL, out cache._name, out cache._docString, out cache._helpContext, out cache._helpFile);
                _cachedTextFields = cache;
                return _cachedTextFields.Value;
            }
        }

        private void InitCommon()
        {
            //TODO : InitCommon
            TypeInfos = new TypeInfoInternalWrapperCollection(this);
            //// ReSharper disable once SuspiciousTypeConversion.Global 
            //// there is no direct implementation but it can be reached via
            //// IUnknown::QueryInterface which is implicitly done as part of casting
            //HasVBEExtensions = _target_ITypeLib is IVBEProject;
        }

        public TypeLibInternalWrapper(ITypeLib typeLib)
        {
            this._ITypeLib = typeLib;
            InitCommon();
        }

        public int GetSafeTypeInfoByIndex(int index, out ITypeInfoInternalWrapper outTI)
        {
            outTI = null;

            using (var typeInfoPtr = AddressableVariables.Create<IntPtr>())
            {
                var hr = _target_ITypeLib.GetTypeInfo(index, typeInfoPtr.Address);
                if (ComHelper.HRESULT_FAILED(hr))
                {
                    return HandleBadHRESULT(hr);
                }

                var outVal = InternalTypeApiFactory.GetTypeInfoInternalWrapper(typeInfoPtr.Value);
                _cachedTypeInfos = _cachedTypeInfos ?? new DisposableList<ITypeInfoInternalWrapper>();
                _cachedTypeInfos.Add(outVal);
                outTI = outVal;

                return hr;
            }
        }

        int ITypeLibInternalWrapper.GetSafeTypeInfoByIndex(int index, out ITypeInfoInternalWrapper outTI)
        {
            var result = GetSafeTypeInfoByIndex(index, out var outTIW);
            outTI = outTIW;
            return result;
        }

        private ComTypes.TYPELIBATTR? _cachedLibAttribs;
        public ComTypes.TYPELIBATTR Attributes
        {
            get
            {
                if (_cachedLibAttribs.HasValue)
                {
                    return _cachedLibAttribs.Value;
                }

                using (var typeLibAttributesPtr = AddressableVariables.CreatePtrTo<ComTypes.TYPELIBATTR>())
                {
                    var hr = _target_ITypeLib.GetLibAttr(typeLibAttributesPtr.Address);
                    if (ComHelper.HRESULT_FAILED(hr))
                    {
                        return _cachedLibAttribs.Value;
                    }

                    _cachedLibAttribs = typeLibAttributesPtr.Value.Value;   // dereference the ptr, then the content
                    var pTypeLibAttr = typeLibAttributesPtr.Value.Address; // dereference the ptr, and take the contents address
                    _target_ITypeLib.ReleaseTLibAttr(pTypeLibAttr);         // can release immediately as _cachedLibAttribs is a copy
                }
                return _cachedLibAttribs.Value;
            }
        }

        public IntPtr GetCOMReferencePtr()
            => RdMarshal.GetComInterfaceForObject(this, typeof(ITypeLibInternal));

        int HandleBadHRESULT(int hr)
        {
            return hr;
        }

        // ITypeLibInternal virtuals to be implemented by the derived class.

        public override int GetTypeInfoCount()
        {
            var retVal = _target_ITypeLib.GetTypeInfoCount();
            return retVal;
        }

        public override int GetTypeInfo(int index, IntPtr ppTI)
        {
            // We have to wrap the ITypeInfo returned by GetTypeInfo
            var hr = GetSafeTypeInfoByIndex(index, out var ti);
            if (ComHelper.HRESULT_FAILED(hr)) return HandleBadHRESULT(hr);

            RdMarshal.WriteIntPtr(ppTI, ti.GetCOMReferencePtr());
            return hr;
        }

        public override int GetTypeInfoType(int index, IntPtr pTKind)
        {
            var hr = _target_ITypeLib.GetTypeInfoType(index, pTKind);
            if (ComHelper.HRESULT_FAILED(hr)) return HandleBadHRESULT(hr);

            var tKind = RdMarshal.ReadInt32(pTKind);
            RdMarshal.WriteInt32(pTKind, tKind);

            return hr;
        }
        public override int GetTypeInfoOfGuid(ref Guid guid, IntPtr ppTInfo)
        {
            var hr = _target_ITypeLib.GetTypeInfoOfGuid(guid, ppTInfo);
            if (ComHelper.HRESULT_FAILED(hr)) return HandleBadHRESULT(hr);

            var pTInfo = RdMarshal.ReadIntPtr(ppTInfo);

            using (var outVal = InternalTypeApiFactory.GetTypeInfoInternalWrapper(pTInfo)) // takes ownership of the COM reference [pTInfo]
            {
                RdMarshal.WriteIntPtr(ppTInfo, outVal.GetCOMReferencePtr());

                _cachedTypeInfos = _cachedTypeInfos ?? new DisposableList<ITypeInfoInternalWrapper>();
                _cachedTypeInfos.Add(outVal);
            }

            return hr;
        }

        public override int GetLibAttr(IntPtr ppTLibAttr)
        {
            var hr = _target_ITypeLib.GetLibAttr(ppTLibAttr);
            return ComHelper.HRESULT_FAILED(hr) ? HandleBadHRESULT(hr) : hr;
        }
        public override int GetTypeComp(IntPtr ppTComp)
        {
            var hr = _target_ITypeLib.GetTypeComp(ppTComp);
            return ComHelper.HRESULT_FAILED(hr) ? HandleBadHRESULT(hr) : hr;
        }
        public override int GetDocumentation(int index, IntPtr strName, IntPtr strDocString, IntPtr dwHelpContext, IntPtr strHelpFile)
        {
            var hr = _target_ITypeLib.GetDocumentation(index, strName, strDocString, dwHelpContext, strHelpFile);
            return ComHelper.HRESULT_FAILED(hr) ? HandleBadHRESULT(hr) : hr;
        }
        public override int IsName(string szNameBuf, int lHashVal, IntPtr pfName)
        {
            var hr = _target_ITypeLib.IsName(szNameBuf, lHashVal, pfName);
            return ComHelper.HRESULT_FAILED(hr) ? HandleBadHRESULT(hr) : hr;
        }
        public override int FindName(string szNameBuf, int lHashVal, IntPtr ppTInfo, IntPtr rgMemId, IntPtr pcFound)
        {
            var hr = _target_ITypeLib.FindName(szNameBuf, lHashVal, ppTInfo, rgMemId, pcFound);
            return ComHelper.HRESULT_FAILED(hr) ? HandleBadHRESULT(hr) : hr;
        }
        public override void ReleaseTLibAttr(IntPtr pTLibAttr)
        {
            _target_ITypeLib.ReleaseTLibAttr(pTLibAttr);
        }

        public override void Dispose()
        {
            if (_isDisposed) return;
            _isDisposed = true;
            _cachedTypeInfos?.Dispose();
           // _typeLibPointer.Dispose();
        }

    }
}
