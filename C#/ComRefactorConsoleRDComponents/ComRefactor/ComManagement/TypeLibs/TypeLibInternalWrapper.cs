using Rubberduck.VBEditor.ComManagement.TypeLibs.Unmanaged;
using ComTypes = System.Runtime.InteropServices.ComTypes;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Utility;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract;
using System.Collections.ObjectModel;

namespace ComRefactorConsole.ComRefactorr.ComManagement.TypeLibs
{
    internal class TypeLibInternalWrapper :  TypeLibInternalSelfMarshalForwarderBase
    {
      
       private readonly ReadOnlyCollection<ITypeInfoInternal> _cachedTypeInfos;

        private ComPointer<ITypeLibInternal> _typeLibPointer;

        private ITypeLib _typeLib;

        private ITypeLibInternal _target_ITypeLib => _typeLibPointer.Interface;

        ////TODO : Collection of  ITypeInfo
        //public ITypeInfoWrapperCollection TypeInfos { get; private set; }

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
            //TypeInfos = new TypeInfoWrapperCollection(this);
            //// ReSharper disable once SuspiciousTypeConversion.Global 
            //// there is no direct implementation but it can be reached via
            //// IUnknown::QueryInterface which is implicitly done as part of casting
            //HasVBEExtensions = _target_ITypeLib is IVBEProject;
        }

        private void InitFromRawPointer(IntPtr rawObjectPtr, bool addRef)
        {
            if (!UnmanagedMemoryHelper.ValidateComObject(rawObjectPtr))
            {
                throw new ArgumentException("Expected COM object, but validation failed.");
            }

            _typeLibPointer = ComPointer<ITypeLibInternal>.GetObject(rawObjectPtr, addRef);
            InitCommon();
        }

        /// <summary>
        ///// Constructor -- should be called via <see cref="TypeApiFactory"/> only.
        /// </summary>
        /// <param name="rawObjectPtr">The raw unmanaged ITypeLib pointer</param>
        /// <param name="addRef">
        /// Indicates that the pointer was obtained via unorthodox methods, such as
        /// direct memory read. Setting the parameter will effect an IUnknown::AddRef
        /// on the pointer. 
        /// </param>
        internal TypeLibInternalWrapper(IntPtr rawObjectPtr, bool addRef)
        {
            InitFromRawPointer(rawObjectPtr, addRef);
        }


        private bool _isDisposed;
        public override void Dispose()
        {
            if (_isDisposed) return;
            _isDisposed = true;


            // TODO : _cachedTypeInfos?.Dispose();
            //_cachedTypeInfos?.Dispose();
            _typeLibPointer.Dispose();
        }

        public int GetSafeTypeInfoByIndex(int index, out ITypeInfoInternal outTI)
        {
            outTI = null;

            using (var typeInfoPtr = AddressableVariables.Create<IntPtr>())
            {
                var hr = _target_ITypeLib.GetTypeInfo(index, typeInfoPtr.Address);
                if (ComHelper.HRESULT_FAILED(hr))
                {
                    return HandleBadHRESULT(hr);
                }

                // TODO : 
                //var outVal = TypeApiFactory.GetTypeInfoWrapper(typeInfoPtr.Value);
                //_cachedTypeInfos = _cachedTypeInfos ?? new DisposableList<ITypeInfoWrapper>();
                //_cachedTypeInfos.Add(outVal);
                //outTI = outVal;

                return hr;
            }
        }


        //public void GetTypeInfo(int index, out ITypeInfo typeInfo)
        //{
        //   _typeLib.GetTypeInfo(index, out typeInfo);
        //}

        public IntPtr GetCOMReferencePtr()
            => RdMarshal.GetComInterfaceForObject(this, typeof(ITypeLibInternal));

        int HandleBadHRESULT(int hr)
        {
            return hr;
        }

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


            // TODO : RdMarshal.WriteIntPtr(ppTI, ti.GetCOMReferencePtr());
            //RdMarshal.WriteIntPtr(ppTI, ti.GetCOMReferencePtr());
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

            // TODO : 
            //using (var outVal = TypeApiFactory.GetTypeInfoWrapper(pTInfo)) // takes ownership of the COM reference [pTInfo]
            //{
            //    RdMarshal.WriteIntPtr(ppTInfo, outVal.GetCOMReferencePtr());

            //    _cachedTypeInfos = _cachedTypeInfos ?? new DisposableList<ITypeInfoWrapper>();
            //    _cachedTypeInfos.Add(outVal);
            //}

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


    }
}
