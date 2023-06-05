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

namespace ComRefactorConsole.ComRefactorr.ComManagement.TypeLibs
{
    internal class TypeLibInternalWrapper
    {
        //TODO : ITypeInternal??
        //private DisposableList<ITypeInfoWrapper> _cachedTypeInfos;

        private ComPointer<ITypeLibInternal> _typeLibPointer;

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
    }
}
