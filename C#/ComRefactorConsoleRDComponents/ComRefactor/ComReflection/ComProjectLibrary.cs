﻿using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.Serialization;
using Rubberduck.VBEditor.ComManagement.TypeLibs;
using Rubberduck.VBEditor.Utility;
using TYPEATTR = System.Runtime.InteropServices.ComTypes.TYPEATTR;
using TYPELIBATTR = System.Runtime.InteropServices.ComTypes.TYPELIBATTR;
using TYPEKIND = System.Runtime.InteropServices.ComTypes.TYPEKIND;
using SYSKIND = System.Runtime.InteropServices.ComTypes.SYSKIND;

using Rubberduck.Parsing.ComReflection;
using System.Xml.Linq;

namespace ComRefactor.ComReflection
{
    // Modified Rubberduck.Parsing.ComReflection.ComProject

    [DataContract]
    [KnownType(typeof(ComBase))]
    [DebuggerDisplay("{" + nameof(Name) + "}")]
    public class ComProjectLibrary : ComBase
    {
        public static readonly ConcurrentDictionary<Guid, ComType> KnownTypes = new ConcurrentDictionary<Guid, ComType>();
        public static readonly ConcurrentDictionary<Guid, ComEnumeration> KnownEnumerations = new ConcurrentDictionary<Guid, ComEnumeration>();
        public static readonly ConcurrentDictionary<Guid, ComAlias> KnownAliases = new ConcurrentDictionary<Guid, ComAlias>();

        private ITypeLib _typeLib;  //TODO : Added
        public int TypeCount;       //TODO : Added

        public TYPELIBATTR Attributes; //TODO : Added

        [DataMember(IsRequired = true)]
        public string Path { get; set; }

        [DataMember(IsRequired = true)]
        public long MajorVersion { get; set; }

        [DataMember(IsRequired = true)]
        public long MinorVersion { get; set; }

        // YGNI...
        // ReSharper disable once NotAccessedField.Local
#pragma warning disable IDE0052 // Remove unread private members
        public readonly TypeLibTypeFlags wLibFlags;
#pragma warning restore IDE0052 // Remove unread private members

        [DataMember(IsRequired = true)]
        private readonly List<ComAlias> _aliases = new List<ComAlias>();
        public IEnumerable<ComAlias> Aliases => _aliases;

        [DataMember(IsRequired = true)]
        private readonly List<ComInterface> _interfaces = new List<ComInterface>();
        public IEnumerable<ComInterface> Interfaces => _interfaces;

        [DataMember(IsRequired = true)]
        private readonly List<ComEnumeration> _enumerations = new List<ComEnumeration>();
        public IEnumerable<ComEnumeration> Enumerations => _enumerations;

        [DataMember(IsRequired = true)]
        private readonly List<ComCoClass> _classes = new List<ComCoClass>();
        public IEnumerable<ComCoClass> CoClasses => _classes;

        [DataMember(IsRequired = true)]
        private readonly List<ComModule> _modules = new List<ComModule>();
        public IEnumerable<ComModule> Modules => _modules;

        [DataMember(IsRequired = true)]
        private readonly List<ComStruct> _structs = new List<ComStruct>();
        public IEnumerable<ComStruct> Structs => _structs;

        //Note - Enums and Types should enumerate *last*. That will prevent a duplicate module in the unlikely(?)
        //instance where the TypeLib defines a module named "Enums" or "Types".
        public IEnumerable<IComType> Members => _modules.Cast<IComType>()
            .Union(_interfaces)
            .Union(_classes)
            .Union(_enumerations)
            .Union(_structs);

        public ComProjectLibrary(ITypeLib typeLibrary, string path) : base(null, typeLibrary, -1)
        {
            Path = path;

            try
            {
                typeLibrary.GetLibAttr(out IntPtr attribPtr);
                using (DisposalActionContainer.Create(attribPtr, typeLibrary.ReleaseTLibAttr))
                {
                    var typeAttr = Marshal.PtrToStructure<TYPELIBATTR>(attribPtr);

                    Attributes = typeAttr; // TODO: Added

                    MajorVersion = typeAttr.wMajorVerNum;  // TODO : Using Attrributes
                    MinorVersion = typeAttr.wMinorVerNum;  // TODO : Using Attrributes
                    wLibFlags = (TypeLibTypeFlags)typeAttr.wLibFlags; // TODO : Using Attrributes
                    Guid = typeAttr.guid;  // TODO : Using Attrributes
                }
            }
            catch (COMException) { }
            this._typeLib = typeLibrary;
            LoadModules(typeLibrary);
        }

        // TODO : Added returrn the interface for the comCoClass required
        // https://stackoverflow.com/questions/4937060/how-to-check-if-listt-element-contains-an-item-with-a-particular-property-valu/4937099#4937099
        // https://stackoverflow.com/questions/15456845/getting-a-list-item-by-index/15456851#15456851

        // returrn the ComInterface object found
        public ComInterface FindComCoClassInterface(string comCoClassName)
        {

            int index = this._classes.FindIndex(item => item.Name == comCoClassName);
            if (index >= 0)
            {
                ComCoClass comCoClass = this._classes[index];
                return comCoClass.DefaultInterface;
            }
            return null;

        }

        public ComCoClass FindComCoClass(string comCoClassName)
        {
            int index = this._classes.FindIndex(item => item.Name == comCoClassName);
            if (index >= 0)
            {
                ComCoClass comCoClass = this._classes[index];
                return comCoClass;
            }
            return null;
        }


        private void LoadModules(ITypeLib typeLibrary)
        {
            var typeCount = typeLibrary.GetTypeInfoCount();
            TypeCount = typeCount; // TODO : Added
            for (var index = 0; index < typeCount; index++)
            {
                try
                {
                    typeLibrary.GetTypeInfo(index, out ITypeInfo info);
                    info.GetTypeAttr(out var typeAttributesPointer);
                    using (DisposalActionContainer.Create(typeAttributesPointer, info.ReleaseTypeAttr))
                    {
                        var typeAttributes = Marshal.PtrToStructure<TYPEATTR>(typeAttributesPointer);
                        KnownTypes.TryGetValue(typeAttributes.guid, out var type);
                        
                        switch (typeAttributes.typekind)
                        {
                            case TYPEKIND.TKIND_ENUM:
                                var enumeration = type ?? new ComEnumeration(this, typeLibrary, info, typeAttributes, index);
                                Debug.Assert(enumeration is ComEnumeration);
                                _enumerations.Add(enumeration as ComEnumeration);
                                if (type == null && !enumeration.Guid.Equals(Guid.Empty))
                                {
                                    KnownTypes.TryAdd(typeAttributes.guid, enumeration);
                                }
                                break;
                            case TYPEKIND.TKIND_COCLASS:
                                var coclass = type ?? new ComCoClass(this, typeLibrary, info, typeAttributes, index);
                                Debug.Assert(coclass is ComCoClass && !coclass.Guid.Equals(Guid.Empty));
                                _classes.Add(coclass as ComCoClass);
                                if (type == null)
                                {
                                    KnownTypes.TryAdd(typeAttributes.guid, coclass);
                                }
                                break;
                            case TYPEKIND.TKIND_DISPATCH:
                            case TYPEKIND.TKIND_INTERFACE:
                                var intface = type ?? new ComInterface(this, typeLibrary, info, typeAttributes, index);
                                Debug.Assert(intface is ComInterface && !intface.Guid.Equals(Guid.Empty));
                                _interfaces.Add(intface as ComInterface);
                                if (type == null)
                                {
                                    KnownTypes.TryAdd(typeAttributes.guid, intface);
                                }
                                break;
                            case TYPEKIND.TKIND_RECORD:
                                var structure = new ComStruct(this, typeLibrary, info, typeAttributes, index);
                                _structs.Add(structure);
                                break;
                            case TYPEKIND.TKIND_MODULE:
                                var module = type ?? new ComModule(this, typeLibrary, info, typeAttributes, index);
                                Debug.Assert(module is ComModule);
                                _modules.Add(module as ComModule);
                                if (type == null && !module.Guid.Equals(Guid.Empty))
                                {
                                    KnownTypes.TryAdd(typeAttributes.guid, module);
                                }
                                break;
                            case TYPEKIND.TKIND_ALIAS:
                                var alias = new ComAlias(this, typeLibrary, info, index, typeAttributes);
                                _aliases.Add(alias);
                                if (alias.Guid != Guid.Empty)
                                {
                                    KnownAliases.TryAdd(alias.Guid, alias);
                                }
                                break;
                            case TYPEKIND.TKIND_UNION:
                                //TKIND_UNION is not a supported member type in VBA.
                                break;
                            default:
                                throw new NotImplementedException($"Didn't expect a TYPEATTR with multiple typekind flags set in {Path}.");
                        }
                    }
                }
                catch (COMException) { }
            }
            ApplySpecificLibraryTweaks();
        }

        private void ApplySpecificLibraryTweaks()
        {
            if (!Name.ToUpper().Equals("EXCEL")) return;
            var application = _classes.SingleOrDefault(x => x.Guid.ToString().Equals("00024500-0000-0000-c000-000000000046"));
            var worksheetFunction = _interfaces.SingleOrDefault(i => i.Guid.ToString().Equals("00020845-0000-0000-c000-000000000046"));
            if (application != null && worksheetFunction != null)
            {
                application.AddInterface(worksheetFunction);
            }
        }
    }
}
