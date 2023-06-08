using ComRefactor.ComReflection.TypeLibs.Abstract;
using System;
using System.Runtime.InteropServices.ComTypes;

namespace ComRefactor.ComManagement.TypeLibs.Abstract
{
    public interface ITypeLibInternalWrapper : ITypeLib, IDisposable
    {
        string Name { get; }
        string DocString { get; }
        int HelpContext { get; }
        string HelpFile { get; }
        //bool HasVBEExtensions { get; }
        int TypesCount { get; }

        //TODO : 
        ITypeInfoInternalWrapperCollection TypeInfos { get; }

        System.Runtime.InteropServices.ComTypes.TYPELIBATTR Attributes { get; }

        int GetSafeTypeInfoByIndex(int index, out ITypeInfoInternalWrapper outTI);
    }
}