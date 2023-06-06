using ComRefactor.ComReflection.TypeLibs.Abstract;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract;
using System.Collections.Generic;

namespace ComRefactor.ComManagement.TypeLibs.Abstract
{
    public interface ITypeInfoInternalWrapperCollection
    {
        int Count { get; }

        ITypeInfoInternalWrapper GetTypeInfo(int index);
        ITypeInfoInternalWrapper Find(string searchTypeName);
        ITypeInfoInternalWrapper Get(string searchTypeName);
        IEnumerator<ITypeInfoInternalWrapper> GetEnumerator();

    }
}
