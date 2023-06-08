using ComRefactor.ComReflection.TypeLibs.Abstract;
using System.Collections.Generic;

namespace ComRefactor.ComManagement.TypeLibs.Abstract
{
    public interface ITypeInfoInternalWrapperCollection
    {
        int Count { get; }

        ITypeInfoInternalWrapper GetItemByIndex(int index);
        ITypeInfoInternalWrapper Find(string searchTypeName);
        ITypeInfoInternalWrapper Get(string searchTypeName);
        IEnumerator<ITypeInfoInternalWrapper> GetEnumerator();

    }
}