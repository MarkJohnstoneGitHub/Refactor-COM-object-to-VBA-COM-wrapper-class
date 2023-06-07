using ComRefactor.ComReflection.TypeLibs.Abstract;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract;
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


//public interface ITypeInfoWrapperCollection
//{
//    int Count { get; }
//    ITypeInfoWrapper GetItemByIndex(int index);
//    ITypeInfoWrapper Find(string searchTypeName);
//    ITypeInfoWrapper Get(string searchTypeName);
//    IEnumerator<ITypeInfoWrapper> GetEnumerator();
//}