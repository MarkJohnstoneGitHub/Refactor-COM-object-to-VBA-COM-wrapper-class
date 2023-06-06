using System;
using System.Diagnostics;
using System.Runtime.InteropServices.ComTypes;
using ComRefactorConsole.ComRefactorr.ComManagement.TypeLibs;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract;
//using Rubberduck.VBEditor.ComManagement.TypeLibs.DebugInternal;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Unmanaged;

// TODO The tracers are broken - using them will cause a NRE inside the 
// unmanaged boundary. If we need to enable them for diagnostics, this needs
// to be fixed first. 

namespace Rubberduck.VBEditor.ComManagement.TypeLibs
{
    /// <summary>
    /// Abstracts out the creation of the custom implementations of
    /// <see cref="ITypeLib"/> and <see cref="ITypeInfo"/>, mainly to
    /// make it easier to compose the implementation. For example, tracing
    /// can be enabled via the factory with appropriate compilation flags. 
    /// </summary>
    internal static class InternalTypeApiFactory
    {
        internal static ITypeLibInternalWrapper GetTypeLibWrapper(IntPtr rawObjectPtr, bool addRef)
        {
            ITypeLibInternalWrapper wrapper = new TypeLibInternalWrapper(rawObjectPtr, addRef);
            //TraceWrapper(ref wrapper);
            return wrapper;
        }

        //[Conditional("TRACE_TYPEAPI")]
        //private static void TraceWrapper(ref ITypeLibInternalWrapper wrapper)
        //{
        //    wrapper = new TypeLibWrapperTracer(wrapper, (ITypeLibInternal)wrapper);
        //}

        //internal static ITypeLibInternalWrapper GetTypeInfoWrapper(IntPtr rawObjectPtr, int? parentUserFormUniqueId = null)
        //{
        //    ITypeLibInternalWrapper wrapper = new TypeLibInternalWrapper(rawObjectPtr, parentUserFormUniqueId);
        //    //TraceWrapper(ref wrapper);
        //    return wrapper;
        //}


        //internal static ITypeInfoInternalWrapper GetTypeInfoWrapper(ITypeInfo rawTypeInfo)
        //{
        //    TypeInfoInternalWrapper wrapper = new TypeInfoInternalWrapper(rawTypeInfo);
        //    //TraceWrapper(ref wrapper);
        //    return wrapper;
        //}

        internal static ITypeInfoInternalWrapper GetTypeInfoWrapper(ITypeInfo rawTypeInfo)
        {
            ITypeInfoInternalWrapper wrapper = new TypeInfoInternalWrapper(rawTypeInfo);
            //TraceWrapper(ref wrapper);
            return wrapper;
        }

        internal static ITypeInfoInternalWrapper GetTypeInfoWrapper(IntPtr value)
        {
            ITypeInfoInternalWrapper wrapper = new TypeInfoInternalWrapper(value);
            //TraceWrapper(ref wrapper);
            return wrapper;
        }

        //[Conditional("TRACE_TYPEAPI")]
        //private static void TraceWrapper(ref ITypeInfoWrapper wrapper)
        //{
        //    wrapper = new TypeInfoWrapperTracer(wrapper, (ITypeInfoInternal)wrapper);
        //}
    }
}
