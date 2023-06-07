using System;
using System.Runtime.InteropServices;

namespace ComRefactor.ComManagement.TypeLibs.Unmanaged
{
    // https://stackoverflow.com/questions/17339928/c-sharp-how-to-convert-object-to-intptr-and-back/52103996#52103996 
    public static class ObjectHandleExtensions
    {
        public static IntPtr ToIntPtr(this object target)
        {
            return GCHandle.Alloc(target).ToIntPtr();
        }

        public static GCHandle ToGcHandle(this object target)
        {
            return GCHandle.Alloc(target);
        }

        public static IntPtr ToIntPtr(this GCHandle target)
        {
            return GCHandle.ToIntPtr(target);
        }
    }
}
