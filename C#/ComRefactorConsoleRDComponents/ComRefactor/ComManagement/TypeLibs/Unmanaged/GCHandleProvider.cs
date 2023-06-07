using System;
using System.Runtime.InteropServices;

namespace ComRefactor.ComManagement.TypeLibs.Unmanaged
{

    // https://stackoverflow.com/questions/17339928/c-sharp-how-to-convert-object-to-intptr-and-back/52103996#52103996 
    public class GCHandleProvider : IDisposable
    {
        public GCHandleProvider(object target)
        {
            Handle = target.ToGcHandle();
        }

        public IntPtr Pointer => Handle.ToIntPtr();

        public GCHandle Handle { get; }

        private void ReleaseUnmanagedResources()
        {
            if (Handle.IsAllocated) Handle.Free();
        }

        public void Dispose()
        {
            ReleaseUnmanagedResources();
            GC.SuppressFinalize(this);
        }

        ~GCHandleProvider()
        {
            ReleaseUnmanagedResources();
        }
    }
}
