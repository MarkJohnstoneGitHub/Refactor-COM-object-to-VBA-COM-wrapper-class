using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Rubberduck.Parsing.ComReflection;

namespace ComRefactor.ComReflection
{

    // https://github.com/rubberduck-vba/Rubberduck/blob/next/Rubberduck.Parsing/ComReflection/ComLibraryProvider.cs
    // https://learn.microsoft.com/en-us/windows/win32/api/oaidl/nn-oaidl-itypelib


    public class ComLibrary //: IComLibrary
    {

        #region Native Stuff
        /// <summary>
        /// Controls how a type library is registered.
        /// </summary>
        private enum REGKIND
        {
            /// <summary>
            /// Use default register behavior.
            /// </summary>
            REGKIND_DEFAULT = 0,
            /// <summary>
            /// Register this type library.
            /// </summary>
            REGKIND_REGISTER = 1,
            /// <summary>
            /// Do not register this type library.
            /// </summary>
            REGKIND_NONE = 2
        }

        [DllImport("oleaut32.dll", CharSet = CharSet.Unicode)]
        private static extern int LoadTypeLibEx(string strTypeLibName, REGKIND regKind, out ITypeLib TypeLib);
        #endregion

        
        public static ITypeLib LoadTypeLibrary(string libraryPath)
        {
            LoadTypeLibEx(libraryPath, REGKIND.REGKIND_NONE, out var typeLibrary);
            return typeLibrary;
        }

        public static IComDocumentation GetComDocumentation(ITypeLib typelib)
        {
            try
            {
                return new ComDocumentation(typelib, ComDocumentation.LibraryIndex);
            }
            catch
            {
                return null;
            }
        }
    }
}
