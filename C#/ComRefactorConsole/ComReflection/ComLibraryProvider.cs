using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Rubberduck.Parsing.ComReflection;

namespace ComRefactorConsole.ComReflection
{

    // https://github.com/rubberduck-vba/Rubberduck/blob/next/Rubberduck.Parsing/ComReflection/ComLibraryProvider.cs
    // https://learn.microsoft.com/en-us/windows/win32/api/oaidl/nn-oaidl-itypelib


    public class ComLibraryProvider
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

        //public ReferenceInfo GetReferenceInfo(ITypeLib typelib, string name, string path)
        //{
        //    try
        //    {
        //        typelib.GetLibAttr(out var attributes);
        //        using (DisposalActionContainer.Create(attributes, typelib.ReleaseTLibAttr))
        //        {
        //            var typeAttr = Marshal.PtrToStructure<System.Runtime.InteropServices.ComTypes.TYPELIBATTR>(attributes);

        //            return new ReferenceInfo(typeAttr.guid, name, path, typeAttr.wMajorVerNum, typeAttr.wMinorVerNum);
        //        }
        //    }
        //    catch
        //    {
        //        return ReferenceInfo.Empty;
        //    }
        //}
    }
}
