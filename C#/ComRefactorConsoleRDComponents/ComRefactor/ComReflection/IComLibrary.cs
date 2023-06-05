using System.Runtime.InteropServices.ComTypes;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.ComReflection
{
    public interface IComLibrary
    {
        ITypeLib LoadTypeLibrary(string libraryPath);
        IComDocumentation GetComDocumentation(ITypeLib typelib);

    }
}
