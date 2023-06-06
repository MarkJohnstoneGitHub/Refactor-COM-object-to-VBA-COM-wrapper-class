using System.Runtime.InteropServices.ComTypes;
using Rubberduck.Parsing.ComReflection;

namespace ComRefactorConsole.ComReflection
{
    public interface IComLibrary
    {
        ITypeLib LoadTypeLibrary(string libraryPath);
        IComDocumentation GetComDocumentation(ITypeLib typelib);

    }
}
