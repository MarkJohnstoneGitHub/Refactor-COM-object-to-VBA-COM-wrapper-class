using System;
using System.IO;
using System.Runtime.InteropServices.ComTypes;
using ComRefactor.ComReflection;
using ComRefactor.Refactoring.CodeBuilder.VBA;
using Rubberduck.Parsing.ComReflection;

namespace ComRefactorConsole
{
    // https://github.com/MarkJohnstoneGitHub/Refactor-COM-object-to-VBA-COM-wrapper-class
    // Issues: 
    // TODO : Issue with methods using VBA reserved words eg method Date

    internal class Program
    {
        static void Main(string[] args)
        {
            string typeLibraryPath = args[0];  // path of type library
            string comClassName = args[1];     // COM object name required to wrap
            string outputPath = args[2];       // output path for VBA COM wrapper 
            string moduleName = args[3];       // VBA class name

            Boolean isPredeclaredId = false;
            //  https://stackoverflow.com/questions/49590754/convert-a-string-to-a-boolean-in-c-sharp
            Boolean.TryParse(args[4], out isPredeclaredId);  // Select predeclared VBA COM wrapper

            if (File.Exists(typeLibraryPath))
            {
                ComLibraryProvider libraryInfo = new ComLibraryProvider();
                ITypeLib  typeLib = libraryInfo.LoadTypeLibrary(typeLibraryPath);
                if (typeLib != null)
                {
                    ComProject projectTypeLib = new ComProject(typeLib, typeLibraryPath);
                    VBAComWrapper codebuilder = new VBAComWrapper(projectTypeLib, comClassName, moduleName, isPredeclaredId);

                    //TODO : If  comClassName not found??
                    String codeModule = codebuilder.CodeModule();
                    System.IO.File.WriteAllText(outputPath, codeModule);

                    // https://stackoverflow.com/questions/3826763/get-full-path-without-filename-from-path-that-includes-filename
                    // https://stackoverflow.com/questions/674479/how-do-i-get-the-directory-from-a-files-full-path

                }
                else
                {
                }
            }
        }

    }
}
