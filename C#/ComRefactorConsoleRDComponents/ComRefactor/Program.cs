using System;
using System.IO;
using ComRefactor.ComReflection;
using ComRefactor.Refactoring.CodeBuilder.VBA;
using Rubberduck.Parsing.ComReflection;

namespace ComRefactorConsole
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string typeLibraryPath = args[0];
            string outputPath = args[1];
            string comClassName = args[2];

            Boolean isPredeclaredId = false;
            //  https://stackoverflow.com/questions/49590754/convert-a-string-to-a-boolean-in-c-sharp
            Boolean.TryParse(args[3], out isPredeclaredId);

            if (File.Exists(typeLibraryPath))
            {
                ComLibraryInfo libraryInfo = new ComLibraryInfo();
                ComProjectLibrary projectTypeLib = libraryInfo.GetLibraryInfoFromPath(typeLibraryPath);
                ComCoClass comCoClass = projectTypeLib.FindComCoClass(comClassName);
                if (comCoClass != null)
                {
                    string codeModule = null;
                    VBAComWrapper codebuilder = new VBAComWrapper(comCoClass, comClassName,isPredeclaredId);
                    codeModule = codebuilder.CodeModule();
                    System.IO.File.WriteAllText(outputPath, codeModule);

                    // Appears to be working finding the Com CoClass default interface required eg. "DateTime" from DotNetLib.tlb
                    // Successfully created class for all methods and descriptions

                    // TODO : Issue with methods using VBA reserved words eg method Date

                    // https://stackoverflow.com/questions/3826763/get-full-path-without-filename-from-path-that-includes-filename
                    // https://stackoverflow.com/questions/674479/how-do-i-get-the-directory-from-a-files-full-path
                    // TODO : For testing obtain the file path from outputPath to document the ComInterface 
                    // TODO : Document ComInterface comCoClassInterface
                    // TODO : Refactor comCoClassInterface to extract VBA Com wrapper class
                    // TODO : Require to investigate how RD extracts class and/or interface and include RD components required
                    // TODO : First attempt just extract a VBA class without Com wrapper references
                    // TODO : Modify RD refactoring a VBA interface to extract a class wrapping  the Com object required


                }
                else 
                { 
                }
            }
        }

    }
}
