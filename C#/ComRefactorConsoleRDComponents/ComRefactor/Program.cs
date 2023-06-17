using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using ComRefactor.ComManagement.TypeLibs.Utility;
using ComRefactor.ComReflection;
using ComRefactor.Refactoring.CodeBuilder.VBA;
using Rubberduck.Parsing.ComReflection;
using RD = Rubberduck.VBEditor.ComManagement.TypeLibs.Utility;

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
                ComInterface comCoClassInterface = projectTypeLib.FindComCoClassInterface(comClassName);
                if (comCoClassInterface != null)
                {
                    string codeModule = null;
                    VBAComWrapper codebuilder = new VBAComWrapper(comCoClassInterface, comClassName,isPredeclaredId);
                    codeModule = codebuilder.CodeModule();
                    System.IO.File.WriteAllText(outputPath, codeModule);

                    // Appears to be working finding the Com CoClass default interface required eg. "DateTime" from DotNetLib.tlb
                    // Successfully created class for all methods and descriptions

                    // TODO : Attributes for enumeration, default member etc.
                    // TODO : Issue with optional parameters
                    // TODO : Issue with methods using VBA reserved words eg method Date
                    // TODO : Implement wrapping of early binding Com object
                    // TODO : Rubberduck annotations for description etc.

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

        public static string DocumentComProject(ComProjectLibrary projectTypeLib)
        {
            var output = new RD.StringLineBuilder();
            projectTypeLib.Document(output);
            return output.ToString();
        }

        // https://stackoverflow.com/questions/43875454/is-there-a-way-to-view-com-entries-by-traversing-a-tlb-file-in-net
        public static void ParseTypeLib(string path,ITypeLib typeLib)
        {
            int count = typeLib.GetTypeInfoCount();
            IntPtr ipLibAtt = IntPtr.Zero;
            typeLib.GetLibAttr(out ipLibAtt);

            var typeLibAttr = (System.Runtime.InteropServices.ComTypes.TYPELIBATTR)
                Marshal.PtrToStructure(ipLibAtt, typeof(System.Runtime.InteropServices.ComTypes.TYPELIBATTR));
            Guid tlbId = typeLibAttr.guid;


            for (int i = 0; i < count; i++)
            {
                ITypeInfo typeInfo = null;
                typeLib.GetTypeInfo(i, out typeInfo);

                //figure out what guids, typekind, and names of the thing we're dealing with
                IntPtr ipTypeAttr = IntPtr.Zero;
                typeInfo.GetTypeAttr(out ipTypeAttr);

                //unmarshal the pointer into a structure into something we can read
                var typeattr = (System.Runtime.InteropServices.ComTypes.TYPEATTR)
                    Marshal.PtrToStructure(ipTypeAttr, typeof(System.Runtime.InteropServices.ComTypes.TYPEATTR));

                System.Runtime.InteropServices.ComTypes.TYPEKIND typeKind = typeattr.typekind;
                Guid typeId = typeattr.guid;

                //get the name of the type
                string strName, strDocString, strHelpFile;
                int dwHelpContext;
                typeLib.GetDocumentation(i, out strName, out strDocString, out dwHelpContext, out strHelpFile);


                if (typeKind == System.Runtime.InteropServices.ComTypes.TYPEKIND.TKIND_COCLASS)
                {
                    string xmlComClassFormat = "<comClass clsid=\"{0}\" tlbid=\"{1}\" description=\"{2}\" progid=\"{3}.{4}\"></comClass>";
                    string comClassXml = String.Format(xmlComClassFormat,
                        typeId.ToString("B").ToUpper(),
                        tlbId.ToString("B").ToUpper(),
                        strDocString,
                        path, strName
                        );
                    Debug.WriteLine(comClassXml + Environment.NewLine);
                    Console.WriteLine(comClassXml + Environment.NewLine);
                }
                else if (typeKind == System.Runtime.InteropServices.ComTypes.TYPEKIND.TKIND_INTERFACE)
                {
                    string xmlProxyStubFormat = "<comInterfaceExternalProxyStub name=\"{0}\" iid=\"{1}\" tlbid=\"{2}\" proxyStubClsid32=\"{3}\"></comInterfaceExternalProxyStub>";
                    string proxyStubXml = String.Format(xmlProxyStubFormat,
                        strName,
                        typeId.ToString("B").ToUpper(),
                        tlbId.ToString("B").ToUpper(),
                        "{00020424-0000-0000-C000-000000000046}"
                    );
                    Debug.WriteLine(proxyStubXml + Environment.NewLine);
                    Console.WriteLine(proxyStubXml + Environment.NewLine);
                }
            }
            return;
        }

    }
}
