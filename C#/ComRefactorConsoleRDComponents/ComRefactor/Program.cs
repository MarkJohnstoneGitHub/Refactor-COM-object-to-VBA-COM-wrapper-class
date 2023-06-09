﻿using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.Remoting.Contexts;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Documents;
using ComRefactor.ComManagement.TypeLibs.Abstract;
using ComRefactor.ComManagement.TypeLibs.Utility;
using ComRefactorConsole.ComRefactor;
using ComRefactorConsole.ComRefactor.ComManagement.TypeLibs.Utility;
using ComRefactorConsole.ComReflection;
using ComRefactorr.ComManagement.TypeLibs;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Utility;

namespace ComRefactorConsole
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string typeLibraryPath = args[0];
            string outputPath = args[1];

            if (File.Exists(typeLibraryPath))
            {

                ComLibraryInfo libraryInfo = new ComLibraryInfo();
                ComProject projectTypeLib = libraryInfo.GetLibraryInfoFromPath(typeLibraryPath);
                string output = DocumentComProject(projectTypeLib);
                System.IO.File.WriteAllText(outputPath, output);
            }
        }

        /// <summary>
        /// Documents the type library 
        /// </summary>
        /// <param name="projectTypeLib">Low-level ITypeLib wrapper</param>
        /// <returns>text document, in a non-standard format, useful for debugging purposes</returns>
        public static string DocumentTypeLibInternal(ITypeLibInternalWrapper projectTypeLib)
        {
            var output = new StringLineBuilder();
            projectTypeLib.Document(output);
            return output.ToString();
        }


        public static string DocumentComProject(ComProject projectTypeLib)
        {
            var output = new StringLineBuilder();
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
