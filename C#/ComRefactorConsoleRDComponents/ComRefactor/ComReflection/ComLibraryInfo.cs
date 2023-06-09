﻿//using Rubberduck.Parsing.ComReflection;
using Rubberduck.Parsing.ComReflection;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices.ComTypes;

namespace ComRefactor.ComReflection
{
    public class ComLibraryInfo
    {
        public static readonly List<string> TypeLibraryExtensions = new List<string> { ".tlb" , ".olb", ".tlb", ".dll", ".ocx", ".exe" };
        private readonly IComLibraryProvider _libraryProvider = new ComLibraryProvider();
        private ITypeLib _typelib;
        

        public ComProjectLibrary GetLibraryInfoFromPath(string path)
        {
            try
            {
                var extension = Path.GetExtension(path)?.ToLowerInvariant() ?? string.Empty;
                if (string.IsNullOrEmpty(extension))
                {
                    return null;
                }

                // LoadTypeLibrary will attempt to open files in the host, so only attempt on possible COM servers.
                if (TypeLibraryExtensions.Contains(extension))
                {
                    this._typelib = _libraryProvider.LoadTypeLibrary(path);
                    return new ComProjectLibrary(this._typelib, path); 
                }
                return null; 
            }
            catch
            {
                // Most likely this is unloadable. If not, it we can't fail here because it could have come from the Apply
                // button in the AddRemoveReferencesDialog. Wait for it...  :-P
                return null; 
            }
        }

    }
}
