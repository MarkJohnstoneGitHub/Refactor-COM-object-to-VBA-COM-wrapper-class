using Rubberduck.Parsing.ComReflection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ComRefactor.Refactoring.CodeBuilder.VBA
{
    public static class CodeModuleHeader
    {
        public static String Header => "VERSION 1.0 CLASS\r\nBEGIN\r\n  MultiUse = -1  'True\r\nEND";
    }
}
