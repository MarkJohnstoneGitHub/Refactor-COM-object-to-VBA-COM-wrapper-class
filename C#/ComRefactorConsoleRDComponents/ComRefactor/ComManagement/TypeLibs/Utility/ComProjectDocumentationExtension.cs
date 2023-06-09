using ComRefactor.ComManagement.TypeLibs.Abstract;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Utility;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ComRefactorConsole.ComRefactor.ComManagement.TypeLibs.Utility
{
    internal static class ComProjectDocumentationExtension
    {
        public static void Document(this ComProject _this, StringLineBuilder output)
        {
            output.AppendLine();
            output.AppendLine("================================================================================");
            output.AppendLine();

            output.AppendLine("ITypeLib: " + _this.Name);

              
            output.AppendLineNoNullChars("- Documentation: " + _this.Documentation.DocString);
            output.AppendLineNoNullChars("- HelpContext: " + _this.Documentation.HelpContext);
            output.AppendLineNoNullChars("- HelpFile: " + _this.Documentation.HelpFile);

            output.AppendLine("- Guid: " + _this.Guid);
            
            //output.AppendLine("- Lcid: " + _this.Attributes.lcid);
            //output.AppendLine("- SysKind: " + _this.Attributes.syskind);
            //output.AppendLine("- LibFlags: " + _this.Attributes.wLibFlags);
            output.AppendLine("- MajorVer: " + _this.MajorVersion);
            output.AppendLine("- MinorVer: " + _this.MinorVersion);


            //output.AppendLine("- TypeCount: " + _this.TypesCount);

        }

    }
}
