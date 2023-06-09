using ComRefactor.ComReflection;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Utility;

namespace ComRefactor.ComManagement.TypeLibs.Utility
{
    internal static class ComProjectDocumentationExtension
    {
        public static void Document(this ComProjectLibrary _this, StringLineBuilder output)
        {
            output.AppendLine();
            output.AppendLine("================================================================================");
            output.AppendLine();

            output.AppendLine("ITypeLib: " + _this.Name);

              
            output.AppendLineNoNullChars("- Documentation: " + _this.Documentation.DocString);
            output.AppendLineNoNullChars("- HelpContext: " + _this.Documentation.HelpContext);
            output.AppendLineNoNullChars("- HelpFile: " + _this.Documentation.HelpFile);

            output.AppendLine("- Guid: " + _this.Attributes.guid);
            output.AppendLine("- Lcid: " + _this.Attributes.lcid);
            output.AppendLine("- SysKind: " + _this.Attributes.syskind);
            output.AppendLine("- wLibFlags: " + _this.Attributes.wLibFlags);
            output.AppendLine("- MajorVer: " + _this.Attributes.wMajorVerNum);
            output.AppendLine("- MinorVer: " + _this.Attributes.wMinorVerNum);


            output.AppendLine("- TypeCount: " + _this.TypeCount);

        }

    }
}
