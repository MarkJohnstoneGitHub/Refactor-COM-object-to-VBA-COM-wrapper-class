using Rubberduck.Parsing.ComReflection;
using System;
using System.Text;

namespace ComRefactor.Refactoring.CodeBuilder.VBA
{

    // TODO : Research VBA Events
    // https://stackoverflow.com/questions/41023670/can-we-use-interfaces-and-events-together-at-the-same-time
    // http://www.cpearson.com/excel/Events.aspx
    // https://stackoverflow.com/questions/39511528/exposing-net-events-to-com

    public class VBAComWrapper
    {
        static int _capacity = 255;
        static int _maxCapacity = 65536;

        private StringBuilder _codeBuilder;
        private ComInterface _template;
        private String _moduleName;
        private bool _isPredeclaredId;

        public VBAComWrapper(ComInterface template, String moduleName, bool isPredeclaredId = false ) 
        {
            _codeBuilder = new StringBuilder(_capacity, _maxCapacity);
            _template = template;
            _moduleName = moduleName;
            _isPredeclaredId = isPredeclaredId;
            BuildCodeModule();
        }

        public String CodeModule()
        {
            return _codeBuilder.ToString();
        }

        private void BuildCodeModule()
        {
            this._codeBuilder.AppendLine(CodeModuleHeader.Header);
            CodeModuleHeaderAttributes headerAttributes = new CodeModuleHeaderAttributes(this._moduleName, _template.Documentation.DocString, this._isPredeclaredId);
            this._codeBuilder.AppendLine(headerAttributes.CodeModule());
            if (this._isPredeclaredId)
            {
                this._codeBuilder.AppendLine(AnnotationPredeclaredId());
            }

            if (this._template.Documentation.DocString != null)
            {
                this._codeBuilder.AppendLine(AnnotationModuleDescription()); 
            }
            this._codeBuilder.AppendLine();
            this._codeBuilder.AppendLine(CodeModuleOptionExplicit.OptionExplicit);
            this._codeBuilder.AppendLine();

            //TODO : require to handle if VBA class module exceeds maximum size 65536
            foreach (var methodInfo in this._template.Members)
            {
                if (!methodInfo.IsRestricted)
                {
                    CodeModuleMethod method = new CodeModuleMethod(methodInfo);
                    _codeBuilder.AppendLine(method.CodeModule());
                }
            }
        }

        // https://github.com/rubberduck-vba/Rubberduck/wiki/VB_Attribute-Annotations#module-annotations
        private string AnnotationPredeclaredId()
        {
            return "'@PredeclaredId";
        }

        // https://github.com/rubberduck-vba/Rubberduck/wiki/VB_Attribute-Annotations#module-annotations
        private string AnnotationModuleDescription()
        {
            return "'@ModuleDescription(\"" + _template.Documentation.DocString + "\")";
        }


    }
}
