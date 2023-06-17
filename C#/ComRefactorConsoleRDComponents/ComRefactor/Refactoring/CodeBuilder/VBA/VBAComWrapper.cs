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

        public VBAComWrapper(ComInterface template, String moduleName) 
        {
            _codeBuilder = new StringBuilder(_capacity, _maxCapacity);
            _template = template;
            _moduleName = moduleName;
            BuildCodeModule();
        }

        public String CodeModule()
        {
            return _codeBuilder.ToString();
        }

        private void BuildCodeModule()
        {
            this._codeBuilder.AppendLine(CodeModuleHeader.Header);
            CodeModuleHeaderAttributes headerAttributes = new CodeModuleHeaderAttributes(this._moduleName, _template.Documentation.DocString, true);
            this._codeBuilder.AppendLine(headerAttributes.CodeModule());
            this._codeBuilder.Append(CodeModuleOptionExplicit.OptionExplicit);
            //TODO : RD annotations for predeclardId and module description
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


    }
}
