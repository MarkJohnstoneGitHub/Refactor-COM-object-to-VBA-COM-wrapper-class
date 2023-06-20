using Rubberduck.Parsing.ComReflection;
using Rubberduck.Parsing.Symbols;
using System;
using System.Text;

namespace ComRefactor.Refactoring.CodeBuilder.VBA
{

    // TODO : Research VBA Events
    // https://stackoverflow.com/questions/41023670/can-we-use-interfaces-and-events-together-at-the-same-time
    // http://www.cpearson.com/excel/Events.aspx
    // https://stackoverflow.com/questions/39511528/exposing-net-events-to-com
    // Example of a custom collection https://nolongerset.com/strongly-typed-collection-classes-the-easy-way/
    // https://codereview.stackexchange.com/questions/60504/building-a-better-collection-enumerable-in-vba

    public class VBAComWrapper
    {
        static int _capacity = 255;
        static int _maxCapacity = 65536;

        private StringBuilder _codeBuilder;
        private ComCoClass _comCoClass;
        private ComInterface _comInterface => this._comCoClass.DefaultInterface;

        private String _moduleName;
        private bool _isPredeclaredId;

        public VBAComWrapper(ComCoClass comCoClass, String moduleName, bool isPredeclaredId = false)
        {
            _codeBuilder = new StringBuilder(_capacity, _maxCapacity);
            _comCoClass = comCoClass;
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
            CodeModuleHeaderAttributes headerAttributes = new CodeModuleHeaderAttributes(this._moduleName, _comInterface.Documentation.DocString, this._isPredeclaredId);
            this._codeBuilder.AppendLine(headerAttributes.CodeModule());
            if (this._isPredeclaredId)
            {
                this._codeBuilder.AppendLine(AnnotationPredeclaredId());
            }

            if (this._comInterface.Documentation.DocString != null)
            {
                this._codeBuilder.AppendLine(AnnotationModuleDescription()); 
            }
            this._codeBuilder.AppendLine();
            this._codeBuilder.AppendLine(CodeModuleOptionExplicit.OptionExplicit);
            this._codeBuilder.AppendLine();

            //private variable to wrapping Com object
            this._codeBuilder.AppendLine(VariableComObject());
            this._codeBuilder.AppendLine();

            //TODO : require to handle if VBA class module exceeds maximum size 65536
            foreach (var methodInfo in this._comInterface.Members)
            {
                if (!methodInfo.IsRestricted)
                {
                    CodeModuleMethod method = new CodeModuleMethod(methodInfo, this._moduleName);
                    _codeBuilder.AppendLine(method.CodeModule());
                }
            }
        }

        //require name of COM object being wrapped
        //output eg. Private mDateTime As DotNetLib.DateTime
        private String VariableComObject()
        {
            string output;
            output = "Private " + "m" + _comCoClass.Name + " As " + _comCoClass.Parent.Name + "." + _comCoClass.Name;   
            return output; 
        }    

        // https://github.com/rubberduck-vba/Rubberduck/wiki/VB_Attribute-Annotations#module-annotations
        private string AnnotationPredeclaredId()
        {
            return "'@PredeclaredId";
        }

        // https://github.com/rubberduck-vba/Rubberduck/wiki/VB_Attribute-Annotations#module-annotations
        private string AnnotationModuleDescription()
        {
            return "'@ModuleDescription(\"" + _comInterface.Documentation.DocString + "\")";
        }


    }
}
