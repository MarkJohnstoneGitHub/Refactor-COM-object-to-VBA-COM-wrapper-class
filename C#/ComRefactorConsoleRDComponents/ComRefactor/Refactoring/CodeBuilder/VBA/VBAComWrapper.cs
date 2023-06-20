using Rubberduck.Parsing.ComReflection;
using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
        const string Indent = "   ";

        static int _capacity = 255;
        static int _maxCapacity = 65536;

        private StringBuilder _codeBuilder;
        private ComCoClass _comCoClass;
        private ComInterface _comInterface => this._comCoClass.DefaultInterface;
        private bool _isPredeclaredId;
        public string QualifierName => _comCoClass.Parent.Name + "." + _comCoClass.Name;


        public String ModuleName;
        public string ComVariableName => "m" + _comCoClass.Name;

        public VBAComWrapper(ComCoClass comCoClass, String moduleName, bool isPredeclaredId = false)
        {
            _codeBuilder = new StringBuilder(_capacity, _maxCapacity);
            _comCoClass = comCoClass;
            ModuleName = moduleName;
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
            CodeModuleHeaderAttributes headerAttributes = new CodeModuleHeaderAttributes(this.ModuleName, _comInterface.Documentation.DocString, this._isPredeclaredId);
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

            //VBA Class_Initialize and Class_Terminate
            this._codeBuilder.AppendLine(ClassInitialize());
            this._codeBuilder.AppendLine(ClassTerminate());

            //Internal Properties
            this._codeBuilder.AppendLine(InternalPropertyGetComObject());
            this._codeBuilder.AppendLine(InternalPropertySetComObject());
            this._codeBuilder.AppendLine(InternalPropertySelf());

            //TODO : require to handle if VBA class module exceeds maximum size 65536
            foreach (var methodInfo in this._comInterface.Members)
            {
                if (!methodInfo.IsRestricted)
                {
                    //CodeModuleMethod method = new CodeModuleMethod(methodInfo, this.ModuleName);
                    CodeModuleMethod method = new CodeModuleMethod(methodInfo, this.ModuleName,this.ComVariableName );

                    // TODO : update to AppendLine method.Signiture
                    // TODO : loop through member attributes to AppendLine
                    // TODO : AppendLine reference to COM object being wrapped, also indent
                    // TODO : AppendLine member End

                    _codeBuilder.AppendLine(method.CodeModule());
                }
            }
        }

        //require name of COM object being wrapped
        //output eg. Private mDateTime As DotNetLib.DateTime
        private String VariableComObject()
        {
            string output;
            output = "Private " + this.ComVariableName + " As " + QualifierName;   
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


        // Eg.
        // Private Sub Class_Initialize()
        //    Set mCol = New VBA.Collection
        //End Sub
        private string ClassInitialize()
        {
            StringBuilder methodInitialize = new StringBuilder();
            methodInitialize.AppendLine("Private Sub Class_Initialize()");
            string codeLine = Indent + "Set " + ComVariableName + " = " + "New " + QualifierName;
            methodInitialize.AppendLine(codeLine);
            methodInitialize.AppendLine("End Sub");
            return methodInitialize.ToString();
        }

        private string ClassTerminate()
        {
            StringBuilder methodInitialize = new StringBuilder();
            methodInitialize.AppendLine("Private Sub Class_Terminate()");
            string codeLine = Indent + "Set " + ComVariableName + " = " + "Nothing";
            methodInitialize.AppendLine(codeLine);
            methodInitialize.AppendLine("End Sub");
            return methodInitialize.ToString();
        }

        //Internal propeties

        private string InternalPropertyGetComObject()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("Friend Property Get ComObject() As " + QualifierName);
            sb.AppendLine(Indent + "Set ComObject = " + ComVariableName);
            sb.AppendLine("End Property");
            return sb.ToString();
        }

        private string InternalPropertySetComObject()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("Friend Property Set ComObject(ByVal " + "obj"+_comCoClass.Name +   " As " + QualifierName +")");
            sb.AppendLine(Indent + "Set ComObject = " + ComVariableName);
            sb.AppendLine("End Property");
            return sb.ToString();
        }

        private string InternalPropertySelf()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("Friend Property Get Self() As " + ModuleName);
            sb.AppendLine(Indent + "Set Self = Me");
            sb.AppendLine("End Property");
            return sb.ToString();
        }
    }
}
