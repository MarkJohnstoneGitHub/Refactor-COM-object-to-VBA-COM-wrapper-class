using ComRefactor.Refactoring.CodeBuilder.VBA;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Text;

namespace ComRefactor.Refactoring.CodeBuilder.VBA
{

    // TODO : Issue with method names using VBA reserved words
    // TODO : Issue with Com method names in Lowercase quickfix done
    // TODO : Issue for interface being used when referring to the current Com Object being implemented
    //  To replace interface with implementing object for the Com object being wrapped check against its default interface
    //  For other objects would require to find implementing object in the Com Libary and/or external libraries?
    //  require a list of ComClass Names, default interface and if interface is implemented by multiple objects?
    //  If interface has multiple implementations??? Use interface or if implementation not found.
    //  Eg. IFormatProvider used  in DotNetLib.ToString3(ByRef provider As IFormatProvider) As String
    //  IFormatProvider to stay as is as has multiple implementations contained in mscorlib
    //  eg  CultureInfo : ICloneable, IFormatProvider implements IFormatProvider
    // https://learn.microsoft.com/en-us/dotnet/api/system.iformatprovider.getformat?view=netframework-4.8.1
    // Maybe check if not in list of default implementation use interface reference?
    // Easiest way to fully implement wrapping a Com Object would require access to any external type libraries referenced?
    // Usually would be in the same type libary?
    // Maybe could be VBA
    // Or for DotNetLib \Windows\Microsoft.NET\Framework64\v4.0.30319\mscorlib.tlb
    // ComBase has parent property. 
    // See ComCoClass might be enumerable lists that might help or ComProject ??
    //    TypeLibTypeFlags
    //    Indicates that the interface derives from IDispatch, either directly or indirectly.
    //    FDispatchable = 0x1000,

    // If TKIND_INTERFACE then inface and get implementation object name and or qualifier name?

    // https://stackoverflow.com/questions/45063037/find-dependent-type-libraries-in-typelib-file-through-code
    // ITypeInfo::GetContainingTypeLib
    // https://stackoverflow.com/a/45090861 


    // TODO : Check for parameters and return type equal to ComCoClass default interface
    // eg. 
    // TODO : Issue missing Function As clause

    // https://www.engram9.info/access-2007-vba/reserved-word-list.html

    public class CodeModuleMethod
    {
        private const string Quote = "\"";

        String _memberName;
        String _comObjectVariable;

        public ComMember MethodInfo;

        public String ModuleName; 

        List<CodeModuleParameter> _parametersCode = new List<CodeModuleParameter>(); // TODO :

        public string ParametersDeclaration => "(" + String.Join(", ", _parametersCode) + ")";

        public IEnumerable<CodeModuleParameter> ParametersCode => this._parametersCode;

        private IEnumerable<ComParameter> _parameters => this.MethodInfo.Parameters;


        public CodeModuleMethod(ComMember methodInfo, String moduleName, String comObjectVariable)
        {
            this.MethodInfo = methodInfo;
            this._memberName = FirstLetterToUpper(this.MethodInfo.Name);
            this.ModuleName = moduleName;
            this._comObjectVariable = comObjectVariable;
            CodeParamaters();
        }

        public string Name 
        {
            get => _memberName;
            private set => _memberName = value;
        } 

        //TODO: Missing function and property get As clause
        public String Signature()
        {
            return Declaration() + ParametersDeclaration + " " + ReturnType();
        }
        public String CodeModule() 
        {
            StringBuilder method = new StringBuilder();

            // Annotations
            // TODO : Add enumeration annoation
            if (this.MethodInfo.IsDefault)
            {
                method.AppendLine(AnnotationDefaultMember());
            }

            if (this.MethodInfo.Documentation.DocString != null)
            {
                method.AppendLine(AnnotationMemberDescription());
            }

            method.AppendLine(Signature());

            // Member attributes
            if (this.MethodInfo.Documentation.DocString != null)
            {
                method.AppendLine(AttributeDescription()); 
            }
            if (this.MethodInfo.IsDefault)
            {
                method.AppendLine(AttributeDefaultMember());
            }

            //TODO method attribute for enumeration
            //TODO: reference to Com object  being wrapped
            if (this._comObjectVariable != null) 
            {
                // TODO : AppendLine reference to COM object being wrapped, also indent
                //method.AppendLine();
            }

            method.AppendLine(DeclarationEnd());

            return method.ToString();
        }


        public String Declaration()
        {
            var type = string.Empty;
            switch (this.MethodInfo.Type)
            {
                case DeclarationType.Function:
                    type = "Function";
                    break;
                case DeclarationType.Procedure:
                    type = "Sub";
                    break;
                case DeclarationType.PropertyGet:
                    type = "Property Get";
                    break;
                case DeclarationType.PropertyLet:
                    type = "Property Let";
                    break;
                case DeclarationType.PropertySet:
                    type = "Property Set";
                    break;
                case DeclarationType.Event:
                    type = "Event";
                    break;
            }
            return $"{(this.MethodInfo.IsHidden || this.MethodInfo.IsRestricted ? "Private" : "Public")} {type} {this.Name}";
        }

        public String DeclarationEnd()
        {
            var type = string.Empty;
            switch (this.MethodInfo.Type)
            {
                case DeclarationType.Function:
                    type = "End Function";
                    break;
                case DeclarationType.Procedure:
                    type = "End Sub";
                    break;
                case DeclarationType.PropertyGet:
                    type = "End Property";
                    break;
                case DeclarationType.PropertyLet:
                    type = "End Property";
                    break;
                case DeclarationType.PropertySet:
                    type = "End Property";
                    break;
                    //case DeclarationType.Event:  // TODO : research VBA events
                    //    type = "End Sub";
                    //    break;
            }
            return type;
        }


        //'@DefaultMember
        //Attribute Item.VB_UserMemId = 0
        public String AttributeDefaultMember()
        {
            if (this.MethodInfo.IsDefault)
            {
                return $"Attribute {Name}.VB_UserMemId = 0";
            }
            return String.Empty;
        }

        public String AttributeDescription()
        {
            if (this.MethodInfo.Documentation.DocString != null)
            {
                return $"Attribute {Name}.VB_Description = {Quote}{this.MethodInfo.Documentation.DocString}{Quote}";
            }
            return String.Empty;
        }


        private void CodeParamaters()
        {
            foreach (var parameter in this.MethodInfo.Parameters)
            {
                _parametersCode.Add(new CodeModuleParameter(parameter,this)); 
            }
        }


        // TODO : Issue with return types being interface, how to handle for other interfaces returned that are in the current library?
        // TODO : Posssible pass in the Com library object to check and replace with qualified name eg. DotNetLib.TimeSpan? or keep as ITimeSpan?
        public String ReturnType()
        {
            if (this.MethodInfo.Type == DeclarationType.Function || this.MethodInfo.Type == DeclarationType.PropertyGet)
            {
                string returnType = this.MethodInfo.AsTypeName.TypeName;

                //If the return type is the interface name replace with moduleName/new class name  ???
                if (returnType != null) 
                { 
                    if (returnType == this.MethodInfo.Parent.Name)
                    {
                        returnType = this.ModuleName;  //If new class name selected use instead of default moduleName
                    }
                }
                else 
                {
                    // TODO : throw error???
                }

                return "As " + returnType;
            }
            return null;
        }


        //Eg. For a function
        // Public Function CreateFromTicks(ByVal ticks As LongLong, Optional ByVal kind As DateTimeKind = DateTimeKind_Unspecified) As IDateTime
        //     With New DateTime
        //          Set .CreateFromTicks = mDateTime.CreateFromTicks(ticks, kind)
        //     End With
        // End Function
        // Eg. returns  Set CreateFromTicks = mDateTime.CreateFromTicks(ticks, kind)
        // Require "Set" if return type is an object
        // Require the list of parameter names
        // return declaration type,

        // Dim pvtDateTime as DateTime
        // pvtDateTime = New DateTime
        // Set CreateFromTicks = mDateTime.CreateFromTicks(ticks, kind)

        //eg output this.DotNetLibDateTime.CreateFromTicks(ticks, kind)
        private String ComObjectWrapperReference()
        {
            List<String> parameterNames = new List<String>();

            foreach (var parameter in this._parameters)
            {
                parameterNames.Add(parameter.Name);
            }
            String joined = "(" + String.Join(", ", parameterNames) + ")";
            
            return this._comObjectVariable + this.MethodInfo.Name + joined;
        }

        // TODO: eg returns Set CreateFromTicks = mDateTime.CreateFromTicks(ticks, kind)
        private String ComObjectVariableDeclaration() 
        {

            if (this.MethodInfo.Type == DeclarationType.Function || this.MethodInfo.Type == DeclarationType.PropertyGet)
            {
                string assignment = this.MethodInfo.Name + " = ";
                if (this.MethodInfo.AsTypeName.IsByRef)
                {
                    assignment = "Set " + assignment;

                    // Require return object eg Date how to object implementation from interface?
                    //TODO : create code block to create object to return
                    // EG.
                    //     With New DateTime
                    //          Set .CreateFromTicks = mDateTime.CreateFromTicks(ticks, kind)
                    //     End With
                    return assignment;
                }
                else
                {
                    return assignment;
                }
                // TODO require to determine if object for Set
               
            }
            return string.Empty;
        }


        //https://github.com/rubberduck-vba/Rubberduck/wiki/VB_Attribute-Annotations#member-annotations
        public string AnnotationMemberDescription()
        {
            return "'@Description(\"" + this.MethodInfo.Documentation.DocString + "\")";
        }

        public string AnnotationDefaultMember()
        {
            if (this.MethodInfo.IsDefault)
            {
                return "'@DefaultMember";
            }
            return String.Empty;

        }


        // https://stackoverflow.com/questions/4135317/make-first-letter-of-a-string-upper-case-with-maximum-performance/4135491#4135491
        // TODO : Move to appropriate location
        private string FirstLetterToUpper(string str)
        {
            if (str == null)
                return null;

            if (str.Length > 1)
                return char.ToUpper(str[0]) + str.Substring(1);

            return str.ToUpper();
        }



    }

}



// TODO : Issue with optional parameters if equal null 
// TODO : Move to CodeModuleParrameters
//public String Parameters()
//{
//    List<String> declarationParameters = new List<String>();

//    foreach (var parameter in this.MethodInfo.Parameters)
//    {
//        String declarationName;
//        if (parameter.TypeName == this.MethodInfo.Parent.Name)
//        {
//            string parameterTypeName = this.ModuleName;

//            declarationName = parameter.DeclarationName; // TODO: update to class name i.e. default is the  comCoClass.Name or new module name ???

//            //See _comParameter
//            //declarationName = $"{(parameter.IsOptional ? "Optional " : string.Empty)}{(parameter.IsByRef ? "ByRef" : "ByVal")} {parameter.Name} As {parameterTypeName}{(parameter.IsOptional && parameter.DefaultValue != null ? " = " : string.Empty)}{(parameter.IsOptional && parameter.DefaultValue != null ? _typeName.IsEnumMember ? DefaultAsEnum : DefaultValue : string.Empty)}";
//        }
//        else
//        {
//            declarationName = parameter.DeclarationName;
//        }

//        declarationParameters.Add(declarationName);
//    }
//    String joined = "(" + String.Join(", ", declarationParameters) + ")";
//    return joined;
//}
