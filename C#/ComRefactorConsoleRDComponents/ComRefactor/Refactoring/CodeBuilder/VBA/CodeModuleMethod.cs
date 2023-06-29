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
        const string Indent = "   ";

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
            return Declaration() + ParametersDeclaration + (ReturnType() == null ? string.Empty : " As " + ReturnType());
        }
        public String CodeModule() 
        {
            StringBuilder method = new StringBuilder();

            if (this.MethodInfo.IsDefault)
            {
                method.AppendLine(AnnotationDefaultMember());
            }

            if (this.MethodInfo.IsEnumerator)
            {
                method.AppendLine(AnnotationEnumerator());
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

            if (this.MethodInfo.IsEnumerator)
            {
                method.AppendLine(AttributeEnumerator());
            }

            //TODO: reference to Com object  being wrapped
            if (this._comObjectVariable != null) 
            {
                method.Append(ComObjectVariableDeclaration());
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

        public String AttributeEnumerator()
        {
            if (this.MethodInfo.IsEnumerator)
            {
                return $"Attribute {Name}.VB_UserMemId = -4";
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


        // TODO : Issue with return types being interface, how to handle for other interfaces returned that are in the current type library or external type libraries?
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

                return returnType;  //TODO : Remove "As "
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
            String joinParameters = "(" + String.Join(", ", parameterNames) + ")";
            
            return $"{this._comObjectVariable}.{this.MethodInfo.Name}{joinParameters}"; //this._comObjectVariable + this.MethodInfo.Name + joinParameters;

        }

        // TODO: eg returns Set CreateFromTicks = mDateTime.CreateFromTicks(ticks, kind)

        // Issue :

        //Public Function CreateFromTicks(ByVal ticks As LongLong, Optional ByVal kind As DateTimeKind = DateTimeKind_Unspecified) As DateTime
        //Attribute CreateFromTicks.VB_Description = "Initializes a new instance of the DateTime structure to a specified number of ticks and to Coordinated Universal Time (UTC) or local time."
        //       With New DateTime
        //          Set .CreateFromTicks = this.DotNetLibDateTime.CreateFromTicks(ticks, kind)
        //       End With
        //End Function

        // Expected output for return types of current object wrapping
        // Public Function CreateFromTicks(ByVal ticks As LongLong, Optional ByVal kind As DateTimeKind = DateTimeKind_Unspecified) As DateTime
        //     With New DateTime
        //         Set .ComObject = this.DotNetLibDateTime.CreateFromTicks(ticks, kind)
        //         Set Create = .Self
        //     End With
        // End Function

        // Note for 
        //Public Property Get TimeOfDay() As ITimeSpan
        //Attribute TimeOfDay.VB_Description = "Gets the time of day for this instance."
        //   With New ITimeSpan
        //      Set .TimeOfDay = this.DotNetLibDateTime.TimeOfDay()
        //   End With
        //End Property

        //Expected output
        //Public Property Get TimeOfDay() As DotNetLib.TimeSpan
        //   With New DotNetLib.TimeSpan
        //       Set .TimeOfDay = this.DotNetLibDateTime.TimeOfDay()
        //   End With
        // End Property
        private String ComObjectVariableDeclaration() 
        {
            if (this.MethodInfo.Type == DeclarationType.Function || this.MethodInfo.Type == DeclarationType.PropertyGet)
            {
                string assignment = this._memberName + " = ";
                if (this.MethodInfo.AsTypeName.IsByRef)
                {
                    StringBuilder sb = new StringBuilder();
                    //get return object name make sure not default interface
                    sb.AppendLine($"{Indent}With New {ReturnType()}");  
                    sb.AppendLine($"{Indent}{Indent}Set .{assignment}{ComObjectWrapperReference()}");
                    sb.AppendLine($"{Indent}End With");

                    // Require return object eg Date how to object implementation from interface?
                    //TODO : create code block to create object to return
                    //TODO : Issue with return types being an interface
                    // EG.
                    //     With New DateTime
                    //          Set .CreateFromTicks = mDateTime.CreateFromTicks(ticks, kind)
                    //     End With
                    return sb.ToString();
                }
                else
                {
                    return $"{Indent}{assignment}{ComObjectWrapperReference()}" +Environment.NewLine; ;
                }
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

        // https://rubberduckvba.blog/2019/12/14/rubberduck-annotations/
        public string AnnotationEnumerator()
        {
            if (this.MethodInfo.IsEnumerator)
            {
                return "'@Enumerator";
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
