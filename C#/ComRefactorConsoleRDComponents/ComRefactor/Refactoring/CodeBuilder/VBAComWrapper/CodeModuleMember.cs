using Rubberduck.Parsing.ComReflection;
using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using System.Linq;
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
    // Easiest way to fully implementedInterface wrapping a Com Object would require access to any external type libraries referenced?
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

    public class CodeModuleMember
    {
        private const string Quote = "\"";
        const string Indent = "   ";

        String _memberName;
        String _comObjectVariable;


        public ComProject Project { get; private set; }
        public ComMember Member;

        public String ModuleName; 

        List<CodeModuleParameter> _parametersCode = new List<CodeModuleParameter>();

        public string ParametersDeclaration => "(" + String.Join(", ", _parametersCode) + ")";

        public IEnumerable<CodeModuleParameter> ParametersCode => this._parametersCode;

        private IEnumerable<ComParameter> _parameters => this.Member.Parameters;


        public CodeModuleMember(ComProject project, ComMember member, String moduleName, String comObjectVariable)
        {
            this.Project = project;
            this.Member = member;
            this._memberName = FirstLetterToUpper(this.Member.Name);
            this.ModuleName = moduleName;
            this._comObjectVariable = comObjectVariable;
            CodeParamaters();
        }

        public string Name 
        {
            get => _memberName;
            private set => _memberName = value;
        } 

        public String Signature()
        {
            return Declaration() + ParametersDeclaration + (ReturnType() == null ? string.Empty : " As " + ReturnType());
        }
        public String CodeModule() 
        {
            StringBuilder method = new StringBuilder();

            if (this.Member.IsDefault)
            {
                method.AppendLine(AnnotationDefaultMember());
            }

            if (this.Member.IsEnumerator)
            {
                method.AppendLine(AnnotationEnumerator());
            }

            if (this.Member.Documentation.DocString != null)
            {
                method.AppendLine(AnnotationMemberDescription());
            }

            method.AppendLine(Signature());

            // Member attributes
            if (this.Member.Documentation.DocString != null)
            {
                method.AppendLine(AttributeDescription()); 
            }
            if (this.Member.IsDefault)
            {
                method.AppendLine(AttributeDefaultMember());
            }

            if (this.Member.IsEnumerator)
            {
                method.AppendLine(AttributeEnumerator());
            }

            if (this._comObjectVariable != null) 
            {
                method.Append(ComObjectVariableDeclaration());
            }

            method.AppendLine(DeclarationEnd());

            return method.ToString();
        }

        public String Declaration()
        {
            var type = string.Empty;
            switch (this.Member.Type)
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
            return $"{(this.Member.IsHidden || this.Member.IsRestricted ? "Private" : "Public")} {type} {this.Name}";
        }

        public String DeclarationEnd()
        {
            var type = string.Empty;
            switch (this.Member.Type)
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
            if (this.Member.IsDefault)
            {
                return $"Attribute {Name}.VB_UserMemId = 0";
            }
            return String.Empty;
        }

        public String AttributeDescription()
        {
            if (this.Member.Documentation.DocString != null)
            {
                return $"Attribute {Name}.VB_Description = {Quote}{this.Member.Documentation.DocString}{Quote}";
            }
            return String.Empty;
        }

        public String AttributeEnumerator()
        {
            if (this.Member.IsEnumerator)
            {
                return $"Attribute {Name}.VB_UserMemId = -4";
            }
            return String.Empty;
        }

        private void CodeParamaters()
        {
            foreach (var parameter in this.Member.Parameters)
            {
                _parametersCode.Add(new CodeModuleParameter(this, parameter)); 
            }
        }

        public String ReturnType()
        {
            if (this.Member.Type == DeclarationType.Function || this.Member.Type == DeclarationType.PropertyGet)
            {
                string returnType = this.Member.AsTypeName.TypeName;

                if (this.Member.AsTypeName.Type.IsDispatch)
                {
                    //If function/PropertyGet return type GUID equals COM object GUID then return type is new module name
                    if (this.Member.AsTypeName.Type.DispatchGuid == this.Member.Parent.Guid)
                    {
                        returnType = this.ModuleName;
                    }
                    else
                    {
                        IEnumerable<ComCoClass> implementedInterface = this.Member.Project.FindImplementedInterface(this.Member.AsTypeName.Type.DispatchGuid);
                        if (implementedInterface != null)
                        {
                            if (implementedInterface.Count() == 1)
                            {
                                //TODO : Issue using qualified name and variable name
                                //returnType = $"{this.Member.AsTypeName.Type.Project.Name}.{implementedInterface.First().Name}"; //qualified name of IDispatch object
                                returnType = $"{implementedInterface.First().Name}";
                            }
                            else
                            {
                                returnType = this.Member.AsTypeName.TypeName;
                            }
                        }
                    }
                }
                else
                {
                    returnType = this.Member.AsTypeName.TypeName;
                }

                if (this.Member.AsTypeName.IsArray)
                {
                    returnType = returnType + "()";
                }
                return returnType;
            }
            return null;
        }

        private String ComObjectWrapperReference()
        {
            List<String> parameterNames = new List<String>();

            foreach (var parameter in this._parameters)
            {
                //i.e. if parameter type is Com object being wrapped
                if (parameter.Type.DispatchGuid ==  this.Member.Parent.Guid)
                {
                    parameterNames.Add($"{parameter.Name}.ComObject");
                }
                else 
                { 
                    parameterNames.Add(parameter.Name);
                }
            }
            String joinParameters = "(" + String.Join(", ", parameterNames) + ")";
            return $"{this._comObjectVariable}.{this.Member.Name}{joinParameters}"; 
        }

        private String ComObjectVariableDeclaration() 
        {
            if (this.Member.Type == DeclarationType.Function || this.Member.Type == DeclarationType.PropertyGet)
            {
                if (this.Member.AsTypeName.IsByRef)
                {
                    StringBuilder sb = new StringBuilder();
                    // if member return type is same as COM object being wrapped.
                    if (this.Member.AsTypeName.Type.DispatchGuid == this.Member.Parent.Guid)
                    {
                        sb.AppendLine($"{Indent}With New {ReturnType()}");
                        sb.AppendLine($"{Indent}{Indent}Set .ComObject = {ComObjectWrapperReference()}");
                        sb.AppendLine($"{Indent}{Indent}Set {this._memberName} = .Self");
                        sb.AppendLine($"{Indent}End With");
                    }
                    else
                    {
                        sb.AppendLine($"{Indent}Set {this._memberName}  = {ComObjectWrapperReference()}");
                    }
                    return sb.ToString();
                }
                else
                {
                    return $"{Indent}{this._memberName} = {ComObjectWrapperReference()}" + Environment.NewLine;
                }
            }
            return string.Empty;
        }

        //https://github.com/rubberduck-vba/Rubberduck/wiki/VB_Attribute-Annotations#member-annotations
        public string AnnotationMemberDescription()
        {
            return "'@Description(\"" + this.Member.Documentation.DocString + "\")";
        }

        public string AnnotationDefaultMember()
        {
            if (this.Member.IsDefault)
            {
                return "'@DefaultMember";
            }
            return String.Empty;

        }

        // https://rubberduckvba.blog/2019/12/14/rubberduck-annotations/
        public string AnnotationEnumerator()
        {
            if (this.Member.IsEnumerator)
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
