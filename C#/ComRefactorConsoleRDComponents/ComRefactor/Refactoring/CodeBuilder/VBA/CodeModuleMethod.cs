using Rubberduck.Parsing.ComReflection;
using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Text;

namespace ComRefactor.Refactoring.CodeBuilder.VBA
{

    // TODO : Issue with method names using VBA reserved words
    // TODO : Issue with Com method names in Lowercase quickfix done
    // TODO : Issue for interface being used when referring to the current Com Object being implemented
    // TODO : Check for parameters and return type equal to ComCoClass default interface
    // eg. 
    // TODO : Issue missing Function As clause

    // https://www.engram9.info/access-2007-vba/reserved-word-list.html

    public class CodeModuleMethod
    {
        private const string Quote = "\"";

        ComMember _methodInfo;
        String _memberName;
        String _comCoClassName; 

        public CodeModuleMethod(ComMember methodInfo, String comCoClassName)
        {
            this._methodInfo = methodInfo;
            this._memberName = FirstLetterToUpper(this._methodInfo.Name);
            this._comCoClassName = comCoClassName;
        }
        
        public string Name 
        {
            get => _memberName;
            private set => _memberName = value;
        } 


        //TODO: Missing function and property get As clause
        public String Signature()
        {
            return Declaration() + Parameters() + " " + ReturnType();
        }
        public String CodeModule() 
        {
            StringBuilder method = new StringBuilder();

            if (this._methodInfo.Documentation.DocString != null)
            {
                method.AppendLine(AnnotationMemberDescription());
            }
            method.AppendLine(Signature());
            if (this._methodInfo.Documentation.DocString != null)
            {
                method.AppendLine(DescriptionAttribute()); 
            }
            //TODO method attributes
            //TODO: reference to Com object  being wrapped
            method.AppendLine(DeclarationEnd());

            return method.ToString();
        }

        public String DeclarationEnd()
        {
            var type = string.Empty;
            switch (this._methodInfo.Type)
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

        public String Declaration()
        {
            var type = string.Empty;
            switch (this._methodInfo.Type)
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
            return $"{(this._methodInfo.IsHidden || this._methodInfo.IsRestricted ? "Private" : "Public")} {type} {this.Name}";
        }

        public String DescriptionAttribute()
        {
            string output = String.Empty;
            if (this._methodInfo.Documentation.DocString != null)
            {
                //Attribute Example.VB_Description = "This is a description of a procedure"
                output = "Attribute " + Name + ".VB_Description" + " = " + Quote + this._methodInfo.Documentation.DocString + Quote;
            }
            return output;
        }


        // TODO : Issue with optional parameters if equal null 
        // TODO : Move to CodeModuleParrameters
        public String Parameters()
        {
            List<String> declarationParameters = new List<String>();

            foreach (var parameter in this._methodInfo.Parameters)
            {
                String declarationName;
                if (parameter.TypeName == this._methodInfo.Parent.Name)
                {
                    string parameterTypeName = this._comCoClassName;

                    declarationName = parameter.DeclarationName; // TODO: update to class name i.e. default is the  comCoClass.Name or new module name ???

                    //See ComParameter
                    //declarationName = $"{(parameter.IsOptional ? "Optional " : string.Empty)}{(parameter.IsByRef ? "ByRef" : "ByVal")} {parameter.Name} As {parameterTypeName}{(parameter.IsOptional && parameter.DefaultValue != null ? " = " : string.Empty)}{(parameter.IsOptional && parameter.DefaultValue != null ? _typeName.IsEnumMember ? DefaultAsEnum : DefaultValue : string.Empty)}";
                }
                else
                {
                    declarationName = parameter.DeclarationName;
                }

                declarationParameters.Add(declarationName);
            }
            String joined = "(" + String.Join(", ", declarationParameters) + ")";
            return joined;
        }


        // TODO : Issue with return types being interface, how to handle for other interfaces returned that are in the current library?
        // TODO : Posssible pass in the Com library object to check and replace with qualified name eg. DotNetLib.TimeSpan? or keep as ITimeSpan?
        public String ReturnType()
        {
            if (this._methodInfo.Type == DeclarationType.Function || this._methodInfo.Type == DeclarationType.PropertyGet)
            {
                string returnType = this._methodInfo.AsTypeName.TypeName;

                //If the return type is the interface name replace with comCoClassName/new class name  ???
                if (returnType != null) 
                { 
                    if (returnType == this._methodInfo.Parent.Name)
                    {
                        returnType = this._comCoClassName;  //If new class name selected use instead of default comCoClassName
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

        //https://github.com/rubberduck-vba/Rubberduck/wiki/VB_Attribute-Annotations#member-annotations
        public string AnnotationMemberDescription()
        {
            return "'@Description(\"" + this._methodInfo.Documentation.DocString + "\")";
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
