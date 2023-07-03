using Rubberduck.Parsing.ComReflection;

namespace ComRefactor.Refactoring.CodeBuilder.VBA
{
    public class CodeModuleParameter
    {
        ComParameter _comParameter;
        CodeModuleMethod _codeMethod;  //parent of parameter

        //paramater name
        public string Name
        {
            get
            {
                //if parameter name equals coClassName rename to module name, module maybe the coClass.Name or a new module name
                if (_comParameter.TypeName == _codeMethod.MethodInfo.Parent.Name)
                {
                    return _codeMethod.ModuleName;
                }
                else
                // TODO If (this._comParameter.IsByRef) //replace with quantative name of object i.e. is the default interface
                {
                    //TODO : may require to rename for interface returned when only one implementation i.e. the default interface.  If cant find it's ComCoClass or has multiple implementations then use interface?
                    return _comParameter.TypeName;
                }
            }
        }


        //require parent method info for a parameter
        public CodeModuleParameter(ComParameter comParameter, CodeModuleMethod parentCodeModuleMethod)
        {
            _comParameter = comParameter;
            _codeMethod = parentCodeModuleMethod;
        }

        //From ComParameter DeclarationName required to add public property Type to ComParameter to access required properties.
        //Test for if _comParameter.IsArrray
        public override string ToString()
        {
            return  $"{(_comParameter.IsOptional ? "Optional " : string.Empty)}{(_comParameter.IsByRef ? "ByRef" : "ByVal")} {_comParameter.Name} As {this.Name}{(_comParameter.IsOptional && _comParameter.DefaultValue != null ? " = " : string.Empty)}{(_comParameter.IsOptional && _comParameter.DefaultValue != null ? _comParameter.Type.IsEnumMember ? _comParameter.DefaultAsEnum : _comParameter.DefaultValue : string.Empty)}";
        }

    }
}
