using Rubberduck.Parsing.ComReflection;

namespace ComRefactor.Refactoring.CodeBuilder.VBA
{
    public class CodeModuleParameter
    {
        CodeModuleMethod _parentMember;
        ComParameter _parameter;

        public string Name
        {
            get
            {
                //if parameter name equals coClassName rename to module name, module maybe the coClass.Name or a new module name
                if (_parameter.Type.IsDispatch)
                {
                    if (_parameter.Type.DispatchGuid == _parentMember.MethodInfo.Parent.Guid)
                    {
                        return _parentMember.ModuleName;
                    }
                    else
                    {
                        //TODO: Eg. convert ITimeSpan to TimeSpan
                        return _parameter.TypeName;
                    }

                }
                else
                {
                    return  _parameter.TypeName;
                }
                

                //if (this.MethodInfo.AsTypeName.Type.DispatchGuid == this.MethodInfo.Parent.Guid)
                //{

                //}

                //if (_parameter.TypeName == _parentMember.MethodInfo.Parent.Name)
                //{
                //    return _parentMember.ModuleName;
                //}
                //else
                //// TODO If (this._parameter.IsByRef) //replace with quantative name of object i.e. is the default interface
                //{
                //    //TODO : may require to rename for interface returned when only one implementation i.e. the default interface.  If cant find it's ComCoClass or has multiple implementations then use interface?
                //    return _parameter.TypeName;
                //}
            }
        }


        //require parent method info for a parameter
        public CodeModuleParameter(CodeModuleMethod parentMember,ComParameter parameter)
        {
            _parentMember = parentMember;
            _parameter = parameter;

        }

        //From ComParameter DeclarationName required to add public property Type to ComParameter to access required properties.
        //Test for if _parameter.IsArrray
        public override string ToString()
        {
            return  $"{(_parameter.IsOptional ? "Optional " : string.Empty)}{(_parameter.IsByRef ? "ByRef" : "ByVal")} {_parameter.Name} As {this.Name}{(_parameter.IsOptional && _parameter.DefaultValue != null ? " = " : string.Empty)}{(_parameter.IsOptional && _parameter.DefaultValue != null ? _parameter.Type.IsEnumMember ? _parameter.DefaultAsEnum : _parameter.DefaultValue : string.Empty)}";
        }

    }
}
