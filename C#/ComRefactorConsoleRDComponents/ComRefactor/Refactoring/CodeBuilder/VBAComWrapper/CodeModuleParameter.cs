using Rubberduck.Parsing.ComReflection;
using System;
using System.Reflection;

namespace ComRefactor.Refactoring.CodeBuilder.VBA
{
    public class CodeModuleParameter
    {
        CodeModuleMember _parentMember;
        ComParameter _parameter;

        public string Name
        {
            get
            {

                if (_parameter.Type.IsDispatch)
                {
                    //if parameter name equals coClassName rename to module name, module maybe the coClass.Name or a new module name
                    if (_parameter.Type.DispatchGuid == _parentMember.Member.Parent.Guid)
                    {
                        return _parentMember.ModuleName;
                    }
                    else
                    {
                        //Search KnowTypes using GUID if exists in type library
                        if (ComProject.KnownTypes.TryGetValue(_parameter.Type.DispatchGuid, out var type))
                        {
                            return (String.IsNullOrEmpty(type.Project.Name) ? _parameter.TypeName : $"{_parameter.Project.Name}.{_parameter.TypeName}");
                        }
                        else
                        {
                            return _parameter.TypeName;
                        }
                    }
                }
                else if (this._parameter.Type.IsEnumMember)
                {
                    //Search KnowTypes using GUID if exists in type library
                    if (ComProject.KnownTypes.TryGetValue(_parameter.Type.EnumGuid, out var type))
                    {
                        return (String.IsNullOrEmpty(type.Project.Name) ? _parameter.TypeName : $"{_parameter.Project.Name}.{_parameter.TypeName}");
                    }
                    else 
                    { 
                        return _parameter.TypeName; 
                    }
                }
                else
                {
                    return _parameter.TypeName;
                }
            }
        }

        public CodeModuleParameter(CodeModuleMember parentMember,ComParameter parameter)
        {
            _parentMember = parentMember;
            _parameter = parameter;

        }

        public override string ToString()
        {
            return  $"{(_parameter.IsOptional ? "Optional " : string.Empty)}{(_parameter.IsByRef ? "ByRef" : "ByVal")} {_parameter.Name}{(_parameter.IsArray ? "()" : string.Empty)} As {this.Name}{(_parameter.IsOptional && _parameter.DefaultValue != null ? " = " : string.Empty)}{(_parameter.IsOptional && _parameter.DefaultValue != null ? _parameter.Type.IsEnumMember ? _parameter.DefaultAsEnum : _parameter.DefaultValue : string.Empty)}";
        }

    }
}
