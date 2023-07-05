using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;

namespace Rubberduck.Parsing.ComReflection
{
    //TODO : Rubberduck required GUID for TKIND_DISPATCH for ComParameter ??
    //TODO : Why Guid for Enum and Alias? why not jut one GUID for a parameter?
    //TODO : Probably requires refactoring with GUID property and TYPEKIND ?
    //TODO : See ComParameter

    [DataContract]
    [KnownType(typeof(ComProject))]
    public class ComTypeName
    {
        [DataMember(IsRequired = true)]
        public Guid EnumGuid { get; private set; } = Guid.Empty;  
        public bool IsEnumMember => !EnumGuid.Equals(Guid.Empty);

        [DataMember(IsRequired = true)]
        public Guid AliasGuid { get; private set; } = Guid.Empty;
        public bool IsAliased => !AliasGuid.Equals(Guid.Empty);

        public bool IsDispatch => !DispatchGuid.Equals(Guid.Empty);  // TODO : Rubberduck Added
        public Guid DispatchGuid { get; private set; } = Guid.Empty; // TODO : Rubberduck Added

        public ComProject Project { get; set; }

        [DataMember(IsRequired = true)]
        private string _rawName;
        public string Name
        {
            get
            {
                if (IsEnumMember && ComProject.KnownEnumerations.TryGetValue(EnumGuid, out var enumeration))
                {
                    return enumeration.Name;
                }

                if (IsAliased && ComProject.KnownAliases.TryGetValue(AliasGuid, out var alias))
                {
                    return alias.Name;
                }

                // TODO: Rubberduck added
                // Search the current project for the implementation of the interface
                // If only one implementation i.e. default implementation use that?
                if (IsDispatch && ComProject.KnownTypes.TryGetValue(DispatchGuid, out var dispatch))
                {
                    IEnumerable<ComCoClass> implementedInterface = Project.FindImplementedInterface(DispatchGuid);

                    if (implementedInterface != null)
                    {
                        if (implementedInterface.Count() == 1)
                        {
                            return $"{implementedInterface.First().Name}";
                        }
                        else
                        {
                            return dispatch.Name;
                        }
                    }
                    return dispatch.Name;
                }

                if (Project == null)
                {
                    return _rawName;
                }

                var softAlias = Project.Aliases.FirstOrDefault(x => x.Name.Equals(_rawName));
                return softAlias == null ? _rawName : softAlias.TypeName;
            }
        }

        public ComTypeName(ComProject project, string name)
        {
            Project = project;
            _rawName = name;
        }

        public ComTypeName(ComProject project, string name, Guid enumGuid, Guid aliasGuid) : this(project, name)
        {
            EnumGuid = enumGuid;
            AliasGuid = aliasGuid;
        }

        //TODO : Rubberduck added
        public ComTypeName(ComProject project, string name, Guid enumGuid, Guid aliasGuid, Guid dispatchGuid) : this(project, name)
        {
            EnumGuid = enumGuid;
            AliasGuid = aliasGuid;
            DispatchGuid = dispatchGuid;
        }
    }
}
