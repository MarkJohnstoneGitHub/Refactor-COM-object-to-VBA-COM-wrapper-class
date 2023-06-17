using Rubberduck.VBEditor.ComManagement.TypeLibs.Utility;
using System;
using System.Collections.Generic;
using System.Runtime.Serialization;
using System.Text;

namespace ComRefactor.Refactoring.CodeBuilder.VBA
{
    [DataContract]
    public class CodeModuleHeaderAttributes
    {
        private Dictionary<string, object> Attributes = new Dictionary<string, object>();
        public CodeModuleHeaderAttributes(String moduleName, String description = null, bool isPredeclared = false) 
        {
            Attributes.Add("VB_Name", moduleName);  //TODO : validate cannot be null or empty and check VBA module name rules
            Attributes.Add("VB_Description", description);
            Attributes.Add("VB_GlobalNameSpace", false);
            Attributes.Add("VB_Creatable", false);
            Attributes.Add("VB_PredeclaredId", isPredeclared);
            Attributes.Add("VB_Exposed ", false);
        }

        public String CodeModule()
        {
            StringBuilder attributeHeader = new StringBuilder();

            foreach (var attribute in Attributes)
            {
                String attrributeValue;
                if  (attribute.Value != null) 
                {
                    if (attribute.Key == "VB_Name" || attribute.Key == "VB_Description" )
                    {
                        String quote = "\"";
                        attrributeValue = quote + attribute.Value + quote;
                    }
                    else
                    {
                        attrributeValue = attribute.Value.ToString();
                    }
                    attributeHeader.AppendLine("Attribute" + " " + attribute.Key + " = " + attrributeValue);
                }      
            }
            return attributeHeader.ToString();
        }



    }
}
