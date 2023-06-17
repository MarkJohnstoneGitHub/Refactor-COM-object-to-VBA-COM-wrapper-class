using Rubberduck.Parsing.ComReflection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ComRefactorConsole.ComRefactor.Refactoring.CodeBuilder.VBA
{
    internal class CodeModuleParameters
    {
        private IEnumerable<ComParameter> _parameters;

        public CodeModuleParameters(IEnumerable<ComParameter> parameters)
        {
            _parameters = parameters;
        }

        public String Parameters()
        {
            List<String> declarationParameters = new List<String>();

            foreach (var parameter in this._parameters)
            {
                declarationParameters.Add(parameter.DeclarationName);
            }
            String joined = "(" + String.Join(", ", declarationParameters) + ")";
            return joined;
        }

    }
}
