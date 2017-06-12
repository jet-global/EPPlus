using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Information
{
	public class IsLogical : ExcelFunction
	{
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			var functionArguments = arguments as FunctionArgument[] ?? arguments.ToArray();
			if (this.ValidateArguments(arguments, 1) == false)
				return new CompileResult(eErrorType.Value);
			var v = GetFirstValue(arguments);
			return CreateResult(v is bool, DataType.Boolean);
		}
	}
}
