using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	public class Fact : ExcelFunction
	{
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (ValidateArguments(arguments, 1) == false)
				return new CompileResult(eErrorType.Value);
			var number = ArgToDecimal(arguments, 0);
			if (number < 0)
				return new CompileResult(eErrorType.NA);
			var result = 1d;
			for (var x = 1; x < number; x++)
			{
				result *= x;
			}
			return CreateResult(result, DataType.Integer);
		}
	}
}
