using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	public class Quotient : ExcelFunction
	{
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (ValidateArguments(arguments, 2) == false)
				return new CompileResult(eErrorType.Value);
			var num = ArgToDecimal(arguments, 0);
			var denom = ArgToDecimal(arguments, 1);
			if ((int)denom == 0)
				return new CompileResult(eErrorType.Div0);
			var result = (int)(num / denom);
			return CreateResult(result, DataType.Integer);
		}
	}
}
