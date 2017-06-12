using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	public class Large : ExcelFunction
	{
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ValidateArguments(arguments, 2) == false)
				return new CompileResult(eErrorType.Value);
			var args = arguments.ElementAt(0);
			var index = ArgToInt(arguments, 1) - 1;
			var values = ArgsToDoubleEnumerable(new List<FunctionArgument> { args }, context);
			if (index < 0 || index >= values.Count())
				return new CompileResult(eErrorType.Num);
			var result = values.OrderByDescending(x => x).ElementAt(index);
			return CreateResult(result, DataType.Decimal);
		}
	}
}
