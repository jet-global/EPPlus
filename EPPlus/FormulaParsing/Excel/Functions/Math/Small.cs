using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	public class Small : ExcelFunction
	{
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentsAreValid(arguments, 2, out eErrorType argumentError) == false)
				return new CompileResult(argumentError);
			var args = arguments.ElementAt(0);
			var index = ArgToInt(arguments, 1) - 1;
			var values = ArgsToDoubleEnumerable(new List<FunctionArgument> { args }, context);
			if (index < 0 || index >= values.Count())
				return new CompileResult(eErrorType.Num);
			var result = values.OrderBy(x => x).ElementAt(index);
			return CreateResult(result, DataType.Decimal);
		}
	}
}
