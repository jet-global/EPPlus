using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	public class Fact : ExcelFunction
	{
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ValidateArguments(arguments, 1) == false)
				return new CompileResult(eErrorType.Value);
			if (arguments.ElementAt(0).ValueIsExcelError)
				return new CompileResult(arguments.ElementAt(0).ValueAsExcelErrorValue);
			if (!ConvertUtil.TryParseDateObjectToOADate(arguments.ElementAt(0).Value, out double parsedNumberAsDouble))
				return new CompileResult(eErrorType.Value);
			var number = (int)parsedNumberAsDouble;
			if (number < 0)
				return new CompileResult(eErrorType.Num);
			var result = 1d;
			for (var x = 1; x <= number; x++)
			{
				result *= x;
			}
			return CreateResult(result, DataType.Integer);
		}
	}
}
