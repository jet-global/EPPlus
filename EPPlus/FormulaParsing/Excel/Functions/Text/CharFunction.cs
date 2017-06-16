using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Text
{
	public class CharFunction : ExcelFunction
	{
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentsAreValid(arguments, 1, out eErrorType argumentError) == false)
				return new CompileResult(argumentError);
			var number = ArgToInt(arguments, 0);
			if (number < 1 || number > 255)
				return new CompileResult(eErrorType.Value);
			return CreateResult(((char)number).ToString(), DataType.String);
		}
	}
}
