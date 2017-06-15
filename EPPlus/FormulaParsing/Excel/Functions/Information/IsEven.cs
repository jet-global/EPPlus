using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Information
{
	public class IsEven : ExcelFunction
	{
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentsAreValid(arguments, 1, out eErrorType argumentError) == false)
				return new CompileResult(argumentError);
			var arg1 = GetFirstValue(arguments);//arguments.ElementAt(0);
			if (!ConvertUtil.IsNumeric(arg1))
			{
				return new CompileResult(eErrorType.Value);
			}
			var number = (int)System.Math.Floor(ConvertUtil.GetValueDouble(arg1));
			return CreateResult(number % 2 == 0, DataType.Boolean);
		}
	}
}
