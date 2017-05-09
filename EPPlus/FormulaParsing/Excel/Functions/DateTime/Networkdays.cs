using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime.Workdays;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
	public class Networkdays : ExcelFunction
	{
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			var functionArguments = arguments as FunctionArgument[] ?? arguments.ToArray();
			if (ValidateArguments(arguments, 2) == false)
				return new CompileResult(eErrorType.Value);
			var startDate = System.DateTime.FromOADate(ArgToInt(functionArguments, 0));
			var endDate = System.DateTime.FromOADate(ArgToInt(functionArguments, 1));
			var calculator = new WorkdayCalculator();
			var result = calculator.CalculateNumberOfWorkdays(startDate, endDate);
			if (functionArguments.Length > 2)
			{
				result = calculator.ReduceWorkdaysWithHolidays(result, functionArguments[2]);
			}

			return new CompileResult(result.NumberOfWorkdays, DataType.Integer);
		}
	}
}
