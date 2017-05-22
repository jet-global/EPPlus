using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime.Workdays;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
	public class Workday : ExcelFunction
	{
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			var functionArguments = arguments as FunctionArgument[] ?? arguments.ToArray();
			if (ValidateArguments(arguments, 2) == false)
				return new CompileResult(eErrorType.Value);
			
			var serialNumberCandidate = arguments.ElementAt(0).Value;
			if (serialNumberCandidate == null)
				return new CompileResult(eErrorType.NA);
			else if (serialNumberCandidate is string)
			{
				var isDateString = System.DateTime.TryParse(serialNumberCandidate.ToString(),out System.DateTime output);
				if (!isDateString)
					return new CompileResult(eErrorType.Value);
			}
			var dateSerial = ArgToDecimal(arguments, 0);
			if (dateSerial < 0)
				return new CompileResult(eErrorType.Num);
			var startDate = System.DateTime.FromOADate(dateSerial);
			var nWorkDays = ArgToInt(functionArguments, 1);
			var resultDate = System.DateTime.MinValue;

			var calculator = new WorkdayCalculator();
			var result = calculator.CalculateWorkday(startDate, nWorkDays);
			if (functionArguments.Length > 2)
			{
				result = calculator.AdjustResultWithHolidays(result, functionArguments[2]);
			}

			if(dateSerial == 0)
				if(dateSerial % 5 == 0)
					return CreateResult(result.EndDate.ToOADate()-1, DataType.Date);
			return CreateResult(result.EndDate.ToOADate(), DataType.Date);
		}
	}
}
