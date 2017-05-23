using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime.Workdays;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
	public class Workday : ExcelFunction
	{
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			var functionArguments = arguments as FunctionArgument[] ?? arguments.ToArray();
			if (ValidateArguments(arguments, 2) == false)
				return new CompileResult(eErrorType.Value);
			if (arguments.ElementAt(1).Value is null)
				return new CompileResult(eErrorType.NA);

			var serialNumberCandidate = arguments.ElementAt(0).Value;
			if (serialNumberCandidate == null)
				return new CompileResult(eErrorType.NA);
			else if (serialNumberCandidate is string)
			{
				var isDateString = ConvertUtil.TryParseDateString(serialNumberCandidate.ToString(), out System.DateTime output);
				if (!isDateString)
					return new CompileResult(eErrorType.Value);
			}
			else if (serialNumberCandidate is int)
				if (ArgToInt(functionArguments, 1) < 0)
					return new CompileResult(eErrorType.Num);

			var dateSerial = ArgToDecimal(arguments, 0);
			if (dateSerial < 0)
				return new CompileResult(eErrorType.Num);

			var workDaysCandidate = arguments.ElementAt(1).Value;
			if (workDaysCandidate is string)
			{
				var isWorkDayString = ConvertUtil.TryParseDateString(workDaysCandidate.ToString(), out System.DateTime output2);
				if (!isWorkDayString)
					return new CompileResult(eErrorType.Value);
			}
			

			if (dateSerial == 0)
				if (ArgToInt(arguments, 1) < 0)
					return new CompileResult(eErrorType.Num);

			var startDate = System.DateTime.FromOADate(dateSerial);

			var workDateSerial = ArgToDecimal(arguments, 1);

			var resultDate = System.DateTime.MinValue;

			var calculator = new WorkdayCalculator();
			var result = calculator.CalculateWorkday(startDate, (int)workDateSerial);




			if (functionArguments.Length > 2)
			{
				for (int i = 2; i < functionArguments.Length; i++)
				{
					var holidayCandidate = arguments.ElementAt(i).Value;
					if (holidayCandidate is string)
					{
						var isHolidayString = ConvertUtil.TryParseDateString(holidayCandidate, out System.DateTime output3);
						if (!isHolidayString)
							return new CompileResult(eErrorType.Value);
					}
					else if (holidayCandidate is int)
					{
						var holidaySerial = ArgToInt(arguments, i);
						if (holidaySerial < 0)
							return new CompileResult(eErrorType.Num);
						break;
					}
					result = calculator.AdjustResultWithHolidays(result, functionArguments[i]);
				}
			}



			if(dateSerial == 0)
				if(dateSerial % 5 == 0)
					return CreateResult(result.EndDate.ToOADate()-1, DataType.Date);
			return CreateResult(result.EndDate.ToOADate(), DataType.Date);
		}
	}
}
