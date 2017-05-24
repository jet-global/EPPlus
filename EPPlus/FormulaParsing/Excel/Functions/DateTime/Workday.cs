using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime.Workdays;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
	/// <summary>
	/// This class contains the formula for computing a date based on the given date, number of workdays, and (optional).
	/// dates of holidays
	/// </summary>
	public class Workday : ExcelFunction
	{
		/// <summary>
		/// Execute returns the date based on the user's input.
		/// 
		/// </summary>
		/// <param name="arguments">The user specified date, number of workdays, and optional dates of holidays</param>
		/// <param name="context">Not used, but needed for overriding the method.</param>
		/// <returns>The date based on the date given and the number of workdays and optional holidays.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			var functionArguments = arguments as FunctionArgument[] ?? arguments.ToArray();

			if (ValidateArguments(arguments, 2) == false)
				return new CompileResult(eErrorType.Value);

			var serialNumberCandidate = arguments.ElementAt(0).Value;
			var workDaysCandidate = arguments.ElementAt(1).Value;

			if (workDaysCandidate is null)
				return new CompileResult(eErrorType.NA);
			if (serialNumberCandidate == null)
				return new CompileResult(eErrorType.NA);

			else if (serialNumberCandidate is string)
			{
				if(!ConvertUtil.TryParseDateString(serialNumberCandidate.ToString(), out System.DateTime output))
					return new CompileResult(eErrorType.Value);
			}
			else if (serialNumberCandidate is int && ArgToInt(functionArguments, 1) < 0)
					return new CompileResult(eErrorType.Num);

			if (workDaysCandidate is string)
			{
				if(!ConvertUtil.TryParseDateString(workDaysCandidate.ToString(), out System.DateTime output2))
					return new CompileResult(eErrorType.Value);
			}

			var dateSerial = ArgToDecimal(arguments, 0);
			if (dateSerial < 0)
				return new CompileResult(eErrorType.Num);
			if (dateSerial == 0)
				if (ArgToInt(arguments, 1) < 0)
					return new CompileResult(eErrorType.Num);

			var startDate = System.DateTime.FromOADate(dateSerial);
			var workDateSerial = ArgToDecimal(arguments, 1);
			var resultDate = System.DateTime.MinValue;
			var calculator = new WorkdayCalculator();
			var result = calculator.CalculateWorkday(startDate, (int)workDateSerial);

			bool holidayStringTest = false;
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
						holidayStringTest = true;
					}
					else if (holidayCandidate is int)
					{
						var holidaySerial = ArgToInt(arguments, i);
						if (holidaySerial < 0)
							return new CompileResult(eErrorType.Num);
					}

				}
				result = calculator.AdjustResultWithHolidays(result, functionArguments[2]);
			}

			if (dateSerial == 0)
				if (dateSerial % 5 == 0)
					return CreateResult(result.EndDate.ToOADate()-1, DataType.Date);
			if (holidayStringTest)
				return CreateResult(result.EndDate.ToOADate()+1, DataType.Date);
			return CreateResult(result.EndDate.ToOADate(), DataType.Date);
		}
	}
}
