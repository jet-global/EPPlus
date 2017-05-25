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
		/// </summary>
		/// <param name="arguments">The user specified date, number of workdays, and optional dates of holidays</param>
		/// <param name="context">Not used, but needed for overriding the method.</param>
		/// <returns>The date based on the date given and the number of workdays and optional holidays.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			System.DateTime output;
			var functionArguments = arguments as FunctionArgument[] ?? arguments.ToArray();

			if (ValidateArguments(arguments, 2) == false)
				return new CompileResult(eErrorType.Value);

			var serialNumberCandidate = arguments.ElementAt(0).Value;
			var workDaysCandidate = arguments.ElementAt(1).Value;

			if (workDaysCandidate == null)
				return new CompileResult(eErrorType.NA);
			if (serialNumberCandidate == null)
				return new CompileResult(eErrorType.NA);

			bool isSerialIntZero = (serialNumberCandidate is int dateAsInt && dateAsInt == 0);
			bool isSerialDoubleZero = (serialNumberCandidate is double dateAsDouble && dateAsDouble == 0.00);

			if (ConvertUtil.TryParseDateObject(serialNumberCandidate, out output, out eErrorType? error) || isSerialIntZero || isSerialDoubleZero)
			{
				if (serialNumberCandidate is int && ArgToInt(functionArguments, 1) < 0)
					return new CompileResult(eErrorType.Num);

				if (workDaysCandidate is string)
				{
					if (!ConvertUtil.TryParseDateString(workDaysCandidate.ToString(), out output))
						return new CompileResult(eErrorType.Value);
				}

				var dateSerial = this.ArgToDecimal(arguments, 0);
				if (dateSerial < 0)
					return new CompileResult(eErrorType.Num);
				if (isSerialIntZero && this.ArgToInt(arguments, 1) < 0)
						return new CompileResult(eErrorType.Num);
				if (isSerialDoubleZero && this.ArgToInt(arguments, 1) < 0)
					return new CompileResult(eErrorType.Num);

				var startDate = System.DateTime.FromOADate(dateSerial);
				var workDateSerial = this.ArgToDecimal(arguments, 1);
				var resultDate = System.DateTime.MinValue;
				var calculator = new WorkdayCalculator();
				var dateResult = calculator.CalculateWorkday(startDate, (int)workDateSerial);

				if (functionArguments.Length > 2)
				{
					for (int i = 2; i < functionArguments.Length; i++)
					{ 
						var holidayCandidate = arguments.ElementAt(i).Value;
						bool isHolidayZero = (serialNumberCandidate is int holAsint && holAsint == 0);

						if (holidayCandidate is int holAsInt && holAsInt < 0)
							return new CompileResult(eErrorType.Num);
						
						if (holidayCandidate is string && !ConvertUtil.TryParseDateString(holidayCandidate, out output))
							return new CompileResult(eErrorType.Value);

						dateResult = calculator.AdjustResultWithHolidays(dateResult, functionArguments[2]);
					}
				}
				if(isSerialIntZero || isSerialDoubleZero)
					return CreateResult(dateResult.EndDate.ToOADate()-1, DataType.Date);
				return CreateResult(dateResult.EndDate.ToOADate(), DataType.Date);
			}
			else
				return new CompileResult(error.Value);
		}
	}
}
