using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
	/// <summary>
	/// Return the fraction of the year represented by the number of whole days between two dates.
	/// </summary>
	public class Yearfrac : ExcelFunction
	{
		/// <summary>
		/// Given two dates, represented as a string, int, double, or <see cref="System.DateTime"/>,
		/// and an optional type of day count basis to use, return the fraction of the year beteen those two dates.
		/// </summary>
		/// <param name="arguments">The arguments used to calculate the fraction of the year.</param>
		/// <param name="context">Used to determine the context for the Days360 function used in this method.</param>
		/// <returns>Returns the fraction of the year between the two given dates as a double, or an <see cref="ExcelErrorValue"/> if the input is invalid.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			var functionArguments = arguments as FunctionArgument[] ?? arguments.ToArray();
			if (this.ValidateArguments(functionArguments, 2) == false)
				return new CompileResult(eErrorType.Value);
			var basis = 0;
			if (functionArguments.Count() > 2)
			{
				if (!this.TryGetArgAsInt(functionArguments, 2, out basis))
					return new CompileResult(eErrorType.Value);
				else if (basis < 0 || basis > 4)
					return new CompileResult(eErrorType.Num);
			}
			var date1Candidate = functionArguments[0].Value;
			var date2Candidate = functionArguments[1].Value;
			if (date1Candidate == null || date2Candidate == null)
				return new CompileResult(eErrorType.NA);
			var date1IsValid = ConvertUtil.TryParseDateObjectToOADate(date1Candidate, out double date1Num);
			var date2IsValid = ConvertUtil.TryParseDateObjectToOADate(date2Candidate, out double date2Num);
			if (!date1IsValid)
				return new CompileResult(eErrorType.Value);
			else if (date1Num < 0)
				return new CompileResult(eErrorType.Num);
			else if (!date2IsValid)
				return new CompileResult(eErrorType.Value);
			else if (date2Num < 0)
				return new CompileResult(eErrorType.Num);
			// The date1Num and date2Num provided by TryParseDateObjectToOADate are Excel OADates;
			// they need to be converted back to System.DateTime OADates for the special case
			// of dates before 3/1/1900 (OADate 61 in both Excel and System.DateTime).
			if (date1Num < 61)
				date1Num++;
			if (date2Num < 61)
				date2Num++;
			if (date1Num > date2Num) // Ensure that date1Num always represents the earliest date.
			{
				var t = date1Num;
				date1Num = date2Num;
				date2Num = t;
			}
			var date1 = System.DateTime.FromOADate(date1Num);
			var date2 = System.DateTime.FromOADate(date2Num);
			functionArguments[0] = new FunctionArgument(date1Num);
			functionArguments[1] = new FunctionArgument(date2Num);
			var func = context.Configuration.FunctionRepository.GetFunction("days360");
			var daysBetween = (date2 - date1).TotalDays;
			// Account for the fact that Excel includes an extra day, 2/29/1900 which doesn't actually exist, 
			// between 2/28/1900 and 3/1/1900, which are represented by System.DateTime OADates 60 and 61 respectively.
			if (date1.ToOADate() < 61 && date2.ToOADate() >= 61)
				daysBetween++;
			var yearFracResult = 0d;
			switch (basis)
			{
				case 0:
					var d360Result = System.Math.Abs(func.Execute(functionArguments, context).ResultNumeric);
					// Reproducing Excel's behavior.
					if (date1.Month == 2 && date2.Day == 31)
					{
						var calendar = new GregorianCalendar();
						var daysInFeb = calendar.IsLeapYear(date1.Year) ? 29 : 28;
						if (date1.Day == daysInFeb)
							d360Result++;
					}
					yearFracResult = d360Result / 360d;
					break;
				case 1:
					yearFracResult = System.Math.Abs(daysBetween / this.CalculateAcutalYear(date1, date2));
					break;
				case 2:
					yearFracResult = System.Math.Abs(daysBetween / 360d);
					break;
				case 3:
					yearFracResult = System.Math.Abs(daysBetween / 365d);
					break;
				case 4:
					var args = functionArguments.ToList();
					args.Add(new FunctionArgument(true));
					double? result = System.Math.Abs(func.Execute(args, context).ResultNumeric / 360d);
					yearFracResult = result.Value;
					break;
				default:
					return new CompileResult(eErrorType.Num);
			}
			yearFracResult = System.Math.Round(yearFracResult, 11);
			return this.CreateResult(yearFracResult, DataType.Decimal);
		}

		/// <summary>
		/// Calculates the average number of days in a year for the years between two given dates.
		/// </summary>
		/// <param name="date1">The starting date for calculating the actual year value.</param>
		/// <param name="date2">The ending date for calculating the actual year value.</param>
		/// <returns>Returns the average number of days in a year based on the time between <paramref name="date1"/> and <paramref name="date2"/>.</returns>
		private double CalculateAcutalYear(System.DateTime date1, System.DateTime date2)
		{
			var calendar = new GregorianCalendar();
			var totalDaysPerYear = 0d;
			var numberOfYears = date2.Year - date1.Year + 1;
			if (new System.DateTime(date1.Year + 1, date1.Month, date1.Day) >= date2) // Check if date1 and date2 are a year or less apart.
			{
				numberOfYears = 1;
				totalDaysPerYear = 365;
				if (calendar.IsLeapYear(date1.Year) && date1.Month <= 2)
					totalDaysPerYear = 366;
				else if (calendar.IsLeapYear(date2.Year) && date2.Month > 2)
					totalDaysPerYear = 366;
				else if (date2.Month == 2 && date2.Day == 29)
					totalDaysPerYear = 366;
			}
			else
			{
				for (var y = date1.Year; y <= date2.Year; ++y)
				{
					totalDaysPerYear += calendar.IsLeapYear(y) ? 366 : 365;
				}
			}
			return totalDaysPerYear / numberOfYears;
		}
	}
}
