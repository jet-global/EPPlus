using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime.Workdays;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
	/// <summary>
	/// This class contains the formula for computing the week number based on the ISO calendar given a date. 
	/// </summary>
	public class IsoWeekNum : ExcelFunction
	{
		#region ExcelFunctionOverrides 
		/// <summary>
		/// Execute returns the week number based on the ISO calendar based on the user's input.
		/// </summary>
		/// <param name="arguments">The user specified date</param>
		/// <param name="context">Not used, but needed to overload the method. The context in which the function is being executed.</param>
		/// <returns>The week number based on the given date.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ValidateArguments(arguments, 1) == false)
				return new CompileResult(eErrorType.Value);
			if (arguments.Count() > 1)
				return new CompileResult(eErrorType.NA);

			var dateCandidate = arguments.ElementAt(0);
			if (ConvertUtil.TryParseDateObject(dateCandidate.Value, out System.DateTime date1, out eErrorType? error))
			{
				return this.CreateResult(this.WeekNumber(date1), DataType.Integer);
			}
			else if (dateCandidate.Value is int)
			{
				var dateInt = this.ArgToInt(arguments, 0);
				if (dateInt < 0)
					return new CompileResult(eErrorType.Num);
				var date = System.DateTime.FromOADate(dateInt);
				return this.CreateResult(WeekNumber(date), DataType.Integer);
			}
			else
				return new CompileResult(eErrorType.Value);
		}
		#endregion
		#region Private Methods
		/// <summary>
		/// This implementation was found on http://stackoverflow.com/questions/1285191/get-week-of-date-from-linq-query
		/// </summary>
		/// <param name="fromDate">The date of which the week number will be determined.</param>
		/// <returns>The week number based on the given date.</returns>
		private int WeekNumber(System.DateTime fromDate)
		{
			// Get jan 1st of the year
			var startOfYear = fromDate.AddDays(-fromDate.Day + 1).AddMonths(-fromDate.Month + 1);
			// Get dec 31st of the year
			var endOfYear = startOfYear.AddYears(1).AddDays(-1);
			// ISO 8601 weeks start with Monday
			// The first week of a year includes the first Thursday
			// DayOfWeek returns 0 for sunday up to 6 for saterday
			int[] iso8601Correction = { 6, 7, 8, 9, 10, 4, 5 };
			int nds = fromDate.Subtract(startOfYear).Days + iso8601Correction[(int)startOfYear.DayOfWeek];
			int wk = nds / 7;
			switch (wk)
			{
				case 0:
					// Return weeknumber of dec 31st of the previous year
					return WeekNumber(startOfYear.AddDays(-1));
				case 53:
					// If dec 31st falls before thursday it is week 01 of next year
					if (endOfYear.DayOfWeek < DayOfWeek.Thursday)
						return 1;
					return wk;
				default: return wk;
			}
		}
		#endregion
	}
}
