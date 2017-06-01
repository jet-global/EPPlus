using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
	/// <summary>
	/// This class contains the formula for computing the number of dates between two given dates. 
	/// </summary>
	public class Days360 : ExcelFunction
	{
		private enum Days360Calctype
		{
			European,
			Us
		}
		#region Public ExcelFunction Overrides
		/// <summary>
		/// Execute returns the number of days between two user specified dates using the US or European date standards.
		/// </summary>
		/// <param name="arguments">The user specified dates and method of calendar they want to use. </param>
		/// <param name="context">Not used, but needed to override the method.</param>
		/// <returns></returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ValidateArguments(arguments, 2) == false)
				return new CompileResult(eErrorType.Value);
			if (!this.TryGetArgumentDateValueAtIndex(arguments, 0, out System.DateTime dt1) || !this.TryGetArgumentDateValueAtIndex(arguments, 1, out System.DateTime dt2))
				return new CompileResult(eErrorType.Value);
			var calcType = Days360Calctype.Us;

			if (arguments.Count() > 2)
			{
				var european = false;
				if (arguments.ElementAt(2).Value is string stringVal)
				{
					if (!stringVal.Equals("true", System.StringComparison.CurrentCultureIgnoreCase) && !stringVal.Equals("false", System.StringComparison.CurrentCultureIgnoreCase))
						return new CompileResult(eErrorType.Value);
				}
					
				if (arguments.ElementAt(2).Value is string && ConvertUtil.TryParseNumericString(arguments.ElementAt(2).Value, out double val))
					return new CompileResult(eErrorType.Value);
				if(arguments.ElementAt(2).Value is System.DateTime)
				{
					european = true;
				}
				else
				{
					european = this.ArgToBool(arguments, 2);
					if (arguments.ElementAt(2).Value is int intval && intval > 0)
						european = true;
				}
				if (european) calcType = Days360Calctype.European;
			}

			var startYear = dt1.Year;
			var startMonth = dt1.Month;
			var startDay = dt1.Day;
			var endYear = dt2.Year;
			var endMonth = dt2.Month;
			var endDay = dt2.Day;

			if (calcType == Days360Calctype.European)
			{
				if (startDay == 31) startDay = 30;
				if (endDay == 31) endDay = 30;
			}
			else
			{
				var calendar = new GregorianCalendar();
				var nDaysInFeb = calendar.IsLeapYear(dt1.Year) ? 29 : 28;

				// If the investment is EOM and (Date1 is the last day of February) and (Date2 is the last day of February), then change D2 to 30.
				if (startMonth == 2 && startDay == nDaysInFeb && endMonth == 2 && endDay == nDaysInFeb)
				{
					endDay = 30;
				}
				// If the investment is EOM and (Date1 is the last day of February), then change D1 to 30.
				if (startMonth == 2 && startDay == nDaysInFeb)
				{
					startDay = 30;
				}
				// If D2 is 31 and D1 is 30 or 31, then change D2 to 30 unless first argument is null (because Excel adds a day in this case).
				if (endDay == 31 && (startDay == 30 || startDay == 31) && arguments.First().Value != null)
				{
					endDay = 30;
				}
				// If D1 is 31, then change D1 to 30.
				if (startDay == 31)
				{
					startDay = 30;
				}
			}
			var result = (endYear * 12 * 30 + endMonth * 30 + endDay) - (startYear * 12 * 30 + startMonth * 30 + startDay);
			return this.CreateResult(result, DataType.Integer);
		}
		#endregion

		#region Private Methods
		/// <summary>
		/// Takes a list of arguments and tries to get a date value out of a specified value from the argument list. 
		/// </summary>
		/// <param name="arguments">The list of arguments.</param>
		/// <param name="index">The location of the value we want to try to get a date value from.</param>
		/// <param name="date">The output date value.</param>
		/// <returns>True and the date value if we can get a date value from the argument at the specified index.</returns>
		private bool TryGetArgumentDateValueAtIndex(IEnumerable<FunctionArgument> arguments, int index, out System.DateTime date)
		{
			try
			{
				double dateNumber = (arguments.ElementAt(index).Value == null) ? 0 : this.ArgToDecimal(arguments, index);
				date = System.DateTime.FromOADate(dateNumber);
				return true;
			}
			catch
			{
				date = new System.DateTime();
				return false;
			}
		}
		#endregion
	}
}
