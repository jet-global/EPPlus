using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
	/// <summary>
	/// Return the Excel OADate for the date that is the indicated number of months before or after
	/// the given start date.
	/// </summary>
	public class Edate : ExcelFunction
	{
		/// <summary>
		/// Given a date represented as a string, int, double, or <see cref="System.DateTime"/> object and a number
		/// represented as a string, double, or int, return the Excel OADate for the date that is the given number
		/// of months before or after the given date.
		/// </summary>
		/// <param name="arguments">The given arguments used to calculate the new date.</param>
		/// <param name="context">Unused in the method, but necessary to override the method.</param>
		/// <returns>Returns the Excel OADate of the calculated date, or an <see cref="ExcelErrorValue"/> if either input is invalid.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (!this.ValidateArguments(arguments, 2))
				return new CompileResult(eErrorType.Value);
			var dateObj = arguments.ElementAt(0).Value;
			var monthsObj = arguments.ElementAt(1).Value;
			if (dateObj == null || monthsObj== null)
				return new CompileResult(eErrorType.NA);
			if (dateObj is string dateString && Double.TryParse(dateString, out double dateAsNumber))
				dateObj = dateAsNumber;
			var isZeroDate = ((dateObj is int dateInt && dateInt == 0) ||
				(dateObj is double dateDouble && dateDouble < 1 && dateDouble >= 0));
			if (ConvertUtil.TryParseDateObject(dateObj, out System.DateTime validDate, out eErrorType? error) || isZeroDate)
			{
				var resultDate = new System.DateTime(validDate.Year, validDate.Month, validDate.Day);
				if (isZeroDate)
					resultDate = new System.DateTime(1899, 12, 31);
				if (monthsObj is string monthsString)
				{
					if (Double.TryParse(monthsString, out double numberMonths))
						monthsObj = numberMonths;
					else
						return new CompileResult(eErrorType.Value);
				}
				var monthsToAdd = (monthsObj is double monthsDouble) ? (int)monthsDouble : (int)monthsObj;
				resultDate = resultDate.AddMonths(monthsToAdd);
				var serialNumber = resultDate.ToOADate();
				if (serialNumber < 61)
					serialNumber--;
				// The OADate 0 with 0 months added is a special case.
				if (serialNumber == 0 && isZeroDate)
					return this.CreateResult(0.0, DataType.Date);
				else if (serialNumber < 1)
					return new CompileResult(eErrorType.Num);
				else
					return this.CreateResult(serialNumber, DataType.Date);
			}
			else
				return new CompileResult(error.Value);
		}
	}
}
