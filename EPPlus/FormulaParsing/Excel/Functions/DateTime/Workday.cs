using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime.Workdays;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
	/// <summary>
	/// This class contains the formula for computing a date based on the given date, number of workdays, and (optional).
	/// dates of holidays
	/// </summary>
	public class Workday : WorkdayIntl
	{
		#region Properties
		private int holidayIndex = 2;
		#endregion

		#region ExcelFunction Overrides
		/// <summary>
		/// Execute returns a calculator with default Saturday and Sunday as weekend.
		/// </summary>
		/// <param name="weekend">The user specified weekend code</param>
		/// <returns>A calculator with Saturday and Sunday as default weekend days.</returns>
		protected override WorkdayCalculator SetCalculator(object weekend)
		{
			return new WorkdayCalculator();
		}

		/// <summary>
		/// Execute returns whether or not there is a weekend parameter.
		/// </summary>
		/// <returns>False since Workday does not have a weekend parameter.</returns>
		protected override bool WeekendSpecified(FunctionArgument[] functionArguments)
		{
			return false;
		}

		/// <summary>
		/// Execute returns whether holidays parameter is specified by user.
		/// </summary>
		/// <param name="functionArguments">The array of parameters for function</param>
		/// <returns>A boolean depending on whether or not the holiday parameter is given.</returns>
		protected override bool HolidaysSpecified(FunctionArgument[] functionArguments)
		{
			return functionArguments.Length > 2;
		}

		/// <summary>
		/// Execute returns holiday parameter index.
		/// </summary>
		/// <returns>The index value 2 corresponding to holiday parameter index.</returns>
		protected override int GetHolidayIndex()
		{
			return holidayIndex;
		}

		#endregion
	}
}
