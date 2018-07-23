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
		/// Execute returns a value dummy value for the weekend parameter.
		/// </summary>
		/// <returns>The value 0 (but can be any value except 2).</returns>
		protected override int GetWeekendIndex()
		{
			return 0;
		}

		/// <summary>
		/// Execute returns whether holidays parameter is specified by user.
		/// </summary>
		/// <param name="functionArguments">The array of parameters for function</param>
		/// <returns>A boolean depending on whether or not the holiday parameter is given.</returns>
		protected override bool HolidaysGiven(FunctionArgument[] functionArguments)
		{
			if (functionArguments.Length > 2)
				return true;
			else
				return false;
		}

		/// <summary>
		/// Execute returns holiday parameter index.
		/// </summary>
		/// <returns>The index value 2 corresponding to holiday parameter index.</returns>
		protected override int GetHolidayIndex()
		{
			return 2;
		}
	}
}
