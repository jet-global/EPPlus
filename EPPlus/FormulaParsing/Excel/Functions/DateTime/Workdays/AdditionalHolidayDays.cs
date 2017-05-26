using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime.Workdays
{
	/// <summary>
	/// This class takes a FunctionArgument list of dates and turns them into a list of DateTime objects. 
	/// </summary>
	public class AdditionalHolidayDays
	{
		#region Properties
		private List<System.DateTime> HolidayDates { get; } = new List<System.DateTime>();
		public IEnumerable<System.DateTime> AdditionalDates;
		#endregion

		#region Constructor
		/// <summary>
		/// The constructor for this class calls the method Initalize, which does the transformation of the list. 
		/// </summary>
		/// <param name="holidayArg">The list of FunctionArguments to be turned into DateTime objects.</param>
		public AdditionalHolidayDays(FunctionArgument holidayArg)
		{
			this.AdditionalDates = HolidayDates;

			var holidays = holidayArg.Value as IEnumerable<FunctionArgument>;

			System.DateTime date;
			eErrorType? error;

			if (holidays != null)
			{
				foreach (var holiday in holidays)
				{
					if (ConvertUtil.TryParseDateObject(holiday.Value, out date, out error))
						this.HolidayDates.Add(date);
				}
			}

			var range = holidayArg.Value as ExcelDataProvider.IRangeInfo;
			if (range != null)
			{
				foreach (var cell in range)
				{
					if (ConvertUtil.TryParseDateObject(cell.Value, out date, out error))
						this.HolidayDates.Add(date);
				}
			}

			if (ConvertUtil.TryParseDateObject(holidayArg.Value, out date, out error))
			{
				this.HolidayDates.Add(date);
			}
		}
		#endregion
	}
}