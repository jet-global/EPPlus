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
		private readonly FunctionArgument _holidayArg;
		private List<System.DateTime> HolidayDates { get; } = new List<System.DateTime>();
		/// <summary>
		/// The constructor for this class calls the method Initalize, which does the transformation of the list. 
		/// </summary>
		/// <param name="holidayArg">The list of FunctionArguments to be turned into DateTime objects.</param>
		public AdditionalHolidayDays(FunctionArgument holidayArg)
		{
			_holidayArg = holidayArg;
			Initialize();
		}

		public IEnumerable<System.DateTime> AdditionalDates => HolidayDates;

		/// <summary>
		/// Initalize takes the list of FunctionArguments and turns it into a list of DateTime objects. It handles the cases 
		/// where the arguments are dates as OADates, results of the DATE Function, integers, and strings.
		/// It works for both arrays of dates or dates as values from an Excel worksheet.
		/// </summary>
		private void Initialize()
		{
			var holidays = _holidayArg.Value as IEnumerable<FunctionArgument>;

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

			var range = _holidayArg.Value as ExcelDataProvider.IRangeInfo;
			if (range != null)
			{
				foreach (var cell in range)
				{
					if (ConvertUtil.TryParseDateObject(cell.Value, out date, out error))
						this.HolidayDates.Add(date);
				}
			}

			if (ConvertUtil.TryParseDateObject(_holidayArg.Value, out date, out error))
			{
				HolidayDates.Add(date);
			}
		}
	}
}