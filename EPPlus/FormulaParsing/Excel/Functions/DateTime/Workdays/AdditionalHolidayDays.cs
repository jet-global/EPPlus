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
		private readonly List<System.DateTime> _holidayDates = new List<System.DateTime>();
		/// <summary>
		/// The constructor for this class calls the method Initalize, which does the transformation of the list. 
		/// </summary>
		/// <param name="holidayArg">The list of FunctionArguments to be turned into DateTime objects.</param>
		public AdditionalHolidayDays(FunctionArgument holidayArg)
		{
			_holidayArg = holidayArg;
			Initialize();
		}

		public IEnumerable<System.DateTime> AdditionalDates => _holidayDates;

		/// <summary>
		/// Initalize takes the list of FunctionArguments and turns it into a list of DateTime objects. It handles the cases 
		/// where the arguments are dates as OADates, results of the DATE Function, integers, and strings.
		/// It works for both arrays of dates or dates as values from an Excel worksheet.
		/// </summary>
		private void Initialize()
		{
			bool testBool = false;
			var holidays = _holidayArg.Value as IEnumerable<FunctionArgument>;

			if (holidays != null)
			{
				foreach (var holidayDate in from arg in holidays where ConvertUtil.IsNumeric(arg.Value) select ConvertUtil.GetValueDouble(arg.Value) into dateSerial select System.DateTime.FromOADate(dateSerial))
				{
					_holidayDates.Add(holidayDate);
					testBool = true;
				}

				if (!testBool)
				{
					foreach (var arg in holidays)
					{
						if (!ConvertUtil.TryParseDateString(arg.Value, out System.DateTime date))
							continue;
						_holidayDates.Add(date);
					}
				}
			}

			var range = _holidayArg.Value as ExcelDataProvider.IRangeInfo;
			if (range != null)
			{
				foreach (var holidayDate in from cell in range where ConvertUtil.IsNumeric(cell.Value) select ConvertUtil.GetValueDouble(cell.Value) into dateSerial select System.DateTime.FromOADate(dateSerial))
				{
					_holidayDates.Add(holidayDate);
					testBool = true;
				}

				if(!testBool)
				{
					foreach (var arg in range)
					{
						if (!ConvertUtil.TryParseDateString(arg.Value, out System.DateTime date))
							continue;
						_holidayDates.Add(date);
					}
				}
			}

			if (ConvertUtil.IsNumeric(_holidayArg.Value))
			{
				_holidayDates.Add(System.DateTime.FromOADate(ConvertUtil.GetValueDouble(_holidayArg.Value)));
			}
		}
	}
}