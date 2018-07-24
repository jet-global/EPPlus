/*******************************************************************************
* You may amend and distribute as you like, but don't remove this header!
*
* EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
* See http://www.codeplex.com/EPPlus for details.
*
* Copyright (C) 2011-2017 Jan Källman, Matt Delaney, and others as noted in the source history.
*
* This library is free software; you can redistribute it and/or
* modify it under the terms of the GNU Lesser General Public
* License as published by the Free Software Foundation; either
* version 2.1 of the License, or (at your option) any later version.

* This library is distributed in the hope that it will be useful,
* but WITHOUT ANY WARRANTY; without even the implied warranty of
* MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
* See the GNU Lesser General Public License for more details.
*
* The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
* If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
*
* All code and executables are provided "as is" with no warranty either express or implied. 
* The author accepts no liability for any damage or loss of business that this product may cause.
*
* For code change notes, see the source control history.
*******************************************************************************/
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
		protected override int HolidayIndex { get; } = 2;
		#endregion

		#region ExcelFunction Overrides
		/// <summary>
		/// Returns a calculator with default Saturday and Sunday as weekend.
		/// </summary>
		/// <param name="weekend">The user specified weekend code</param>
		/// <returns>A calculator with Saturday and Sunday as default weekend days.</returns>
		protected override WorkdayCalculator GetCalculator(object weekend)
		{
			return new WorkdayCalculator();
		}

		/// <summary>
		/// Returns whether or not there is a weekend parameter.
		/// </summary>
		/// <returns>False since Workday does not have a weekend parameter.</returns>
		protected override bool WeekendSpecified(FunctionArgument[] functionArguments)
		{
			return false;
		}

		/// <summary>
		/// Returns whether holidays parameter is specified by user.
		/// </summary>
		/// <param name="functionArguments">The array of parameters for function</param>
		/// <returns>A boolean depending on whether or not the holiday parameter is given.</returns>
		protected override bool HolidaysSpecified(FunctionArgument[] functionArguments)
		{
			return functionArguments.Length > 2;
		}
		#endregion
	}
}
