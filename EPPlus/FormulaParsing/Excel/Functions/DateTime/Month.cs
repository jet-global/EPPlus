﻿/* Copyright (C) 2011  Jan Källman
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
 * Code change notes:
 * 
 * Author							Change						Date
 *******************************************************************************
 * Mats Alm   		                Added		                2013-12-03
 *******************************************************************************/
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
	/// <summary>
	/// Returns the corresponding month of the year (as an int) from the given date.
	/// </summary>
	public class Month : ExcelFunction
	{
		/// <summary>
		/// Checks if the input is a valid date, and returns the corresponding month of the year if so.
		/// </summary>
		/// <param name="arguments">The given arguments used to calculate the month of the year.</param>
		/// <param name="context">Unused in the method, but necessary to override the method.</param>
		/// <returns>Returns the numeric month of the year for the given date, or an ExcelErrorValue, depending on if the input is valid.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ValidateArguments(arguments, 1) == false)
				return new CompileResult(eErrorType.Value);
			var dateObj = arguments.ElementAt(0).Value;
			// 0 is a special case and requires specific output.
			if ((dateObj is int dateInt && dateInt == 0) || (dateObj is double dateDouble && dateDouble == 0))
				return this.CreateResult(1, DataType.Integer);
			if (ConvertUtil.TryParseDateObject(dateObj, out System.DateTime date, out eErrorType? error))
				return this.CreateResult(date.Month, DataType.Integer);
			else
				return new CompileResult((eErrorType)error);
		}
	}
}
