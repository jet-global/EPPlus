/* Copyright (C) 2011  Jan Källman
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
using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
	/// <summary>
	/// Returns the appropriate day of the week (as an int) from the given date and return type.
	/// </summary>
	public class Weekday : ExcelFunction
	{
		/// <summary>
		/// Checks if input is valid, and returns the corresponding weekday value if so.
		/// </summary>
		/// <param name="arguments">The given arguments used to calculate the weekday.</param>
		/// <param name="context">Unused in method, but necessary to override method.</param>
		/// <returns>Returns the correct weekday number or an ExcelErrorValue, depending on if the input is valid.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ValidateArguments(arguments, 1) == false)
				return new CompileResult(eErrorType.Value);
			var serialNumberCandidate = arguments.ElementAt(0).Value;

			var returnType = 1;
			if (arguments.Count() > 1 && !this.TryParseObjectAsInt(arguments.ElementAt(1).Value, out returnType, out eErrorType? returnTypeError))
				return new CompileResult(returnTypeError.Value);

			if (ConvertUtil.TryParseDateObjectToOADate(serialNumberCandidate, out double serialNumber) &&
				serialNumber < 1 && serialNumber >= 0)
				serialNumberCandidate = 7;

			

			if (ConvertUtil.TryParseDateObject(serialNumberCandidate, out System.DateTime date, out eErrorType? error))
			{
				if (date.ToOADate() < 61)
					date = new System.DateTime(date.Year, date.Month, date.Day - 1);
				var result = this.CalculateDayOfWeek(date, returnType);
				return this.CreateResult(result, DataType.Integer);
			}
			else
				return new CompileResult(error.Value);
			//if (serialNumberCandidate is null)
			//	return new CompileResult(eErrorType.Num);
			//if (serialNumberCandidate is string)
			//{
			//	var isDateString = System.DateTime.TryParse(serialNumberCandidate.ToString(), out System.DateTime date);
			//	if (!isDateString)
			//		return new CompileResult(eErrorType.Value);
			//}
			//if (arguments.Count() > 1)
			//{
			//	var returnTypeCandidate = arguments.ElementAt(1).Value;
			//	if (returnTypeCandidate is null)
			//		return new CompileResult(eErrorType.Num);
			//	else if (returnTypeCandidate is string)
			//	{
			//		var isValidReturnType = Int32.TryParse(returnTypeCandidate.ToString(), out int result);
			//		if (!isValidReturnType)
			//			return new CompileResult(eErrorType.Value);
			//	}
			//}
			//var serialNumber = this.ArgToDecimal(arguments, 0);
			//if(serialNumber < 0)
			//	return new CompileResult(eErrorType.Num);
			//var returnType = arguments.Count() > 1 ? ArgToInt(arguments, 1) : 1;
			//try
			//{
			//	var result = this.CalculateDayOfWeek(System.DateTime.FromOADate(serialNumber), returnType);
			//	return CreateResult(result, DataType.Integer);
			//}
			//catch (ExcelErrorValueException e)
			//{
			//	return new CompileResult(ExcelErrorValue.Values.ToErrorType(e.ErrorValue.ToString()));
			//}
		}

		private static List<int> _oneBasedStartOnSunday = new List<int> { 1, 2, 3, 4, 5, 6, 7 };
		private static List<int> _oneBasedStartOnMonday = new List<int> { 7, 1, 2, 3, 4, 5, 6 };
		private static List<int> _zeroBasedStartOnMonday = new List<int> { 6, 0, 1, 2, 3, 4, 5 };
		private static List<int> _oneBasedStartOnTuesday = new List<int> { 6, 7, 1, 2, 3, 4, 5 };
		private static List<int> _oneBasedStartOnWednesday = new List<int> { 5, 6, 7, 1, 2, 3, 4 };
		private static List<int> _oneBasedStartOnThursday = new List<int> { 4, 5, 6, 7, 1, 2, 3 };
		private static List<int> _oneBasedStartOnFriday = new List<int> { 3, 4, 5, 6, 7, 1, 2 };
		private static List<int> _oneBasedStartOnSaturday = new List<int> { 2, 3, 4, 5, 6, 7, 1};

		/// <summary>
		/// Returns the appropriate day of the week (as an int) from the given date and return type.
		/// </summary>
		/// <param name="dateTime">The user specified date.</param>
		/// <param name="returnType">The specified return type.</param>
		/// <returns></returns>
		private int CalculateDayOfWeek(System.DateTime dateTime, int returnType)
		{
			var dayIx = (int)dateTime.DayOfWeek;
			switch (returnType)
			{
				case 1:
					return _oneBasedStartOnSunday[dayIx];
				case 2:
					return _oneBasedStartOnMonday[dayIx];
				case 3:
					return _zeroBasedStartOnMonday[dayIx];
				case 11:
					return _oneBasedStartOnMonday[dayIx];
				case 12:
					return _oneBasedStartOnTuesday[dayIx];
				case 13:
					return _oneBasedStartOnWednesday[dayIx];
				case 14:
					return _oneBasedStartOnThursday[dayIx];
				case 15:
					return _oneBasedStartOnFriday[dayIx];
				case 16:
					return _oneBasedStartOnSaturday[dayIx];
				case 17:
					return _oneBasedStartOnSunday[dayIx];
				default:
					throw new ExcelErrorValueException(eErrorType.Num);
			}
		}

		private bool TryParseObjectAsInt(object intCandidate, out int resultInt, out eErrorType? error)
		{
			resultInt = -1;
			error = null;

			if (intCandidate == null)
			{
				error = eErrorType.Num;
				return false;
			}
			else if (intCandidate is string intString && Int32.TryParse(intString, out resultInt))
			{
				return true;
			}
			else if (intCandidate is double intDouble)
			{
				resultInt = (int)intDouble;
				return true;
			}
			else if (intCandidate is int)
			{
				resultInt = (int)intCandidate;
				return true;
			}
			else
			{
				error = eErrorType.Value;
				return false;
			}
		}
	}
}
