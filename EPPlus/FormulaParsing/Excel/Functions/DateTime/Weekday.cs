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
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
	/// <summary>
	/// Returns the appropriate day of the week (as an int) from the given date and return type if given valid input.
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
			if (ValidateArguments(arguments, 1) == false)
				return new CompileResult(eErrorType.Value);

			var element1 = arguments.ElementAt(0).Value;

			if (element1 is null)
				return new CompileResult(eErrorType.Num);

			if (element1 is string)
			{
				foreach(var a in element1.ToString().ToCharArray())
				{
					if ((int)a < 45 || (int)a > 57)
						return new CompileResult(eErrorType.Value);
					if((int)a == 46)
						return new CompileResult(eErrorType.Value);
				}
			}
				

			if (arguments.Count() > 1 && arguments.ElementAt(1).Value is null)
				return new CompileResult(eErrorType.Num);

			if (arguments.Count() >1 && arguments.ElementAt(1).Value is string)
				return new CompileResult(eErrorType.Value);

			var serialNumber = ArgToDecimal(arguments, 0);

			if(serialNumber < 0)
				return new CompileResult(eErrorType.Num);

			var returnType = arguments.Count() > 1 ? ArgToInt(arguments, 1) : 1;

			try
			{
				CalculateDayOfWeek(System.DateTime.FromOADate(serialNumber), returnType);
				return CreateResult(CalculateDayOfWeek(System.DateTime.FromOADate(serialNumber), returnType), DataType.Integer);
			}
			catch (ExcelErrorValueException e)
			{
				return new CompileResult(ExcelErrorValue.Values.ToErrorType(e.ErrorValue.ToString()));
			}
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
	}
}
