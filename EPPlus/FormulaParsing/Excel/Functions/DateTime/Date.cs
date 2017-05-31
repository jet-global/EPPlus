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
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
	public class Date : ExcelFunction
	{
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (ValidateArguments(arguments, 3) == false)
				return new CompileResult(eErrorType.Value);
			
			var yearObj = arguments.ElementAt(0).Value;
			var monthObj = arguments.ElementAt(1).Value;
			var dayObj = arguments.ElementAt(2).Value;

			var isValidYear = this.TryParseNumber(yearObj, out int year);
			var isValidMonth = this.TryParseNumber(monthObj, out int month);
			var isValidDay = this.TryParseNumber(dayObj, out int day);
			if (!isValidYear)
				return new CompileResult(this.getErrorValue(yearObj));
			else if (!isValidMonth)
				return new CompileResult(this.getErrorValue(monthObj));
			else if (!isValidDay)
				return new CompileResult(this.getErrorValue(dayObj));

			if (year >= 10000)
				return new CompileResult(eErrorType.Num);
			else if (year < 1900)
				year += 1900;

			var date = new System.DateTime(year, 1, 1);
			var startDate = date;
			month -= 1;
			day -= 1;
			
			try
			{
				date = date.AddMonths(month);
				startDate = date;
				date = date.AddDays(day);
			}
			catch (ArgumentOutOfRangeException e)
			{
				return new CompileResult(eErrorType.Num);
			}

			var resultOADate = date.ToOADate();
			if (resultOADate < 1)
				return new CompileResult(eErrorType.Num);

			
			if (startDate.ToOADate() < 61 && resultOADate != 61)
				resultOADate--;
			return CreateResult(resultOADate, DataType.Date);
		}

		private bool TryParseNumber(object numberCandidate, out int resultNumber)
		{
			resultNumber = -1;
			if (numberCandidate == null)
			{
				resultNumber = 0;
				return true;
			}
			else if (numberCandidate is string numberString && Double.TryParse(numberString, out double parsedNumber))
			{
				resultNumber = (int)parsedNumber;
				return true;
			}
			else if (numberCandidate is double numberDouble)
			{
				resultNumber = (int)numberDouble;
				return true;
			}
			else if (numberCandidate is int numberInt)
			{
				resultNumber = numberInt;
				return true;
			}
			else
				return false;
		}

		private bool ArgumentIsNumeric(object argument)
		{
			if (argument is string argString)
			{
				return (Double.TryParse(argString, out double result));
			}
			else
				return (argument is double || argument is int);
		}

		private eErrorType getErrorValue(object invalidObject)
		{
			if (invalidObject is string)
				return eErrorType.Value;
			else
				return eErrorType.Num;
		}
	}
}
