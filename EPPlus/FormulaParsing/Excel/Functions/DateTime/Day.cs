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
using OfficeOpenXml.Utils;
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
	/// <summary>
	/// Returns the corresponding day of the month (as an int) from the given date.
	/// </summary>
	public class Day : DateParsingFunction
	{
		/// <summary>
		/// Checks if the input is valid, and returns the corresponding day of the month if so.
		/// </summary>
		/// <param name="arguments">The given arguments used to calculate the day of the month.</param>
		/// <param name="context">Unused in the method, but necessary to override the method.</param>
		/// <returns></returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ValidateArguments(arguments, 1) == false)
				return new CompileResult(eErrorType.Value);			
			var serialNumberCandidate = this.GetFirstValue(arguments);
			var isValidDate = ConvertUtil.TryParseDateObject(serialNumberCandidate, out System.DateTime date, out eErrorType? error);
			if (isValidDate)
				return this.CreateResult(date.Day, DataType.Integer);
			if (serialNumberCandidate is int serialNumberInt)
				serialNumberCandidate = serialNumberInt * 1.0;
			if (serialNumberCandidate is double serialNumberDouble && serialNumberDouble == 0)
				return this.CreateResult(0, DataType.Integer);
			return new CompileResult((eErrorType)error);
		}
	}
}
