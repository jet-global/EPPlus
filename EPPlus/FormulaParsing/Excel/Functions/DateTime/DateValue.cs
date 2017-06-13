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
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
	/// <summary>
	/// Simple implementation of DateValue function, just using .NET built-in
	/// function System.DateTime.TryParse, based on current culture
	/// </summary>
	public class DateValue : ExcelFunction
	{
		/// <summary>
		/// Takes the user specified date and returns that date as a serial number.
		/// </summary>
		/// <param name="arguments">The user specified date to be converted.</param>
		/// <param name="context">Not used, but needed to override the method.</param>
		/// <returns>The date as a serial number.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentsAreValid(arguments, 1, out eErrorType argumentError) == false)
				return new CompileResult(argumentError);
			var dateString = ArgToString(arguments, 0);
			return Execute(dateString);
		}

		/// <summary>
		/// Takes the date as a string and converts it into a serial number, and returns and error if it can't do so. 
		/// </summary>
		/// <param name="dateString">The user specified date as a string.</param>
		/// <returns>The date as a serial number if it can be parsed, otherwise it returns an error value.</returns>
		internal CompileResult Execute(string dateString)
		{
			System.DateTime result;
			System.DateTime.TryParse(dateString, out result);
			return result != System.DateTime.MinValue ?
				 CreateResult(result.ToOADate(), DataType.Date) :
				 CreateResult(ExcelErrorValue.Create(eErrorType.Value), DataType.ExcelError);
		}
	}
}
