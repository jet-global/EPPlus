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
*  * Author							Change						Date
* *******************************************************************************
* * Mats Alm   		                Added		                2013-12-03
* *******************************************************************************
* For code change notes, see the source control history.
*******************************************************************************/
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	/// <summary>
	/// This class contains the formula for the ROUND Function in Excel.
	/// </summary>
	public class Round : ExcelFunction
	{
		/// <summary>
		/// Takes the user specified arguments and rounds the first argument to the number of decimal places
		/// designated by the second argument.
		/// </summary>
		/// <param name="arguments">The user specified arguments.</param>
		/// <param name="context">Not used, but needed to override the method.</param>
		/// <returns>The first arugment rounded to the specifications of the second argument.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentsAreValid(arguments, 2, out eErrorType argumentError) == false)
				return new CompileResult(argumentError);

			if (arguments.ElementAt(0).Value == null)
				return this.CreateResult(0d, DataType.Decimal);

			if (!ConvertUtil.TryParseObjectToDecimal(arguments.ElementAt(0).Value, out double number))
				return new CompileResult(eErrorType.Value);

			if (arguments.ElementAt(1).Value == null)
				return this.CreateResult(System.Math.Round(number, 0), DataType.Decimal);

			if (!ConvertUtil.TryParseObjectToDecimal(arguments.ElementAt(1).Value, out double nDigitsDecimal))
				return new CompileResult(eErrorType.Value);
			var nDigits = (int)nDigitsDecimal;

			if (nDigits > 15)
				nDigits = 15;

			if (nDigits < 0)
			{
				nDigits *= -1;
				var roundedNumber = System.Math.Round(number / System.Math.Pow(10, nDigits));
				var result = (System.Math.Pow(10, nDigits))*(roundedNumber);
				return this.CreateResult(result, DataType.Integer);
			}
			return this.CreateResult(System.Math.Round(number, nDigits), DataType.Decimal);
		}
	}
}
