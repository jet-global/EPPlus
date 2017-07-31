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
* * Mats Alm   		                Added		                 2014-01-06
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
	/// This class contains the formula for computing the ROUNDDOWN Excel function.
	/// </summary>
	public class Rounddown : ExcelFunction
	{
		/// <summary>
		/// Takes the user specified arguments and rounds the first argument down by the specifications of the 
		/// second argument given.
		/// </summary>
		/// <param name="arguments">The user specified arguments.</param>
		/// <param name="context">Not used, but needed to override the method.</param>
		/// <returns>The first argument rounded down by the specifications of the second argument.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentsAreValid(arguments, 2, out eErrorType argumentError) == false)
				return new CompileResult(argumentError);

			var numberCandidate = arguments.ElementAt(0).Value;
			var nDecimalsCandidate = arguments.ElementAt(1).Value;

			if (numberCandidate == null)
				return this.CreateResult(0d, DataType.Decimal);

			if (!ConvertUtil.TryParseObjectToDecimal(numberCandidate, out double number))
				return new CompileResult(eErrorType.Value);

			if (nDecimalsCandidate == null)
				return this.CreateResult(System.Math.Floor(number), DataType.Decimal);

			if (!ConvertUtil.TryParseObjectToDecimal(nDecimalsCandidate, out double nDecimalsDouble))
				return new CompileResult(eErrorType.Value);
			var nDecimals = (int)nDecimalsDouble;

			var nFactor = number < 0 ? -1 : 1;
			number *= nFactor;

			if (nDecimals > 15)
				nDecimals = 15;

			double result;
			if (nDecimals > 0)
				result = RoundDownDecimalNumber(number, nDecimals);
			else
			{
				result = (int)System.Math.Floor(number);
				result = result - (result % System.Math.Pow(10, (nDecimals * -1)));
			}
			return this.CreateResult(result * nFactor, DataType.Decimal);
		}

		/// <summary>
		/// Rounds down the given decimal number.
		/// </summary>
		/// <param name="number">The number to round down.</param>
		/// <param name="nDecimals">The number of decimals to round the number to.</param>
		/// <returns>The given number rounded down to the appropriate number of decimal places.</returns>
		private static double RoundDownDecimalNumber(double number, int nDecimals)
		{
			var integerPart = System.Math.Floor(number);
			var decimalPart = number - integerPart;
			decimalPart = System.Math.Pow(10d, nDecimals) * decimalPart;
			decimalPart = System.Math.Truncate(decimalPart) / System.Math.Pow(10d, nDecimals);
			var result = integerPart + decimalPart;
			return result;
		}
	}
}
