﻿/*******************************************************************************
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
	/// This class contains the formula for the ROUNDUP Function in Excel
	/// </summary>
	public class Roundup : ExcelFunction
	{
		/// <summary>
		/// Takes the first user argument and rounds it up based on the second user specified argument.
		/// </summary>
		/// <param name="arguments">The user specified arguments.</param>
		/// <param name="context">Not used, but needed to override the method.</param>
		/// <returns>The first argument rounded up with respect to the specifications of the second argument.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentsAreValid(arguments, 2, out eErrorType argumentError) == false)
				return new CompileResult(argumentError);

			var numberCandidate = arguments.ElementAt(0).Value;
			var nDigitsCandidate = arguments.ElementAt(1).Value;

			if (numberCandidate == null)
				return CreateResult(0d, DataType.Decimal);

			if (!ConvertUtil.TryParseDateObjectToOADate(numberCandidate, out double numberDecimal))
				return new CompileResult(eErrorType.Value);
			var number = numberDecimal;

			if (nDigitsCandidate == null)
				return this.CreateResult(System.Math.Ceiling(number), DataType.Decimal);

			if (!ConvertUtil.TryParseDateObjectToOADate(nDigitsCandidate, out double digitsDecimal))
				return new CompileResult(eErrorType.Value);
			var nDigits = (int)digitsDecimal;

			if (nDigits > 15)
				nDigits = 15;
			double result = (number >= 0)
				 ? System.Math.Ceiling(number * System.Math.Pow(10, nDigits)) / System.Math.Pow(10, nDigits)
				 : System.Math.Floor(number * System.Math.Pow(10, nDigits)) / System.Math.Pow(10, nDigits);
			return this.CreateResult(result, DataType.Decimal);
		}
	}
}
