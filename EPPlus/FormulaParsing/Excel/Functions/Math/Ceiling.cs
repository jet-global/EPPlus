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
	/// This class contains the formula for the CEILING Function in Excel
	/// </summary>
	public class Ceiling : ExcelFunction
	{
		/// <summary>
		/// Takes the user specified arguments and rounds the first argument up to the nearest multiple of the 
		/// second argument.
		/// </summary>
		/// <param name="arguments">The user specified arguments.</param>
		/// <param name="context">Not used, but needed to override the method.</param>
		/// <returns>The first argument rounded up to the nearest multiple of the second argument.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentsAreValid(arguments, 2, out eErrorType argumentError) == false)
				return new CompileResult(argumentError);

			var numberCandidate = arguments.ElementAt(0).Value;
			var significanceCandidate = arguments.ElementAt(1).Value;
		
			if (numberCandidate == null || significanceCandidate == null)
				return this.CreateResult(0d, DataType.Decimal);

			if (!ConvertUtil.TryParseObjectToDecimal(numberCandidate, out double number))
				return new CompileResult(eErrorType.Value);
			if (!ConvertUtil.TryParseObjectToDecimal(significanceCandidate, out double significance))
				return new CompileResult(eErrorType.Value);

			if (number > 0 && significance < 0)
				return new CompileResult(eErrorType.Num);
			else if (significance < 1 && significance > 0)
			{
				var floor = System.Math.Floor(number);
				var rest = number - floor;
				var nSign = (int)(rest / significance) + 1;
				return this.CreateResult(floor + (nSign * significance), DataType.Decimal);
			}
			else if (significance == 1)
				return this.CreateResult(System.Math.Ceiling(number), DataType.Decimal);
			else if (significance == 0 || number == 0)
				return this.CreateResult(0d, DataType.Decimal);
			else if (number % significance == 0)
				return this.CreateResult(number, DataType.Decimal);
			else if (number < 0 && significance > 0)
			{
				var modNum = -1 * (number % significance);
				var result = number + modNum;
				return this.CreateResult(result, DataType.Decimal);
			}
			else
			{
				var result = number - (number % significance) + significance;
				return this.CreateResult(result, DataType.Decimal);
			}
		}
	}
}
