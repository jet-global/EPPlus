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
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	/// <summary>
	/// This class contains the formula for the CEILING.MATH Excel Function. 
	/// </summary>
	public class CeilingMath : ExcelFunction
	{
		/// <summary>
		/// Takes the first user argument and rounds it up based on the optional second and third arguments.
		/// </summary>
		/// <param name="arguments">The first argument is the number to be rounded up, the second is the significance value to round the
		/// number to, and the third argument is the mode that determines if the number is rounded closer to zero or not (for negative 
		/// numbers only).</param>
		/// <param name="context">Not used, but needed to override the method.</param>
		/// <returns>The first argument rounded up based on the specifications of the second and third 
		/// optional user arguments.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			double divisionResult;
			int multiple;
			bool exactChange;

			if (this.ArgumentsAreValid(arguments, 1, out eErrorType argumentError) == false)
				return new CompileResult(argumentError);
			var numberCandidate = arguments.ElementAt(0).Value;
			var significanceCandidate = arguments.ElementAt(1).Value;
			if (numberCandidate == null)
				return this.CreateResult(0d, DataType.Decimal);
			if (!ConvertUtil.TryParseObjectToDecimal(numberCandidate, out double number))
				return new CompileResult(eErrorType.Value);
			if (significanceCandidate == null)
				return this.CreateResult(System.Math.Ceiling(number), DataType.Decimal);
			if (!ConvertUtil.TryParseObjectToDecimal(significanceCandidate, out double significance))
				return new CompileResult(eErrorType.Value);
			if (significance < 0)
				significance = -1 * significance;
			if (arguments.Count() > 2)
			{
				var modeCandidate = arguments.ElementAt(2).Value ?? 0;
				if (!ConvertUtil.TryParseObjectToDecimal(modeCandidate, out double mode))
					return new CompileResult(eErrorType.Value);
				if (mode == 0 || number > 0)
				{
					if (significance < 1 && significance > 0)
					{
						var floor = System.Math.Floor(number);
						var rest = number - floor;
						var nSign = (int)(rest / significance) + 1;
						return this.CreateResult(floor + (nSign * significance), DataType.Decimal);
					}
					else if (significance == 1)
						return this.CreateResult(System.Math.Ceiling(number), DataType.Decimal);
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
				divisionResult = number / significance;
				multiple = (int)divisionResult;
				exactChange = divisionResult == multiple;
				if (exactChange)
					return this.CreateResult(number, DataType.Decimal);
				else if (significance > 0 && number < 0)
					return this.CreateResult((multiple - 1) * significance, DataType.Decimal);
				else if (number < 0)
					return this.CreateResult((multiple + 1) * significance, DataType.Decimal);
				else
					return this.CreateResult(multiple * significance, DataType.Decimal);
			}
			else if (significance == 1)
				return this.CreateResult(System.Math.Ceiling(number), DataType.Decimal);
			else if (significance == 0 || number == 0)
				return this.CreateResult(0d, DataType.Decimal);
			else if (number % significance == 0)
				return this.CreateResult(number, DataType.Decimal);
			else
			{
				var result = number - (number % significance) + significance;
				return this.CreateResult(result, DataType.Decimal);
			}
		}
	}
}
