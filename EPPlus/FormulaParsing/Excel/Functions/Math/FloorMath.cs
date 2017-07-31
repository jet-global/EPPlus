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
	/// This class contains the formula for the FLOOR.MATH Excel function. 
	/// </summary>
	public class FloorMath : ExcelFunction
	{
		/// <summary>
		/// Takes the user specified argument and rounds the first argument down with respect to the other two 
		/// optional arguments.
		/// </summary>
		/// <param name="arguments">The user specified arguments.</param>
		/// <param name="context">Not used, but needed to override the method.</param>
		/// <returns>The first argument rounded down by the specifications of the other two optional arguments.</returns>
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
				return this.CreateResult(System.Math.Floor(number), DataType.Decimal);
			if (!ConvertUtil.TryParseObjectToDecimal(significanceCandidate, out double significance))
				return new CompileResult(eErrorType.Value);
			if(arguments.Count() > 2)
			{
				var modeCandidate = arguments.ElementAt(2).Value;
				if(modeCandidate == null)
				{
					divisionResult = number / significance;
					multiple = (int)divisionResult;
					exactChange = divisionResult == multiple;
					if (exactChange)
						return this.CreateResult(number, DataType.Decimal);
					else if (significance > 0 && number < 0)
						return this.CreateResult((multiple - 1) * significance, DataType.Decimal);
					else
						return this.CreateResult(multiple * significance, DataType.Decimal);
				}
				if (!ConvertUtil.TryParseObjectToDecimal(modeCandidate, out double mode))
					return new CompileResult(eErrorType.Value);
				if (mode != 0)
				{
					divisionResult = number / significance;
					multiple = (int)divisionResult;
					exactChange = divisionResult == multiple;
					if (exactChange)
						return this.CreateResult(number, DataType.Decimal);
					else
						return this.CreateResult(multiple * significance, DataType.Decimal);
				}
				else
				{
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
	}
}
