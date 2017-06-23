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
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
		/// 
		/// </summary>
		/// <param name="arguments"></param>
		/// <param name="context"></param>
		/// <returns></returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			double number = 0;
			double significance = 0;
			double mode = 0;
			if (this.ArgumentsAreValid(arguments, 1, out eErrorType argumentError) == false)
				return new CompileResult(argumentError);

			var numberCandidate = arguments.ElementAt(0).Value;
			var significanceCandidate = arguments.ElementAt(1).Value;

			if (numberCandidate == null)
				return CreateResult(0d, DataType.Decimal);

			if (!ConvertUtil.TryParseDateObjectToOADate(numberCandidate, out double re))
				return new CompileResult(eErrorType.Value);
			number = re;

			if (significanceCandidate == null)
				return CreateResult(System.Math.Floor(number), DataType.Decimal);

			if (!ConvertUtil.TryParseDateObjectToOADate(significanceCandidate, out double ree))
				return new CompileResult(eErrorType.Value);

			significance = ree;

			if(arguments.Count() > 2)
			{
				var modeCandidate = arguments.ElementAt(2).Value;

				if(modeCandidate == null)
				{
					double divisionResult1 = number / significance;
					int multiple1 = (int)divisionResult1;
					bool exactChange1 = divisionResult1 == multiple1;
					if (exactChange1)
						return this.CreateResult(number, DataType.Decimal);
					else if (significance > 0 && number < 0)
						return this.CreateResult((multiple1 - 1) * significance, DataType.Decimal);
					else
						return this.CreateResult(multiple1 * significance, DataType.Decimal);
					}

				if (!ConvertUtil.TryParseDateObjectToOADate(modeCandidate, out double reee))
					return new CompileResult(eErrorType.Value);
				mode = reee;

				if (mode > 0 || mode < 0)
				{
					var divisionResult3 = number / significance;
					var multiple3 = (int)divisionResult3;
					var exactChange3 = divisionResult3 == multiple3;
					if (exactChange3)
						return this.CreateResult(number, DataType.Decimal);


					else
						return this.CreateResult(multiple3 * significance, DataType.Decimal);
				}
				else
				{
					var divisionResult2 = number / significance;
					var multiple2 = (int)divisionResult2;
					var exactChange2 = divisionResult2 == multiple2;
					if (exactChange2)
						return this.CreateResult(number, DataType.Decimal);
					else if (significance > 0 && number < 0)
						return this.CreateResult((multiple2 - 1) * significance, DataType.Decimal);
					else if (number < 0)
						return this.CreateResult((multiple2 + 1) * significance, DataType.Decimal);
					else
						return this.CreateResult(multiple2 * significance, DataType.Decimal);
				}


			}
			

			var divisionResult = number / significance;
			var multiple = (int)divisionResult;
			var exactChange = divisionResult == multiple;
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
