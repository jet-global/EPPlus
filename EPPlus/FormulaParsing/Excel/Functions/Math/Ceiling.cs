﻿/* Copyright (C) 2011  Jan Källman
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
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	/// <summary>
	/// 
	/// </summary>
	public class Ceiling : ExcelFunction
	{
		/// <summary>
		/// 
		/// </summary>
		/// <param name="arguments"></param>
		/// <param name="context"></param>
		/// <returns></returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentsAreValid(arguments, 2, out eErrorType argumentError) == false)
				return new CompileResult(argumentError);

			var numberCandidate = arguments.ElementAt(0).Value;
			var significanceCandidate = arguments.ElementAt(1).Value;

			if (numberCandidate == null || significanceCandidate == null)
				return CreateResult(0d, DataType.Decimal);

			if (!ConvertUtil.TryParseDateObjectToOADate(numberCandidate, out double number))
				return new CompileResult(eErrorType.Value);

			if (!ConvertUtil.TryParseDateObjectToOADate(significanceCandidate, out double significance))
				return new CompileResult(eErrorType.Value);


			if (number > 0 && significance < 0)
				return new CompileResult(eErrorType.Num);
			
			if (significance < 1 && significance > 0)
			{
				var floor = System.Math.Floor(number);
				var rest = number - floor;
				var nSign = (int)(rest / significance) + 1;
				return CreateResult(floor + (nSign * significance), DataType.Decimal);
			}
			else if (significance == 1)
			{
				return CreateResult(System.Math.Ceiling(number), DataType.Decimal);
			}
			else
			{
				var result = number - (number % significance) + significance;
				return CreateResult(result, DataType.Decimal);
			}
		}
	}
}
