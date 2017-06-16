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
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;
using MathObj = System.Math;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	public class Cot : ExcelFunction
	{
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (ArgumentCountIsValid(arguments, 1) == false)
				return new CompileResult(eErrorType.Value);
			var argument = arguments.First().Value;
			if (argument is string & !ConvertUtil.TryParseDateObjectToOADate(argument, out double result))
			{
				return new CompileResult(eErrorType.Value);
			}
			if (Cosecant(result) == -2)
				return new CompileResult(eErrorType.Div0);
			return CreateResult(Cotangent(result), DataType.Decimal);
		}

		private static double Cotangent(double x)
		{
			if (MathObj.Tan(x) == 0)
				return -2;
			return 1 / MathObj.Tan(x);
		}

		private static double Cosecant(double x)
		{
			if (MathObj.Sin(x) == 0)
				return -2;
			return 1 / MathObj.Sin(x);
		}
	}
}