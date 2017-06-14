/* Copyright (C) 2011  Jan Källman
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
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	public class Floor : ExcelFunction
	{
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentsAreValid(arguments, 2, out eErrorType argumentError) == false)
				return new CompileResult(argumentError);
			var number = ArgToDecimal(arguments, 0);
			var significance = ArgToDecimal(arguments, 1);
			if ((number > 0d && significance < 0))
				return new CompileResult(eErrorType.Num);
			if (significance == 0.0)
				return new CompileResult(eErrorType.Div0);

			if (number == 0.0)
				return base.CreateResult(0.0, DataType.Decimal);
			double divisionResult = number / significance;
			int multiple = (int)divisionResult;
			bool exactChange = divisionResult == multiple;
			if (exactChange)
				return base.CreateResult(number, DataType.Decimal);
			else if (significance > 0 && number < 0)
				return base.CreateResult((multiple - 1) * significance, DataType.Decimal);
			else
				return base.CreateResult(multiple * significance, DataType.Decimal);

		}
	}
}
