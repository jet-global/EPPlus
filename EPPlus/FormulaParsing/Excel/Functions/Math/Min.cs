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
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	/// <summary>
	/// This class contains the formula for calculating the minimum value in a set of arguments. 
	/// </summary>
	public class Min : HiddenValuesHandlingFunction
	{
		/// <summary>
		/// Takes the user specified arguments and returns the minimum value.
		/// </summary>
		/// <param name="arguments">The user specified arguments, which can be a list, array, or cell referece.</param>
		/// <param name="context">The context in which the method is being called.</param>
		/// <returns>The minimum item in the user specified argument list.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentsAreValid(arguments, 1, out eErrorType errorValue) == false)
				return new CompileResult(errorValue);
			var args = arguments.ElementAt(0);
			var argumentValueList = this.ArgsToObjectEnumerable(false, new List<FunctionArgument> { args }, context);
			var values = argumentValueList.Where(arg => ((arg.GetType().IsPrimitive && (arg is bool == false)) || arg is System.DateTime));
			//If the input to the Min Function is not an excel range logical values and string representations of numbers
			//are allowed, even though they are not counted in a cell reference.
			if (!arguments.ElementAt(0).IsExcelRange)
			{
				if(arguments.Count() == 1)
				{
					if (values.Count() > 255)
						return new CompileResult(eErrorType.NA);
					return CreateResult(values.Min(), DataType.Decimal);
				}

				var doublesList = new List<double> { };
				foreach (var item in arguments)
				{
					if (item.ExcelStateFlagIsSet(ExcelCellState.HiddenCell))
						continue;
					if (item.Value is string)
					{
						if (ConvertUtil.TryParseNumericString(item.Value, out double result))
							doublesList.Add(result);
						else if (ConvertUtil.TryParseDateString(item.Value, out System.DateTime dateResult))
							doublesList.Add(dateResult.ToOADate());
					}
					else
						doublesList.Add(this.ArgToDecimal(item.Value));
				}
				if (doublesList.Count() == 0)
					return new CompileResult(eErrorType.Value);
				if (doublesList.Count() > 255)
					return new CompileResult(eErrorType.NA);
				return CreateResult(doublesList.Min(), DataType.Decimal);
			}

			if (values.Count() > 255)
				return new CompileResult(eErrorType.NA);
			return this.CreateResult(values.Min(), DataType.Decimal);
		}
	}
}
