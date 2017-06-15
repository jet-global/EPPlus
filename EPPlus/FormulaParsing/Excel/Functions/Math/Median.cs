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
* * Mats Alm   		                Added		                2015-01-10
* *******************************************************************************
* For code change notes, see the source control history.
*******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	/// <summary>
	/// This class contains the fomula for calculating the median of a set of data. 
	/// </summary>
	public class Median : ExcelFunction
	{
		/// <summary>
		/// Takes the user specified arguments and returns the median of the data. 
		/// </summary>
		/// <param name="arguments">The user specified list, array, or cell reference.</param>
		/// <param name="context">The context in which the method is being called.</param>
		/// <returns></returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			double[] numberArray;
			if (this.ArgumentCountIsValid(arguments, 1) == false)
				return new CompileResult(eErrorType.Value);
			if (arguments.ElementAt(0).Value == null && arguments.Count() == 1)
				return new CompileResult(eErrorType.Num);
			var argumentValueList = this.ArgsToObjectEnumerable(false, new List<FunctionArgument> { arguments.ElementAt(0) }, context);
			foreach (var item in argumentValueList)
			{
				if (item is ExcelErrorValue)
					return new CompileResult((ExcelErrorValue)item);
			}

			if (!arguments.ElementAt(0).IsExcelRange)
			{
				var doubleList = new List<double> { };
				foreach (var item in arguments)
				{
					if (item.ExcelStateFlagIsSet(ExcelCellState.HiddenCell))
						continue;
					if (item.Value is string)
					{
						if (ConvertUtil.TryParseNumericString(item.Value, out double decimalResult))
							doubleList.Add(decimalResult);
						else if (ConvertUtil.TryParseDateString(item.Value, out System.DateTime dateResult))
							doubleList.Add(dateResult.ToOADate());
						else
							return new CompileResult(eErrorType.Value);
					}
					else if(item.Type == null)
						doubleList.Add(0.0);
					else
						doubleList.Add(this.ArgToDecimal(item.Value));
				}
				numberArray = doubleList.ToArray();
			}
			else
				numberArray = ArgsToDoubleEnumerable(arguments, context).ToArray();
		
			if (numberArray.Length == 0)
				return new CompileResult(eErrorType.Num);
			if (numberArray.Length > 255)
				return new CompileResult(eErrorType.NA);

			return this.CreateResult(this.getMedian(numberArray), DataType.Decimal);
		}

		/// <summary>
		/// Returns the median of an array of numbers.
		/// </summary>
		/// <param name="array">The user specified array of doubles.</param>
		/// <returns>The number which is the median.</returns>
		private double getMedian(double[] array)
		{
			Array.Sort(array);

			double result;
			if (array.Length % 2 == 1)
			{
				result = array[array.Length / 2];
			}
			else
			{
				var startIndex = array.Length / 2 - 1;
				result = (array[startIndex] + array[startIndex + 1]) / 2d;
			}
			return result;
		}
	}
}
