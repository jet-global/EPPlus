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
using MathObj = System.Math;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	/// <summary>
	/// Calculates standard deviation based on the entire population given as arguments (ignores logical values and text).
	/// </summary>
	public class StdevP : HiddenValuesHandlingFunction
	{
		/// <summary>
		/// The standard deviation is a measure of how widely values are dispersed from the average value (the mean).
		/// Logical values and text representations of numbers that you type directly into the list of arguments are counted.
		/// If an argument is an array or reference, only numbers in that array or reference are counted.Empty cells, logical values, text, or error values in the array or reference are ignored.
		/// </summary>
		/// <param name="arguments">Up to 254 individual arguments.</param>
		/// <param name="context">Unused, this is information about where the function is being executed.</param>
		/// <returns>The standard deviation based on the entire population.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			//NOTE: This follows the Functionality of excel which is diffrent from the excel documentation.
			//If you pass in a null Stdev.P(1,1,1,,) it will treat those emtpy spaces as zeros insted of ignoring them.
			List<double> listToDoStandardDeviationOn = new List<double>();
			var args = ArgsToDoubleEnumerable(this.IgnoreHiddenValues, false, arguments, context);
			foreach (var item in arguments)
			{
				if (item.IsExcelRange)
				{
					if (item.ValueFirst is double || item.ValueFirst is int || item.ValueFirst == null)
						continue;
					return new CompileResult(eErrorType.Div0);
				}
				if (item.ValueFirst == null)
				{
					listToDoStandardDeviationOn.Add(0.0);
				}
			}
			foreach (var item in args)
			{
				listToDoStandardDeviationOn.Add(item);
			}
			if (!this.TryStandardDeviationEntirePopulation(listToDoStandardDeviationOn, out double standardDeviation))
				return new CompileResult(eErrorType.Value);
			return this.CreateResult(standardDeviation, DataType.Decimal);
		}

		private bool TryStandardDeviationEntirePopulation(IEnumerable<double> values, out double standardDeviation)
		{
			standardDeviation = MathObj.Sqrt(values.Average(v => MathObj.Pow(v - values.Average(), 2)));
			if (standardDeviation == 0 && values.All(x => x == values.First()))
				return false;
			return true;
		}
	}
}
