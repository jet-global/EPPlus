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

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	/// <summary>
	/// Calculates variance based on the entire population (inlcudes logical values and text in the population).
	/// </summary>
	public class Varpa : HiddenValuesHandlingFunction
	{
		/// <summary>
		/// Variance measures how far a data set is spread out.
		/// Arguments can be the following: numbers; names, arrays, or references that contain numbers; text representations of numbers; or logical values, such as TRUE and FALSE, in a reference.
		/// Logical values and text representations of numbers that you type directly into the list of arguments are counted.
		/// </summary>
		/// <param name="arguments">Up too 254 individual arguments.</param>
		/// <param name="context">Unused, this is information about where the function is being executed.</param>
		/// <returns>The variance based on an entire population.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			//NOTE: This follows the Functionality of excel which is diffrent from the excel documentation.
			//If you pass in a null Stdev.S(1,1,1,,) it will treat those emtpy spaces as zeros insted of ignoring them.
			List<double> listToDoVarianceOn = new List<double>();
			var args = ArgsToDoubleEnumerable(this.IgnoreHiddenValues, false, arguments, context);
			foreach (var item in arguments)
			{ 
				if (item.IsExcelRange)
				{
					foreach (var cell in item.ValueAsRangeInfo)
					{
						if (cell.Value is string)
							listToDoVarianceOn.Add(0.0);
					}
					if (item.ValueFirst is double || item.ValueFirst is int || item.ValueFirst == null)
						continue;
				}
				if (item.ValueFirst == null)
					listToDoVarianceOn.Add(0.0);
				if (item.Value is ExcelDataProvider.IRangeInfo itemRange)
				{
					if (item.ValueFirst is bool valueIsABool)
					{
						if (valueIsABool == true)
						{
							listToDoVarianceOn.Add(1);
							continue;
						}
						else
						{
							listToDoVarianceOn.Add(0);
							continue;
						}
					}
					return this.CreateResult(0d, DataType.Decimal);
				}
			}
			foreach (var item in args)
				listToDoVarianceOn.Add(item);
			if (!this.TryVarPopulation(listToDoVarianceOn, out double VarPopulation))
				return new CompileResult(eErrorType.Value);
			return new CompileResult(VarPopulation, DataType.Decimal);
		}

		private bool TryVarPopulation(List<double> listOfDoubles, out double VarPopulation)
		{
			double avg = listOfDoubles.Average();
			double d = listOfDoubles.Aggregate(0.0, (total, next) => total += System.Math.Pow(next - avg, 2));
			VarPopulation = (d / (listOfDoubles.Count()));
			if (VarPopulation == 0 && listOfDoubles.All(x => x == -1))
				return false;
			return true;

		}
	}
}
