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
	/// Estimates standard deviation based on a sample.
	/// </summary>
	public class StdevS : HiddenValuesHandlingFunction
	{
		/// <summary>
		/// The standard deviation is a measure of how widely values are dispersed from the average value (the mean).
		/// Logical values and text representations of numbers that you type directly into the list of arguments are counted.
		/// If an argument is an array or reference, only numbers in that array or reference are counted.Empty cells, logical values, text, or error values in the array or reference are ignored.
		/// </summary>
		/// <param name="arguments">Up too 254 individual arguments.</param>
		/// <param name="context">Unused, this is information about where the function is being executed.</param>
		/// <returns>The standard deviation based on a sample.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			var args = ArgsToDoubleEnumerable(IgnoreHiddenValues, false, arguments, context);
			foreach (var item in arguments)
			{
				if (item.IsExcelRange)
				{
					if (item.ValueFirst is double || item.ValueFirst is int)
						continue;
					return new CompileResult(eErrorType.Div0);
				}
			}
			if (!TryStandardDeviationOnASamplePopulation(args, arguments, out double StanderedDeviation))
				return new CompileResult(eErrorType.Value);
			return CreateResult(StanderedDeviation, DataType.Decimal);
		}

		private static bool TryStandardDeviationOnASamplePopulation(IEnumerable<double> values, IEnumerable<FunctionArgument> arguments, out double StanderedDeviation)
		{
			List<double> listOfValues = new List<double>();
			foreach (var item in values)
			{
				StanderedDeviation = 0.0;
				var checkThis = ConvertUtil.TryParseDateObjectToOADate(item, out double result12);
				if (!ConvertUtil.TryParseDateObjectToOADate(item, out double result))
					return false;
				listOfValues.Add(result);
			}

			StanderedDeviation = MathObj.Sqrt(Var(listOfValues));
			if (StanderedDeviation == 0 && listOfValues.All(x => x == listOfValues.First()))
				return false;
			return true;
		}

		private static double Var(IEnumerable<double> args)
		{
			double avg = args.Average();
			double d = args.Aggregate(0.0, (total, next) => total += System.Math.Pow(next - avg, 2));
			return (d/ (args.Count() - 1));
		}
	}
}
