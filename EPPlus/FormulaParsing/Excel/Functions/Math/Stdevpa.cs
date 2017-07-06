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
	public class Stdevpa : HiddenValuesHandlingFunction
	{
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			List<double> listToDoStandardDeviationOn = new List<double>();
			var args = ArgsToDoubleEnumerable(this.IgnoreHiddenValues, false, arguments, context);
			foreach (var item in arguments)
			{


				if (item.IsExcelRange)
				{
					foreach (var cell in item.ValueAsRangeInfo)
					{
						if (cell.Value is string)
							listToDoStandardDeviationOn.Add(0.0);
					}
					if (item.ValueFirst is double || item.ValueFirst is int || item.ValueFirst == null)
						continue;
				}
				if (item.ValueFirst == null)
				{
					listToDoStandardDeviationOn.Add(0.0);
				}
				if(item.Value is ExcelDataProvider.IRangeInfo itemRange)
				{
					if(item.ValueFirst is bool valueIsABool)
					{
						if (valueIsABool == true)
						{
							listToDoStandardDeviationOn.Add(1);
							continue;
						}
						else
						{
							listToDoStandardDeviationOn.Add(0);
							continue;
						}
					}
					return this.CreateResult(0d, DataType.Decimal);
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

		private bool TryStandardDeviationEntirePopulation(List<double> listToDoStandardDeviationOn, out double standardDeviation)
		{
			standardDeviation = MathObj.Sqrt(this.VarPopulation(listToDoStandardDeviationOn));
			if (standardDeviation == 0 && listToDoStandardDeviationOn.All(x => x == -1))
				return false;
			return true;
		}

		private double VarPopulation(List<double> listOfDoubles)
		{
			double avg = listOfDoubles.Average();
			double d = listOfDoubles.Aggregate(0.0, (total, next) => total += System.Math.Pow(next - avg, 2));
			return (d / (listOfDoubles.Count()));
		}
	}
}
