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
 * Mats Alm   		                Added		                2015-02-01
 *******************************************************************************/
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	public class AverageIfs : MultipleRangeCriteriasFunction
	{
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentCountIsValid(arguments, 3) == false)
				return new CompileResult(eErrorType.Div0);
			

			var functionArguments = arguments as FunctionArgument[] ?? arguments.ToArray();
			var numberOfArguments = functionArguments.Count();
			//var sumRange = ArgsToDoubleEnumerable(true, new List<FunctionArgument> { functionArguments[0] }, context).ToList();
			var sumRange = functionArguments[0].Value as ExcelDataProvider.IRangeInfo;
			if (sumRange == null)
				return new CompileResult(eErrorType.Div0);
			var sumRangeLength = sumRange.Count();
			var argRanges = new List<ExcelDataProvider.IRangeInfo>();
			var criterias = new List<string>();
			var averageIndicesToIgnore = new List<int>();

			for (var currentArgumentIndex = 1; currentArgumentIndex < numberOfArguments; currentArgumentIndex += 2)
			{
				var currentRange = functionArguments[currentArgumentIndex].ValueAsRangeInfo;
				if (currentRange == null)
					return new CompileResult(eErrorType.Value);
				if (currentRange.Count() != sumRangeLength)
					return new CompileResult(eErrorType.Value);
				string criteria = null;
				var thing = functionArguments[currentArgumentIndex + 1];
				if (functionArguments[currentArgumentIndex + 1].Value is ExcelDataProvider.IRangeInfo criteriaRange)
				{
					if (criteriaRange.IsMulti)
						return new CompileResult(eErrorType.Div0);
					else
						criteria = this.GetFirstArgument(thing.ValueFirst).ToString().ToUpper();
				}
				else
					criteria = this.GetFirstArgument(functionArguments[currentArgumentIndex + 1]).ValueFirst.ToString().ToUpper();
				if (criteria == null)
					return new CompileResult(eErrorType.Div0);
				for (var currentRangeIndex = 0; currentRangeIndex < currentRange.Count(); currentRangeIndex++)
				{
					var currentCell = currentRange.ElementAt(currentRangeIndex);
					if (!IfHelper.objectMatchesCriteria(this.GetFirstArgument(currentCell.Value), criteria))
						averageIndicesToIgnore.Add(currentRangeIndex);
				}
			}

			var sumOfValidValues = 0d;
			var numberOfValidValues = 0;
			
			for (var currentAverageIndex = 0; currentAverageIndex < sumRange.Count(); currentAverageIndex++)
			{
				if (averageIndicesToIgnore.Contains(currentAverageIndex))
					continue;
				var currentCellValue = sumRange.ElementAt(currentAverageIndex).Value;
				if (currentCellValue is ExcelErrorValue cellError)
					return new CompileResult(cellError.Type);
				if (currentCellValue is string || currentCellValue is bool || currentCellValue == null)
					continue;
				sumOfValidValues += ConvertUtil.GetValueDouble(currentCellValue);
				numberOfValidValues++;
			}

			if (numberOfValidValues == 0)
				return new CompileResult(eErrorType.Div0);
			else
				return this.CreateResult(sumOfValidValues / numberOfValidValues, DataType.Decimal);

			//for (var ix = 1; ix < 31; ix += 2)
			//{
			//	if (functionArguments.Length <= ix) break;
			//	var rangeInfo = functionArguments[ix].ValueAsRangeInfo;
			//	argRanges.Add(rangeInfo);
			//	var value = functionArguments[ix + 1].Value != null ? functionArguments[ix + 1].Value.ToString() : null;
			//	criterias.Add(value);
			//}
			//IEnumerable<int> matchIndexes = GetMatchIndexes(argRanges[0], criterias[0]);
			//var enumerable = matchIndexes as IList<int> ?? matchIndexes.ToList();
			//for (var ix = 1; ix < argRanges.Count && enumerable.Any(); ix++)
			//{
			//	var indexes = GetMatchIndexes(argRanges[ix], criterias[ix]);
			//	matchIndexes = enumerable.Intersect(indexes);
			//}

			//var result = matchIndexes.Average(index => sumRange[index]);

			//return CreateResult(result, DataType.Decimal);
		}
	}
}
