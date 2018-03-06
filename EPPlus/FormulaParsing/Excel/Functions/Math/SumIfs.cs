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
 * Mats Alm   		                Added		                2015-01-15
 *******************************************************************************/
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	/// <summary>
	/// Returns the sum of all cells that meet multiple criteria.
	/// </summary>
	public class SumIfs : MultipleRangeCriteriasFunction
	{
		/// <summary>
		/// Returns the sum of all cells that meet multiple criteria.
		/// </summary>
		/// <param name="arguments">The arguments used to calculate the sum.</param>
		/// <param name="context">The context for the function.</param>
		/// <returns>Returns the sum of all cells in the given range that pass the given criteria.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (!this.ArgumentsAreValid(arguments, 3, out eErrorType errorType))
				return new CompileResult(errorType);
			var sumRange = arguments.ElementAt(0).Value as ExcelDataProvider.IRangeInfo;
			if (sumRange == null)
				return new CompileResult(0d, DataType.Decimal);
			var indicesOfValidCells = new List<int>();
			for (var argumentIndex = 1; argumentIndex < arguments.Count(); argumentIndex += 2)
			{
				var currentRangeToCompare = arguments.ElementAt(argumentIndex).ValueAsRangeInfo;
				if (currentRangeToCompare == null || !IfHelper.RangesAreTheSameShape(sumRange, currentRangeToCompare))
					return new CompileResult(eErrorType.Value);
				var currentCriterion = IfHelper.ExtractCriterionObject(arguments.ElementAt(argumentIndex + 1), context);

				// This will always look at every cell in the given range of cells to compare. This is done instead of
				// using the iterator provided by the range of cells to compare because the collection of cells that it iterates over
				// does not include empty cells that have not been set since the workbook's creation. This function
				// wants to consider empty cells for comparing with the criterion, but it can be better optimized.
				// A similar problem and optimization opportunity exists in the AverageIf, AverageIfs, SumIf, CountIf, and CountIfs functions.
				var passingIndices = IfHelper.GetIndicesOfCellsPassingCriterion(currentRangeToCompare, currentCriterion);
				if (argumentIndex == 1)
					indicesOfValidCells = passingIndices;
				else
					indicesOfValidCells = indicesOfValidCells.Intersect(passingIndices).ToList();
			}
			double sumOfValidValues = 0d;
			if (sumRange.Count() > 0)
			{
				// Again, all cells, including empty cells, need to be available here. 
				// The IRangeInfo will only provide non-empty cells.
				var allSumValues = sumRange.AllValues();
				foreach (var cellIndex in indicesOfValidCells)
				{
					var currentCellValue = allSumValues.ElementAt(cellIndex);
					if (currentCellValue is ExcelErrorValue cellError)
						return new CompileResult(cellError.Type);
					else if (ConvertUtil.IsNumeric(currentCellValue, true))
						sumOfValidValues += ConvertUtil.GetValueDouble(currentCellValue);
				}
			}
			return this.CreateResult(sumOfValidValues, DataType.Decimal);
		}
	}
}
