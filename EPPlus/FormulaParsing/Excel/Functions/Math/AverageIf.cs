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
* Code change notes:
* 
* Author							Change						Date
********************************************************************************
* Mats Alm   		                Added		                2013-12-03
********************************************************************************/
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	/// <summary>
	/// Returns the average (arithmetic mean) of all the cells in a range that meet a given criterion.
	/// </summary>
	public class AverageIf : HiddenValuesHandlingFunction
	{
		/// <summary>
		/// Returns the average (arithmetic mean) of all the cells in a range that meet a given criterion.
		/// </summary>
		/// <param name="arguments">The arguments used to calculate the average.</param>
		/// <param name="context">The context for the function.</param>
		/// <returns>Returns the average of all cells in the given range that passed the given criterion.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentCountIsValid(arguments, 2) == false)
				return new CompileResult(eErrorType.Value);
			var cellRangeToCheck = arguments.ElementAt(0).Value as ExcelDataProvider.IRangeInfo;
			if (cellRangeToCheck == null)
				return new CompileResult(eErrorType.Value);
			var criterionObject = IfHelper.ExtractCriterionObject(arguments.ElementAt(1), context);
			if (arguments.Count() > 2)
			{
				var cellRangeToAverage = arguments.ElementAt(2).Value as ExcelDataProvider.IRangeInfo;
				if (cellRangeToAverage == null)
					return new CompileResult(eErrorType.Value);
				else
					return this.CalculateAverageUsingAverageRange(cellRangeToCheck, criterionObject, cellRangeToAverage);
			}
			else
				return this.CalculateAverageUsingRange(cellRangeToCheck, criterionObject);
		}

		/// <summary>
		/// Calculates the average value of all cells that match the given criterion. The sizes/shapes of
		/// <paramref name="cellsToCompare"/> and <paramref name="potentialCellsToAverage"/> do not have to be the same;
		/// The size and shape of <paramref name="cellsToCompare"/> is applied to <paramref name="potentialCellsToAverage"/>,
		/// using the first cell in <paramref name="potentialCellsToAverage"/> as a reference point.
		/// </summary>
		/// <param name="cellsToCompare">The range of cells to compare against the <paramref name="comparisonCriterion"/>.</param>
		/// <param name="comparisonCriterion">The criterion dictating which cells should be included in the average calculation.</param>
		/// <param name="potentialCellsToAverage">
		///		If a cell in <paramref name="cellsToCompare"/> passes the criterion, then its
		///		corresponding cell in this cell range will be included in the average calculation.</param>
		/// <returns>Returns the average for all cells that pass the <paramref name="comparisonCriterion"/>.</returns>
		private CompileResult CalculateAverageUsingAverageRange(ExcelDataProvider.IRangeInfo cellsToCompare, object comparisonCriterion, ExcelDataProvider.IRangeInfo potentialCellsToAverage)
		{
			var sumOfValidValues = 0d;
			var numberOfValidValues = 0;

			var startingRowForComparison = cellsToCompare.Address._fromRow;
			var startingColumnForComparison = cellsToCompare.Address._fromCol;
			var endingRowForComparison = cellsToCompare.Address._toRow;
			var endingColumnForComparison = cellsToCompare.Address._toCol;

			// This will always look at every cell in the given range of cells to compare. This is done instead of
			// using the iterator provided by the range of cells to compare because the collection of cells that it iterates over
			// does not include empty cells that have not been set since the workbook's creation. This function
			// wants to consider empty cells for comparing with the criterion, but it can be better optimized.
			// A similar problem and optimization opportunity exists in the AverageIfs, SumIf, SumIfs, CountIf, and CountIfs functions.
			for (var currentRow = startingRowForComparison; currentRow <= endingRowForComparison; currentRow++)
			{
				for (var currentColumn = startingColumnForComparison; currentColumn <= endingColumnForComparison; currentColumn++)
				{
					var currentCellValue = this.GetFirstArgument(cellsToCompare.GetValue(currentRow, currentColumn));
					if (IfHelper.ObjectMatchesCriterion(currentCellValue, comparisonCriterion))
					{
						var relativeRow = currentRow - startingRowForComparison;
						var relativeColumn = currentColumn - startingColumnForComparison;
						var valueOfCellToAverage = potentialCellsToAverage.GetOffset(relativeRow, relativeColumn);
						if (valueOfCellToAverage is ExcelErrorValue cellError)
							return new CompileResult(cellError.Type);
						else if (ConvertUtil.IsNumeric(valueOfCellToAverage, true))
						{
							sumOfValidValues += ConvertUtil.GetValueDouble(valueOfCellToAverage);
							numberOfValidValues++;
						}
					}
				}
			}
			if (numberOfValidValues == 0)
				return new CompileResult(eErrorType.Div0);
			else
				return this.CreateResult(sumOfValidValues / numberOfValidValues, DataType.Decimal);
		}

		/// <summary>
		/// Calculates the average value of all cells that match the given criterion.
		/// </summary>
		/// <param name="potentialCellsToAverage">
		///		The cell range to compare against the given <paramref name="comparisonCriterion"/>
		///		If a cell passes the criterion, then its value is included in the average calculation.</param>
		/// <param name="comparisonCriterion">The criterion dictating which cells should be included in the average calculation.</param>
		/// <returns>Returns the average value for all cells that pass the <paramref name="comparisonCriterion"/>.</returns>
		private CompileResult CalculateAverageUsingRange(ExcelDataProvider.IRangeInfo potentialCellsToAverage, object comparisonCriterion)
		{
			var sumOfValidValues = 0d;
			var numberOfValidValues = 0;
			var valuesToAverage = potentialCellsToAverage.Select(cell => this.GetFirstArgument(cell.Value)).Where(cellValue => IfHelper.ObjectMatchesCriterion(cellValue, comparisonCriterion));
			foreach (var value in valuesToAverage)
			{
				if (value is ExcelErrorValue cellErrorValue)
					return new CompileResult(cellErrorValue.Type);
				else if (ConvertUtil.IsNumeric(value, true))
				{
					sumOfValidValues += ConvertUtil.GetValueDouble(value);
					numberOfValidValues++;
				}
			}
			if (numberOfValidValues == 0)
				return new CompileResult(eErrorType.Div0);
			else
				return this.CreateResult(sumOfValidValues / numberOfValidValues, DataType.Decimal);
		}
	}
}
