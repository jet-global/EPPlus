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
	/// Returns the average (arithmetic mean) of all the cells in a range that meet a given criteria.
	/// </summary>
	public class AverageIf : HiddenValuesHandlingFunction
	{
		/// <summary>
		/// Returns the average (arithmetic mean) of all the cells in a range that meet a given criteria.
		/// </summary>
		/// <param name="arguments">The arguments used to calculate the average.</param>
		/// <param name="context">The context for the function.</param>
		/// <returns>Returns the average of all cells in the given range that passed the given criteria.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentCountIsValid(arguments, 2) == false)
				return new CompileResult(eErrorType.Value);
			var cellRangeToCheck = arguments.ElementAt(0).Value as ExcelDataProvider.IRangeInfo;
			if (cellRangeToCheck == null)
				return new CompileResult(eErrorType.Value);
			//var criteriaString = IfHelper.ExtractCriteriaString(arguments.ElementAt(1), context);
			var criteriaObject = IfHelper.ExtractCriteriaObject(arguments.ElementAt(1), context);
			if (arguments.Count() > 2)
			{
				var cellRangeToAverage = arguments.ElementAt(2).Value as ExcelDataProvider.IRangeInfo;
				if (cellRangeToAverage == null)
					return new CompileResult(eErrorType.Value);
				else
					return this.CalculateAverageUsingAverageRange(cellRangeToCheck, criteriaObject, cellRangeToAverage);
			}
			else
				return this.CalculateAverageUsingRange(cellRangeToCheck, criteriaObject);
		}

		/// <summary>
		/// Calculates the average value of all cells that match the given criteria. The sizes/shapes of
		/// <paramref name="cellsToCompare"/> and <paramref name="potentialCellsToAverage"/> do not have to be the same;
		/// The size and shape of <paramref name="cellsToCompare"/> is applied to <paramref name="potentialCellsToAverage"/>,
		/// using the first cell in <paramref name="potentialCellsToAverage"/> as a reference point.
		/// </summary>
		/// <param name="cellsToCompare">The range of cells to compare against the <paramref name="comparisonCriteria"/>.</param>
		/// <param name="comparisonCriteria">The criteria dictating which cells should be included in the average calculation.</param>
		/// <param name="potentialCellsToAverage">
		///		If a cell in <paramref name="cellsToCompare"/> passes the criteria, then its
		///		corresponding cell in this cell range will be included in the average calculation.</param>
		/// <returns>Returns the average for all cells that pass the <paramref name="comparisonCriteria"/>.</returns>
		private CompileResult CalculateAverageUsingAverageRange(ExcelDataProvider.IRangeInfo cellsToCompare, object comparisonCriteria, ExcelDataProvider.IRangeInfo potentialCellsToAverage)
		{
			var sumOfValidValues = 0d;
			var numberOfValidValues = 0;
			foreach (var cell in cellsToCompare)
			{
				if (comparisonCriteria != null && IfHelper.ObjectMatchesCriteria(this.GetFirstArgument(cell.Value), comparisonCriteria))
				{
					var relativeRow = cell.Row - cellsToCompare.Address._fromRow;
					var relativeColumn = cell.Column - cellsToCompare.Address._fromCol;
					var valueOfCellToAverage = potentialCellsToAverage.GetOffset(relativeRow, relativeColumn);
					if (valueOfCellToAverage is ExcelErrorValue cellError)
						return new CompileResult(cellError.Type);
					if (valueOfCellToAverage is string || valueOfCellToAverage is bool || valueOfCellToAverage == null)
						continue;
					sumOfValidValues += ConvertUtil.GetValueDouble(valueOfCellToAverage, true);
					numberOfValidValues++;
				}
			}
			if (numberOfValidValues == 0)
				return new CompileResult(eErrorType.Div0);
			else
				return this.CreateResult(sumOfValidValues / numberOfValidValues, DataType.Decimal);
		}

		/// <summary>
		/// Calculates the average value of all cells that match the given criteria.
		/// </summary>
		/// <param name="potentialCellsToAverage">
		///		The cell range to compare against the given <paramref name="comparisonCriteria"/>
		///		If a cell passes the criteria, then its value is included in the average calculation.</param>
		/// <param name="comparisonCriteria">The criteria dictating which cells should be included in the average calculation.</param>
		/// <returns>Returns the average value for all cells that pass the <paramref name="comparisonCriteria"/>.</returns>
		private CompileResult CalculateAverageUsingRange(ExcelDataProvider.IRangeInfo potentialCellsToAverage, object comparisonCriteria)
		{
			var sumOfValidValues = 0d;
			var numberOfValidValues = 0;
			foreach (var cell in potentialCellsToAverage)
			{
				var cellValue = this.GetFirstArgument(cell.Value);
				if (comparisonCriteria != null && IfHelper.ObjectMatchesCriteria(cellValue, comparisonCriteria))
				{
					if (cellValue is ExcelErrorValue cellErrorValue)
						return new CompileResult(cellErrorValue.Type);
					else if (IfHelper.IsNumeric(cellValue, true))
					{
						sumOfValidValues += cell.ValueDouble;
						numberOfValidValues++;
					}
				}
				//else if (cell.Value is ExcelErrorValue candidateError)
				//	return new CompileResult(candidateError.Type);
			}
			if (numberOfValidValues == 0)
				return new CompileResult(eErrorType.Div0);
			else
				return this.CreateResult(sumOfValidValues / numberOfValidValues, DataType.Decimal);
		}
	}
}
