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
using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
	/// <summary>
	/// This class contains the formula for obtaining the value from a given cell range with a specified 
	/// row or column value. 
	/// </summary>
	public class Index : ExcelFunction
	{
		/// <summary>
		/// Takes the data and the associated row/column value and returns the value from the cell range
		/// at the given row or column value. 
		/// </summary>
		/// <param name="arguments">The cell range, the row value, and the column value.</param>
		/// <param name="context">The context in which the function is called.</param>
		/// <returns>A <see cref="CompileResult"/> result.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentsAreValid(arguments, 2, out eErrorType argumentError) == false)
				return new CompileResult(argumentError);
			var cellRange = arguments.ElementAt(0);
			var rowDataCandidate = arguments.ElementAt(1);

			var result = new CompileResultFactory();
			if (cellRange.Value is IEnumerable<FunctionArgument> args)
			{
				var index = this.ArgToInt(arguments, 1);
				if (index > args.Count())
					throw new ExcelErrorValueException(eErrorType.Ref);
				var candidate = args.ElementAt(index - 1);
				return base.CreateResult(candidate.Value, DataType.Integer);
			}
			if (rowDataCandidate != null && rowDataCandidate.Value is ExcelErrorValue && rowDataCandidate.Value.ToString() == ExcelErrorValue.Values.NA)
				return base.CreateResult(rowDataCandidate.Value, DataType.Integer);

			// A single cell array ignores the row number and column number arguments, returning the single value in the array.
			if (!cellRange.IsExcelRange)
				return new CompileResult(cellRange.Value, cellRange.DataType);
			else
			{
				var rowCandidate = arguments.ElementAt(1).DataType;
				if (rowCandidate == DataType.Date)
					return new CompileResult(eErrorType.Ref);
				else if (rowCandidate == DataType.Decimal)
					return new CompileResult(eErrorType.Ref);
				var row = this.ArgToInt(arguments, 1);

				if (row == 0)
					return new CompileResult(eErrorType.Value);
				if (row < 0)
					return new CompileResult(eErrorType.Value);

				var column = 1;
				if (arguments.Count() > 2)
				{
					var colCandidate = arguments.ElementAt(2).DataType;
					if (colCandidate == DataType.Date)
						return new CompileResult(eErrorType.Ref);
					else if (colCandidate == DataType.Decimal)
						return new CompileResult(eErrorType.Ref);
					else
						column = this.ArgToInt(arguments, 2);
				}
				else
					if ((arguments.ElementAt(0).ValueAsRangeInfo.Address.Columns > 1) && arguments.ElementAt(0).ValueAsRangeInfo.Address.Rows > 1)
						return new CompileResult(eErrorType.Ref);
				if ((column == 0 && row == 0) || column < 0)
					return new CompileResult(eErrorType.Value);

				var rangeInfo = cellRange.ValueAsRangeInfo;
				if (rangeInfo.Address.Rows == 1 && arguments.Count() < 3)
				{
					column = row;
					row = 1;
				}
				var numColumns = arguments.ElementAt(0).ValueAsRangeInfo.Address.Columns;
				if ((numColumns > 1 && column == 0))
					return new CompileResult(eErrorType.Value);
				if (row > rangeInfo.Address.Rows || column > rangeInfo.Address.Columns)
					return new CompileResult(eErrorType.Ref);
				if (row > rangeInfo.Address._toRow - rangeInfo.Address._fromRow + 1 || column > rangeInfo.Address._toCol - rangeInfo.Address._fromCol + 1)
					return new CompileResult(eErrorType.Value);
				var candidate = rangeInfo.GetOffset(row - 1, column - 1);
				if (column == 0)
					candidate = rangeInfo.GetOffset(row - 1, column);
				return base.CreateResult(candidate, DataType.Integer);
			}
		}
	}
}
