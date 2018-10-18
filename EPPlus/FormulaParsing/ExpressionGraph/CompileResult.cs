/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See http://www.codeplex.com/EPPlus for details.
 *
 * Copyright (C) 2011  Jan Källman
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
 * ******************************************************************************
 * Mats Alm   		                Added       		        2013-03-01 (Prior file history on https://github.com/swmal/ExcelFormulaParser)
 *******************************************************************************/
using System;
using System.Globalization;
using System.Linq;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.ExpressionGraph
{
	/// <summary>
	/// Represents the resulting value of function compilation.
	/// </summary>
	public class CompileResult
	{
		#region Class Variables
		private double? myResultNumeric;
		#endregion

		#region Properties
		/// <summary>
		/// Gets an empty compile result.
		/// </summary>
		public static CompileResult Empty
		{
			get { return EmptyResult; }
		}

		/// <summary>
		/// Gets the result of the <see cref="CompileResult"/>.
		/// </summary>
		public object Result { get; private set; }

		/// <summary>
		/// Gets the result of the <see cref="CompileResult"/>, resolving range references if applicable.
		/// </summary>
		public object ResultValue
		{
			get
			{
				if (this.Result is ExcelDataProvider.IRangeInfo rangeResult)
					return rangeResult.GetValue(rangeResult.Address._fromRow, rangeResult.Address._fromCol);
				else
					return this.Result;
			}
		}

		/// <summary>
		/// Gets the result of the <see cref="CompileResult"/> as a numeric value.
		/// </summary>
		public double ResultNumeric
		{
			get
			{
				// We assume that Result does not change unless it is a range.
				if (myResultNumeric == null)
				{
					if (this.Result is DateTime)
						myResultNumeric = ((DateTime)this.Result).ToOADate();
					else if (this.IsNumeric)
						myResultNumeric = this.Result == null ? 0 : Convert.ToDouble(this.Result);
					else if (this.Result is TimeSpan)
						myResultNumeric = DateTime.FromOADate(0).Add((TimeSpan)this.Result).ToOADate();
					else if (this.Result is ExcelDataProvider.IRangeInfo)
					{
						var c = ((ExcelDataProvider.IRangeInfo)this.Result).FirstOrDefault();
						return c?.ValueDoubleLogical ?? 0;
					}
					// The IsNumericString and IsDateString properties will set _ResultNumeric for efficiency so we just need
					// to check them here.
					else if (!this.IsNumericOrDateString)
						myResultNumeric = 0;
				}
				return myResultNumeric.Value;
			}
		}

		/// <summary>
		/// Gets the <see cref="DataType"/> of the <see cref="CompileResult"/>.
		/// </summary>
		public DataType DataType { get; private set; }

		/// <summary>
		/// Gets a value indicating whether the <see cref="CompileResult"/> is numeric.
		/// </summary>
		public bool IsNumeric
		{
			get
			{
				return this.DataType == DataType.Decimal 
					|| this.DataType == DataType.Boolean 
					|| this.DataType == DataType.Integer 
					|| this.DataType == DataType.Empty 
					|| this.DataType == DataType.Date
					|| this.DataType == DataType.Time;
			}
		}

		/// <summary>
		/// Gets a value determining if this <see cref="CompileResult"/> is a string that can parse to a number or date.
		/// </summary>
		public bool IsNumericOrDateString
		{
			get
			{
				if (this.DataType == DataType.String)
				{
					bool isNumber = ConvertUtil.TryParseNumericString(this.Result, out var doubleResult);
					bool isDate = ConvertUtil.TryParseDateString(this.Result, out var dateResult);
					if (isNumber)
					{
						// If we parse to a number and not a date then we're a number.
						if (!isDate)
							myResultNumeric = doubleResult;
						// If we parse as both a number and a date then we need to validate number group sizes.
						else if (ConvertUtil.ValidateNumberGroupSizes(this.Result.ToString(), CultureInfo.CurrentCulture.NumberFormat))
							myResultNumeric = doubleResult;
						// If number group sizes are incorrect then we are a date.
						else
							myResultNumeric = dateResult.ToOADate();
						return true;
					}
					else if (isDate)
					{
						myResultNumeric = dateResult.ToOADate();
						return true;
					}
				}
				return false;
			}
		}

		/// <summary>
		/// Gets or sets a value indicating whether the <see cref="CompileResult"/> is a result of a subtotal.
		/// </summary>
		public bool IsResultOfSubtotal { get; set; }

		/// <summary>
		/// Gets or sets a value indicating whether the <see cref="CompileResult"/> is for a hidden cell.
		/// </summary>
		public bool IsHiddenCell { get; set; }

		private static CompileResult EmptyResult { get; } = new CompileResult(null, DataType.Empty);
		#endregion

		#region Constructors
		/// <summary>
		/// Instantiates a new <see cref="CompileResult"/>.
		/// </summary>
		/// <param name="result">The result value.</param>
		/// <param name="dataType">The type of the result value.</param>
		public CompileResult(object result, DataType dataType)
		{
			this.Result = result;
			this.DataType = dataType;
		}

		/// <summary>
		/// Instantiates a new <see cref="CompileResult"/> for an error type.
		/// </summary>
		/// <param name="errorType">The error type.</param>
		public CompileResult(eErrorType errorType)
		{
			this.Result = ExcelErrorValue.Create(errorType);
			this.DataType = DataType.ExcelError;
		}

		/// <summary>
		/// Instantiates a new <see cref="CompileResult"/> for an error value.
		/// </summary>
		/// <param name="errorValue">The <see cref="ExcelErrorValue"/> result.</param>
		public CompileResult(ExcelErrorValue errorValue)
		{
			if (errorValue == null)
				throw new ArgumentNullException(nameof(errorValue));
			this.Result = errorValue;
			this.DataType = DataType.ExcelError;
		}
		#endregion
	}
}
