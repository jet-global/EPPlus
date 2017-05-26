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
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Operators
{
	/// <summary>
	/// Provides logic for executing an operator within a formula, such as '+', '-', '=', or an ampersand.
	/// </summary>
	public class Operator : IOperator
	{
		#region Constants
		private const int PrecedencePercent = 2;
		private const int PrecedenceExp = 4;
		private const int PrecedenceMultiplyDevide = 6;
		private const int PrecedenceIntegerDivision = 8;
		private const int PrecedenceModulus = 10;
		private const int PrecedenceAddSubtract = 12;
		private const int PrecedenceConcat = 15;
		private const int PrecedenceComparison = 25;
		#endregion

		#region Class Variables
		private readonly int myPrecedence;
		#endregion

		#region Static Operator Implementations
		#region Backing Properties for lazy loading
		private static IOperator myPlus;
		private static IOperator myMinus;
		private static IOperator myMultiply;
		private static IOperator myDivide;
		private static IOperator myExp;
		private static IOperator myConcat;
		private static IOperator myGreaterThan;
		private static IOperator myEqualsTo;
		private static IOperator myNotEqualsTo;
		private static IOperator myGreaterThanOrEqual;
		private static IOperator myLessThan;
		private static IOperator myLessThanOrEqual;
		private static IOperator myPercent;
		#endregion

		/// <summary>
		/// Gets an <see cref="IOperator"/> that can perform the Plus operation.
		/// </summary>
		public static IOperator Plus
		{
			get
			{
				return myPlus ?? (myPlus = new Operator(OperatorType.Plus, PrecedenceAddSubtract, (l, r) =>
				{
					l = l == null || l.Result == null ? new CompileResult(0, DataType.Integer) : l;
					r = r == null || r.Result == null ? new CompileResult(0, DataType.Integer) : r;
					if (EitherIsError(l, r, out ExcelErrorValue errorVal))
					{
						return new CompileResult(errorVal);
					}
					if (l.DataType == DataType.Integer && r.DataType == DataType.Integer)
					{
						return new CompileResult(l.ResultNumeric + r.ResultNumeric, DataType.Integer);
					}
					else if ((l.IsNumeric || l.IsDateString || l.IsNumericString || l.Result is ExcelDataProvider.IRangeInfo) &&
									 (r.IsNumeric || r.IsDateString || r.IsNumericString || r.Result is ExcelDataProvider.IRangeInfo))
					{
						return new CompileResult(l.ResultNumeric + r.ResultNumeric, DataType.Decimal);
					}
					return new CompileResult(eErrorType.Value);
				}));
			}
		}

		/// <summary>
		/// Gets an <see cref="IOperator"/> that can perform the Minus operation.
		/// </summary>
		public static IOperator Minus
		{
			get
			{
				return myMinus ?? (myMinus = new Operator(OperatorType.Minus, PrecedenceAddSubtract, (l, r) =>
				{
					l = l == null || l.Result == null ? new CompileResult(0, DataType.Integer) : l;
					r = r == null || r.Result == null ? new CompileResult(0, DataType.Integer) : r;
					if (l.DataType == DataType.Integer && r.DataType == DataType.Integer)
					{
						return new CompileResult(l.ResultNumeric - r.ResultNumeric, DataType.Integer);
					}
					else if ((l.IsNumeric || l.IsDateString || l.IsNumericString || l.Result is ExcelDataProvider.IRangeInfo) &&
									 (r.IsNumeric || r.IsDateString || r.IsNumericString || r.Result is ExcelDataProvider.IRangeInfo))
					{
						return new CompileResult(l.ResultNumeric - r.ResultNumeric, DataType.Decimal);
					}

					return new CompileResult(eErrorType.Value);
				}));
			}
		}

		/// <summary>
		/// Gets an <see cref="IOperator"/> that can perform the Multiply operation.
		/// </summary>
		public static IOperator Multiply
		{
			get
			{
				return myMultiply ?? (myMultiply = new Operator(OperatorType.Multiply, PrecedenceMultiplyDevide, (l, r) =>
				{
					l = l ?? new CompileResult(0, DataType.Integer);
					r = r ?? new CompileResult(0, DataType.Integer);
					if (l.DataType == DataType.Integer && r.DataType == DataType.Integer)
					{
						return new CompileResult(l.ResultNumeric * r.ResultNumeric, DataType.Integer);
					}
					else if ((l.IsNumeric || l.IsDateString || l.IsNumericString || l.Result is ExcelDataProvider.IRangeInfo) &&
									 (r.IsNumeric || r.IsDateString || r.IsNumericString || r.Result is ExcelDataProvider.IRangeInfo))
					{
						return new CompileResult(l.ResultNumeric * r.ResultNumeric, DataType.Decimal);
					}
					return new CompileResult(eErrorType.Value);
				}));
			}
		}

		/// <summary>
		/// Gets an <see cref="IOperator"/> that can perform the Divide operation.
		/// </summary>
		public static IOperator Divide
		{
			get
			{
				return myDivide ?? (myDivide = new Operator(OperatorType.Divide, PrecedenceMultiplyDevide, (l, r) =>
				{
					if (!(l.IsNumeric || l.IsDateString || l.IsNumericString || l.Result is ExcelDataProvider.IRangeInfo) ||
							  !(r.IsNumeric || r.IsDateString || r.IsNumericString || r.Result is ExcelDataProvider.IRangeInfo))
					{
						return new CompileResult(eErrorType.Value);
					}
					var left = l.ResultNumeric;
					var right = r.ResultNumeric;
					if (Math.Abs(right - 0d) < double.Epsilon)
					{
						return new CompileResult(eErrorType.Div0);
					}
					else if ((l.IsNumeric || l.IsDateString || l.IsNumericString || l.Result is ExcelDataProvider.IRangeInfo) &&
									 (r.IsNumeric || r.IsDateString || r.IsNumericString || r.Result is ExcelDataProvider.IRangeInfo))
					{
						return new CompileResult(left / right, DataType.Decimal);
					}
					return new CompileResult(eErrorType.Value);
				}));
			}
		}

		/// <summary>
		/// Gets an <see cref="IOperator"/> that can perform the Exponentiation operation.
		/// </summary>
		public static IOperator Exp
		{
			get
			{
				return myExp ?? (myExp = new Operator(OperatorType.Exponentiation, PrecedenceExp, (l, r) =>
				{
					if (l == null && r == null)
					{
						return new CompileResult(eErrorType.Value);
					}
					l = l ?? new CompileResult(0, DataType.Integer);
					r = r ?? new CompileResult(0, DataType.Integer);
					if ((l.IsNumeric || l.IsDateString || l.IsNumericString || l.Result is ExcelDataProvider.IRangeInfo) &&
					 (r.IsNumeric || r.IsDateString || r.IsNumericString || r.Result is ExcelDataProvider.IRangeInfo))
					{
						return new CompileResult(Math.Pow(l.ResultNumeric, r.ResultNumeric), DataType.Decimal);
					}
					return new CompileResult(0d, DataType.Decimal);
				}));
			}
		}

		/// <summary>
		/// Gets an <see cref="IOperator"/> that can perform the Concatenate operation.
		/// </summary>
		public static IOperator Concat
		{
			get
			{
				return myConcat ?? (myConcat = new Operator(OperatorType.Concat, PrecedenceConcat, (l, r) =>
				{
					l = l ?? new CompileResult(string.Empty, DataType.String);
					r = r ?? new CompileResult(string.Empty, DataType.String);
					var lStr = Convert.ToString(l.ResultValue);
					var rStr = Convert.ToString(r.ResultValue);
					return new CompileResult(string.Concat(lStr, rStr), DataType.String);
				}));
			}
		}

		/// <summary>
		/// Gets an <see cref="IOperator"/> that can perform the Greater-than operation.
		/// </summary>
		public static IOperator GreaterThan
		{
			get
			{
				return myGreaterThan ??
						 (myGreaterThan =
							  new Operator(OperatorType.GreaterThan, PrecedenceComparison,
									(l, r) => Compare(l, r, (compRes) => compRes > 0)));
			}
		}

		/// <summary>
		/// Gets an <see cref="IOperator"/> that can perform the Equals operation.
		/// </summary>
		public static IOperator EqualsTo
		{
			get
			{
				return myEqualsTo ??
						 (myEqualsTo =
							  new Operator(OperatorType.Equals, PrecedenceComparison,
									(l, r) => Compare(l, r, (compRes) => compRes == 0)));
			}
		}

		/// <summary>
		/// Gets an <see cref="IOperator"/> that can perform the Not-equals operation.
		/// </summary>
		public static IOperator NotEqualsTo
		{
			get
			{
				return myNotEqualsTo ??
						 (myNotEqualsTo =
							  new Operator(OperatorType.NotEqualTo, PrecedenceComparison,
									(l, r) => Compare(l, r, (compRes) => compRes != 0)));
			}
		}

		/// <summary>
		/// Gets an <see cref="IOperator"/> that can perform the Greater-than-or-equal-to operation.
		/// </summary>
		public static IOperator GreaterThanOrEqual
		{
			get
			{
				return myGreaterThanOrEqual ??
						 (myGreaterThanOrEqual =
							  new Operator(OperatorType.GreaterThanOrEqual, PrecedenceComparison,
									(l, r) => Compare(l, r, (compRes) => compRes >= 0)));
			}
		}

		/// <summary>
		/// Gets an <see cref="IOperator"/> that can perform the Less-than operation.
		/// </summary>
		public static IOperator LessThan
		{
			get
			{
				return myLessThan ??
						 (myLessThan =
							  new Operator(OperatorType.LessThan, PrecedenceComparison,
									(l, r) => Compare(l, r, (compRes) => compRes < 0)));
			}
		}

		/// <summary>
		/// Gets an <see cref="IOperator"/> that can perform the Less-than-or-equal-to operation.
		/// </summary>
		public static IOperator LessThanOrEqual
		{
			get
			{
				return myLessThanOrEqual ?? (myLessThanOrEqual = new Operator(OperatorType.LessThanOrEqual, PrecedenceComparison, (l, r) => Compare(l, r, (compRes) => compRes <= 0)));
			}
		}

		/// <summary>
		/// Gets an <see cref="IOperator"/> that can perform the Percent operation.
		/// </summary>
		public static IOperator Percent
		{
			get
			{
				return myPercent ?? (myPercent = new Operator(OperatorType.Percent, PrecedencePercent, (l, r) =>
					{
						l = l ?? new CompileResult(0, DataType.Integer);
						r = r ?? new CompileResult(0, DataType.Integer);
						if (l.DataType == DataType.Integer && r.DataType == DataType.Integer)
						{
							return new CompileResult(l.ResultNumeric * r.ResultNumeric, DataType.Integer);
						}
						else if ((l.IsNumeric || l.IsDateString || l.IsNumericString || l.Result is ExcelDataProvider.IRangeInfo) &&
						 (r.IsNumeric || r.IsDateString || r.IsNumericString || r.Result is ExcelDataProvider.IRangeInfo))
						{
							return new CompileResult(l.ResultNumeric * r.ResultNumeric, DataType.Decimal);
						}
						return new CompileResult(eErrorType.Value);
					}));
			}
		}
		#endregion

		#region Properties
		int IOperator.Precedence
		{
			get { return this.myPrecedence; }
		}

		OperatorType IOperator.Operator
		{
			get { return this.OperatorType; }
		}

		private Func<CompileResult, CompileResult, CompileResult> Implementation { get; }

		private OperatorType OperatorType { get; }
		#endregion

		#region Constructors
		private Operator() { }

		private Operator(OperatorType @operator, int precedence, Func<CompileResult, CompileResult, CompileResult> implementation)
		{
			this.Implementation = implementation;
			this.myPrecedence = precedence;
			this.OperatorType = @operator;
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Applies the specified <see cref="IOperator"/> given the specified <paramref name="left"/> and <paramref name="right"/> arguments.
		/// </summary>
		/// <param name="left">The left argument to the operator.</param>
		/// <param name="right">The right argument to the operator.</param>
		/// <returns>The result of performing the specified operation on the operands.</returns>
		public CompileResult Apply(CompileResult left, CompileResult right)
		{
			if (left.Result is ExcelErrorValue)
			{
				return new CompileResult(left.Result, DataType.ExcelError);
			}
			else if (right.Result is ExcelErrorValue)
			{
				return new CompileResult(right.Result, DataType.ExcelError);
			}
			return this.Implementation(left, right);
		}

		/// <summary>
		/// Gets a string that contains the operator type.
		/// </summary>
		/// <returns>Gets a string that describes the <see cref="Operator.OperatorType"/>.</returns>
		public override string ToString()
		{
			return "Operator: " + this.OperatorType;
		}
		#endregion

		#region Private Static Methods
		private static object GetObjectWithDefaultValueThatMatchesTheOtherObjectType(CompileResult target, CompileResult other)
		{
			if (target.Result == null)
			{
				if (other.DataType == DataType.String) return string.Empty;
				else return 0d;
			}
			return target.ResultValue;
		}

		private static CompileResult Compare(CompileResult left, CompileResult right, Func<int, bool> comparison)
		{
			if (Operator.EitherIsError(left, right, out ExcelErrorValue errorValue))
			{
				return new CompileResult(errorValue);
			}
			return new CompileResult(comparison(Operator.Compare(left, right)), DataType.Boolean);
		}

		private static int Compare(CompileResult leftInput, CompileResult rightInput)
		{
			object left, right;
			left = Operator.GetObjectWithDefaultValueThatMatchesTheOtherObjectType(leftInput, rightInput);
			right = Operator.GetObjectWithDefaultValueThatMatchesTheOtherObjectType(rightInput, leftInput);
			var leftIsNumeric = ConvertUtil.IsNumeric(left) && !(left is bool);
			var rightIsNumeric = ConvertUtil.IsNumeric(right) && !(right is bool);

			if (leftIsNumeric && rightIsNumeric)
			{
				var leftNumber = ConvertUtil.GetValueDouble(left);
				var rightNumber = ConvertUtil.GetValueDouble(right);
				if (leftNumber.Equals(rightNumber))
				{
					return 0;
				}
				return leftNumber.CompareTo(rightNumber);
			}
			// Numbers are less than text are less than logical values: https://stackoverflow.com/questions/35050151/excel-if-statement-comparing-text-with-number
			else if (leftIsNumeric)
				return -1;
			else if (rightIsNumeric)
				return 1;
			else if (leftInput.DataType == DataType.String && rightInput.DataType == DataType.Boolean)
				return -1;
			else if (leftInput.DataType == DataType.Boolean && rightInput.DataType == DataType.String)
				return 1;
			else if (leftInput.DataType == DataType.Boolean && rightInput.DataType == DataType.Boolean)
			{
				if (left.Equals(right))
					return 0;
				else if (left.Equals(true))
					return 1;
				else
					return -1;
			}
			else if (leftInput.DataType == DataType.String && rightInput.DataType == DataType.String)
			{
				var comparisonResult = Operator.CompareString(left, right);
				return comparisonResult;
			}
			throw new InvalidOperationException($"Comparing operands of the given types {leftInput.DataType.ToString()} and {rightInput.DataType.ToString()} is not supported.");
		}

		private static int CompareString(object l, object r)
		{
			var sl = (l ?? "").ToString();
			var sr = (r ?? "").ToString();
			return string.Compare(sl, sr, System.StringComparison.OrdinalIgnoreCase);
		}

		private static bool EitherIsError(CompileResult l, CompileResult r, out ExcelErrorValue errorVal)
		{
			if (l.DataType == DataType.ExcelError)
			{
				errorVal = (ExcelErrorValue)l.Result;
				return true;
			}
			if (r.DataType == DataType.ExcelError)
			{
				errorVal = (ExcelErrorValue)r.Result;
				return true;
			}
			errorVal = null;
			return false;
		}
		#endregion
	}
}
