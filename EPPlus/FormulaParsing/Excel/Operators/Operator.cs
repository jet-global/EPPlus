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
		private const int PrecedenceMultiplyDivide = 6;
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
		#endregion

		/// <summary>
		/// Gets an <see cref="IOperator"/> that can perform the Plus operation.
		/// </summary>
		public static IOperator Plus
		{
			get
			{
				CompileResult add(CompileResult l, CompileResult r)
				{
					var dataType = Operator.ParseAdditiveOperatorDataType(l.DataType, r.DataType);
					return Operator.CalculateNumericalOperator(l, r, () => new CompileResult(l.ResultNumeric + r.ResultNumeric, dataType));
				}
				return myPlus ?? (myPlus = new Operator(OperatorType.Plus, Operator.PrecedenceAddSubtract, add));
			}
		}

		/// <summary>
		/// Gets an <see cref="IOperator"/> that can perform the Minus operation.
		/// </summary>
		public static IOperator Minus
		{
			get
			{
				CompileResult subtract(CompileResult l, CompileResult r)
				{
					var dataType = Operator.ParseAdditiveOperatorDataType(l.DataType, r.DataType);
					return Operator.CalculateNumericalOperator(l, r, () => new CompileResult(l.ResultNumeric - r.ResultNumeric, dataType));
				} 
				return myMinus ?? (myMinus = new Operator(OperatorType.Minus, Operator.PrecedenceAddSubtract, subtract));
			}
		}

		/// <summary>
		/// Gets an <see cref="IOperator"/> that can perform the Multiply operation.
		/// </summary>
		public static IOperator Multiply
		{
			get
			{
				CompileResult multiply(CompileResult l, CompileResult r)
				{
					var dataType = Operator.ParseMultiplyOperatorDataType(l.DataType, r.DataType);
					return Operator.CalculateNumericalOperator(l, r, () => new CompileResult(l.ResultNumeric * r.ResultNumeric, dataType));
				}
				return myMultiply ?? (myMultiply = new Operator(OperatorType.Multiply, Operator.PrecedenceMultiplyDivide, multiply));
			}
		}

		/// <summary>
		/// Gets an <see cref="IOperator"/> that can perform the Divide operation.
		/// </summary>
		public static IOperator Divide
		{
			get
			{
				CompileResult divide(CompileResult l, CompileResult r)
				{
					return Operator.CalculateNumericalOperator(l, r, () =>
					{
						if (Math.Abs(r.ResultNumeric) < double.Epsilon)
							return new CompileResult(eErrorType.Div0);
						var dataType = Operator.ParseDivideOperatorDataType(l.DataType, r.DataType);
						return new CompileResult(l.ResultNumeric / r.ResultNumeric, dataType);
					});
				}
				return myDivide ?? (myDivide = new Operator(OperatorType.Divide, Operator.PrecedenceMultiplyDivide, divide));
			}
		}

		/// <summary>
		/// Gets an <see cref="IOperator"/> that can perform the Exponentiation operation.
		/// </summary>
		public static IOperator Exp
		{
			get
			{
				CompileResult exponentiate(CompileResult l, CompileResult r)
				{
					var dataType = Operator.ParseGeneralOperatorDataType(l.DataType, r.DataType);
					return Operator.CalculateNumericalOperator(l, r, () => new CompileResult(Math.Pow(l.ResultNumeric, r.ResultNumeric), dataType));
				}
				return myExp ?? (myExp = new Operator(OperatorType.Exponentiation, Operator.PrecedenceExp, exponentiate));
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
					if (l.DataType == DataType.ExcelError)
						return new CompileResult(l.Result as ExcelErrorValue);
					else if (r.DataType == DataType.ExcelError)
						return new CompileResult(r.Result as ExcelErrorValue);
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
					(myGreaterThan = new Operator(
						OperatorType.GreaterThan,
						PrecedenceComparison,
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
					(myEqualsTo = new Operator(
						OperatorType.Equals,
						PrecedenceComparison,
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
					(myNotEqualsTo = new Operator(
						OperatorType.NotEqualTo,
						PrecedenceComparison,
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
					(myGreaterThanOrEqual = new Operator(
						OperatorType.GreaterThanOrEqual,
						PrecedenceComparison,
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
					(myLessThan = new Operator(
						OperatorType.LessThan, 
						PrecedenceComparison,
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
				return myLessThanOrEqual ?? 
					(myLessThanOrEqual = new Operator(
						OperatorType.LessThanOrEqual, 
						PrecedenceComparison, 
						(l, r) => Compare(l, r, (compRes) => compRes <= 0)));
			}
		}

		/// <summary>
		/// Gets an <see cref="IOperator"/> that can perform the 
		/// Excel Percent operation (multiply by .01).
		/// </summary>
		/// <remarks>Note that the right operand to this operator has been set to .01.</remarks>
		public static IOperator Percent => Operator.Multiply;
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
		/// If the left operand is an error type, its error type is returned. If the right operand is an error type,
		/// a #VALUE! error type is returned.
		/// </summary>
		/// <param name="left">The left argument to the operator.</param>
		/// <param name="right">The right argument to the operator.</param>
		/// <returns>The result of performing the specified operation on the operands.</returns>
		public CompileResult Apply(CompileResult left, CompileResult right)
		{
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
		private static CompileResult CalculateNumericalOperator(CompileResult left, CompileResult right, Func<CompileResult> operation)
		{
			if (left == null || right == null)
				return new CompileResult(eErrorType.Value);
			else if (left.DataType == DataType.ExcelError)
				return new CompileResult(left.Result as ExcelErrorValue);
			else if (!Operator.IsNumericType(left))
				return new CompileResult(eErrorType.Value);
			else if (right.DataType == DataType.ExcelError)
				return new CompileResult(right.Result as ExcelErrorValue);
			else if (!Operator.IsNumericType(right))
				return new CompileResult(eErrorType.Value);
			return operation();
		}

		private static bool IsNumericType(CompileResult result)
		{
			return result.IsNumeric || result.IsNumericOrDateString || result.Result is ExcelDataProvider.IRangeInfo;
		}

		private static CompileResult GetObjectWithDefaultValueThatMatchesTheOtherObjectType(CompileResult target, CompileResult other)
		{
			if (target.Result == null)
			{
				if (other.DataType == DataType.String)
					return new CompileResult(string.Empty, other.DataType);
				else if (other.DataType == DataType.Boolean)
					return new CompileResult(false, other.DataType);
				else
					return new CompileResult(0d, other.DataType);
			}
			return target;
		}

		private static CompileResult Compare(CompileResult left, CompileResult right, Func<int, bool> comparison)
		{
			if (Operator.EitherIsError(left, right, out ExcelErrorValue errorValue))
				return new CompileResult(errorValue);
			return new CompileResult(comparison(Operator.Compare(left, right)), DataType.Boolean);
		}

		private static int Compare(CompileResult leftInput, CompileResult rightInput)
		{
			CompileResult leftMatch = Operator.GetObjectWithDefaultValueThatMatchesTheOtherObjectType(leftInput, rightInput);
			CompileResult rightMatch = Operator.GetObjectWithDefaultValueThatMatchesTheOtherObjectType(rightInput, leftInput);
			object left = leftMatch.ResultValue;
			object right = rightMatch.ResultValue;
			var leftIsNumeric = ConvertUtil.IsNumeric(left) && !(left is bool);
			var rightIsNumeric = ConvertUtil.IsNumeric(right) && !(right is bool);

			if (leftIsNumeric && rightIsNumeric)
			{
				var leftNumber = ConvertUtil.GetValueDouble(left);
				var rightNumber = ConvertUtil.GetValueDouble(right);
				if (leftNumber.Equals(rightNumber))
					return 0;
				return leftNumber.CompareTo(rightNumber);
			}
			// Numbers are less than text are less than logical values: https://stackoverflow.com/questions/35050151/excel-if-statement-comparing-text-with-number
			else if (leftIsNumeric)
				return -1;
			else if (rightIsNumeric)
				return 1;
			else if (leftMatch.DataType == DataType.String && rightMatch.DataType == DataType.Boolean)
				return -1;
			else if (leftMatch.DataType == DataType.Boolean && rightMatch.DataType == DataType.String)
				return 1;
			else if (leftMatch.DataType == DataType.Boolean && rightMatch.DataType == DataType.Boolean)
			{
				if (left.Equals(right))
					return 0;
				else if (left.Equals(true))
					return 1;
				else
					return -1;
			}
			else if (leftMatch.DataType == DataType.String && rightMatch.DataType == DataType.String)
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

		private static DataType ParseGeneralOperatorDataType(DataType leftType, DataType rightType)
		{
			return (leftType == DataType.Integer && rightType == DataType.Integer) ? DataType.Integer : DataType.Decimal;
		}

		private static DataType ParseAdditiveOperatorDataType(DataType leftType, DataType rightType)
		{
			if (leftType == DataType.Date)
				return DataType.Date;
			else if (leftType == DataType.Time)
				return DataType.Time;
			else
				return Operator.ParseGeneralOperatorDataType(leftType, rightType);
		}

		private static DataType ParseMultiplyOperatorDataType(DataType leftType, DataType rightType)
		{
			return Operator.ParseMultiplicativeDateTimeOperatorDataType(leftType, rightType) 
				?? Operator.ParseGeneralOperatorDataType(leftType, rightType);
		}

		private static DataType ParseDivideOperatorDataType(DataType leftType, DataType rightType)
		{
			return Operator.ParseMultiplicativeDateTimeOperatorDataType(leftType, rightType) 
				?? DataType.Decimal;
		}

		private static DataType? ParseMultiplicativeDateTimeOperatorDataType(DataType leftType, DataType rightType)
		{
			if (leftType == DataType.Date && rightType == DataType.Date)
				return DataType.Decimal;
			else if (leftType == DataType.Date && rightType == DataType.Time)
				return DataType.Date;
			else if (leftType == DataType.Time && rightType == DataType.Date)
				return DataType.Time;
			else if (leftType == DataType.Time && rightType == DataType.Time)
				return DataType.Time;
			return null;
		}
		#endregion
	}
}
