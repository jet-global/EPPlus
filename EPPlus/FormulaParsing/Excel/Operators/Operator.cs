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
	public class Operator : IOperator
	{
		private const int PrecedencePercent = 2;
		private const int PrecedenceExp = 4;
		private const int PrecedenceMultiplyDevide = 6;
		private const int PrecedenceIntegerDivision = 8;
		private const int PrecedenceModulus = 10;
		private const int PrecedenceAddSubtract = 12;
		private const int PrecedenceConcat = 15;
		private const int PrecedenceComparison = 25;

		private Operator() { }

		private Operator(Operators @operator, int precedence, Func<CompileResult, CompileResult, CompileResult> implementation)
		{
			_implementation = implementation;
			_precedence = precedence;
			_operator = @operator;
		}

		private readonly Func<CompileResult, CompileResult, CompileResult> _implementation;
		private readonly int _precedence;
		private readonly Operators _operator;

		int IOperator.Precedence
		{
			get { return _precedence; }
		}

		Operators IOperator.Operator
		{
			get { return _operator; }
		}

		public CompileResult Apply(CompileResult left, CompileResult right)
		{
			if (left.Result is ExcelErrorValue)
			{
				return new CompileResult(left.Result, DataType.ExcelError);
				//throw(new ExcelErrorValueException((ExcelErrorValue)left.Result));
			}
			else if (right.Result is ExcelErrorValue)
			{
				return new CompileResult(right.Result, DataType.ExcelError);
				//throw(new ExcelErrorValueException((ExcelErrorValue)right.Result));
			}
			return _implementation(left, right);
		}

		public override string ToString()
		{
			return "Operator: " + _operator;
		}

		private static IOperator _plus;
		public static IOperator Plus
		{
			get
			{
				return _plus ?? (_plus = new Operator(Operators.Plus, PrecedenceAddSubtract, (l, r) =>
				{
					l = l == null || l.Result == null ? new CompileResult(0, DataType.Integer) : l;
					r = r == null || r.Result == null ? new CompileResult(0, DataType.Integer) : r;
					ExcelErrorValue errorVal;
					if (EitherIsError(l, r, out errorVal))
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

		private static IOperator _minus;
		public static IOperator Minus
		{
			get
			{
				return _minus ?? (_minus = new Operator(Operators.Minus, PrecedenceAddSubtract, (l, r) =>
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

		private static IOperator _multiply;
		public static IOperator Multiply
		{
			get
			{
				return _multiply ?? (_multiply = new Operator(Operators.Multiply, PrecedenceMultiplyDevide, (l, r) =>
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

		private static IOperator _divide;
		public static IOperator Divide
		{
			get
			{
				return _divide ?? (_divide = new Operator(Operators.Divide, PrecedenceMultiplyDevide, (l, r) =>
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

		public static IOperator Exp
		{
			get
			{
				return new Operator(Operators.Exponentiation, PrecedenceExp, (l, r) =>
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
					 });
			}
		}

		public static IOperator Concat
		{
			get
			{
				return new Operator(Operators.Concat, PrecedenceConcat, (l, r) =>
					 {
						 l = l ?? new CompileResult(string.Empty, DataType.String);
						 r = r ?? new CompileResult(string.Empty, DataType.String);
						 var lStr = Convert.ToString(l.ResultValue);
						 var rStr = Convert.ToString(r.ResultValue);
						 return new CompileResult(string.Concat(lStr, rStr), DataType.String);
					 });
			}
		}

		private static IOperator _greaterThan;
		public static IOperator GreaterThan
		{
			get
			{
				return _greaterThan ??
						 (_greaterThan =
							  new Operator(Operators.GreaterThan, PrecedenceComparison,
									(l, r) => Compare(l, r, (compRes) => compRes > 0)));
			}
		}

		private static IOperator _eq;
		public static IOperator EqualsOperator
		{
			get
			{
				return _eq ??
						 (_eq =
							  new Operator(Operators.Equals, PrecedenceComparison,
									(l, r) => Compare(l, r, (compRes) => compRes == 0)));
			}
		}

		private static IOperator _notEqualsTo;
		public static IOperator NotEqualsTo
		{
			get
			{
				return _notEqualsTo ??
						 (_notEqualsTo =
							  new Operator(Operators.NotEqualTo, PrecedenceComparison,
									(l, r) => Compare(l, r, (compRes) => compRes != 0)));
			}
		}

		private static IOperator _greaterThanOrEqual;
		public static IOperator GreaterThanOrEqual
		{
			get
			{
				return _greaterThanOrEqual ??
						 (_greaterThanOrEqual =
							  new Operator(Operators.GreaterThanOrEqual, PrecedenceComparison,
									(l, r) => Compare(l, r, (compRes) => compRes >= 0)));
			}
		}

		private static IOperator _lessThan;
		public static IOperator LessThan
		{
			get
			{
				return _lessThan ??
						 (_lessThan =
							  new Operator(Operators.LessThan, PrecedenceComparison,
									(l, r) => Compare(l, r, (compRes) => compRes < 0)));
			}
		}

		public static IOperator LessThanOrEqual
		{
			get
			{
				//return new Operator(Operators.LessThanOrEqual, PrecedenceComparison, (l, r) => new CompileResult(Compare(l, r) <= 0, DataType.Boolean));
				return new Operator(Operators.LessThanOrEqual, PrecedenceComparison, (l, r) => Compare(l, r, (compRes) => compRes <= 0));
			}
		}

		private static IOperator _percent;
		public static IOperator Percent
		{
			get
			{
				if (_percent == null)
				{
					_percent = new Operator(Operators.Percent, PrecedencePercent, (l, r) =>
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
						 });
				}
				return _percent;
			}
		}

		private static object GetObjFromOther(CompileResult obj, CompileResult other)
		{
			if (obj.Result == null)
			{
				if (other.DataType == DataType.String) return string.Empty;
				else return 0d;
			}
			return obj.ResultValue;
		}

		private static CompileResult Compare(CompileResult l, CompileResult r, Func<int, bool> comparison)
		{
			ExcelErrorValue errorVal;
			if (EitherIsError(l, r, out errorVal))
			{
				return new CompileResult(errorVal);
			}
			return new CompileResult(comparison(Compare(l, r)), DataType.Boolean);
		}

		private static int Compare(CompileResult leftInput, CompileResult rightInput)
		{
			object left, right;
			left = GetObjFromOther(leftInput, rightInput);
			right = GetObjFromOther(rightInput, leftInput);
			var leftIsNumeric = ConvertUtil.IsNumeric(left) && !(left is bool);
			var rightIsNumeric = ConvertUtil.IsNumeric(right) && !(right is bool);

			if (leftIsNumeric && rightIsNumeric)
			{
				var leftNumber = ConvertUtil.GetValueDouble(left);
				var rightNumber = ConvertUtil.GetValueDouble(right);
				if (Math.Abs(leftNumber - rightNumber) < double.Epsilon)
				{
					return 0;
				}
				return leftNumber.CompareTo(rightNumber);
			}
			// Numbers are less than text are less than logical values: https://stackoverflow.com/questions/35050151/excel-if-statement-comparing-text-with-number
			// I can't find an MSDN source for this, but I've been testing it functionally and I can't find a counterexample.
			// If you find an MSDN source for this, please add a link here. 
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
				var comparisonResult = CompareString(left, right);
				return comparisonResult;
			}
			throw new InvalidOperationException($"Comparing operands of the given types {leftInput.DataType.ToString()} and {rightInput.DataType.ToString()} is not supported.");
		}

		private static int CompareString(object l, object r)
		{
			var sl = (l ?? "").ToString();
			var sr = (r ?? "").ToString();
			return System.String.Compare(sl, sr, System.StringComparison.OrdinalIgnoreCase);
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
	}
}
