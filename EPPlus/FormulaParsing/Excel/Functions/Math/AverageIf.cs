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
using System.Text.RegularExpressions;
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Utilities;
using OfficeOpenXml.Utils;
using Require = OfficeOpenXml.FormulaParsing.Utilities.Require;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	public class AverageIf : HiddenValuesHandlingFunction
	{
		private readonly ExpressionEvaluator _expressionEvaluator;

		public AverageIf()
			 : this(new ExpressionEvaluator())
		{

		}

		public AverageIf(ExpressionEvaluator evaluator)
		{
			Require.That(evaluator).Named("evaluator").IsNotNull();
			_expressionEvaluator = evaluator;
		}

		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentCountIsValid(arguments, 2) == false)
				return new CompileResult(eErrorType.Value);
			var args = arguments.ElementAt(0).Value as ExcelDataProvider.IRangeInfo;
			//var criteriaCandidate = arguments.ElementAt(1).Value;
			var criteria = GetFirstArgument(arguments.ElementAt(1)).ValueFirst != null ? GetFirstArgument(arguments.ElementAt(1)).ValueFirst.ToString() : string.Empty;
			//var criteria = criteriaCandidate != null ? criteriaCandidate.ToString() : string.Empty;
			var retVal = 0d;
			if (args == null)
			{
				var val = GetFirstArgument(arguments.ElementAt(0)).Value;
				if (criteria != null && Evaluate(val, criteria))
				{
					if (arguments.Count() > 2)
					{
						var averageVal = arguments.ElementAt(2).Value;
						var averageRange = averageVal as ExcelDataProvider.IRangeInfo;
						if (averageRange != null)
						{
							retVal = averageRange.First().ValueDouble;
						}
						else
						{
							retVal = ConvertUtil.GetValueDouble(averageVal, true);
						}
					}
					else
					{
						retVal = ConvertUtil.GetValueDouble(val, true);
					}
				}
				else
				{
					throw new ExcelErrorValueException(eErrorType.Div0);
				}
			}
			else if (arguments.Count() > 2)
			{
				var lookupRange = arguments.ElementAt(2).Value as ExcelDataProvider.IRangeInfo;
				retVal = CalculateWithAverageRange(args, criteria, lookupRange);
			}
			else
			{
				retVal = CalculateSingleRange(args, criteria);
			}
			return CreateResult(retVal, DataType.Decimal);
		}

		private double CalculateWithAverageRange(ExcelDataProvider.IRangeInfo range, string criteria, ExcelDataProvider.IRangeInfo sumRange)
		{
			var retVal = 0d;
			var nMatches = 0;
			foreach (var cell in range)
			{
				//if (criteria != null && Evaluate(GetFirstArgument(cell.Value), criteria))
				if (criteria != null && objectMatchesCriteria(GetFirstArgument(cell.Value), criteria))
				{
					var or = cell.Row - range.Address._fromRow;
					var oc = cell.Column - range.Address._fromCol;
					var avgRangeFromRow = sumRange.Address._fromRow;
					var avgRangeFromCol = sumRange.Address._fromCol;
					if (sumRange.Address._fromRow + or <= sumRange.Address._toRow &&
						sumRange.Address._fromCol + oc <= sumRange.Address._toCol)
					{
						var v = sumRange.GetOffset(or, oc);
						nMatches++;
						retVal += ConvertUtil.GetValueDouble(v, true);
					}
				}
			}
			return Divide(retVal, nMatches);
		}

		private double CalculateSingleRange(ExcelDataProvider.IRangeInfo range, string expression)
		{
			var sumOfValidValues = 0d;
			var numberOfValidValues = 0;
			foreach (var candidate in range)
			{
				if (expression != null && IsNumericForAverageIf(candidate.Value) &&
						Evaluate(candidate.Value, expression))
				{
					sumOfValidValues += candidate.ValueDouble;
					numberOfValidValues++;
				}
				else if (candidate.Value is ExcelErrorValue candidateError)
					throw new ExcelErrorValueException(candidateError.Type);
			}
			return this.Divide(sumOfValidValues, numberOfValidValues);
			//var retVal = 0d;
			//var nMatches = 0;
			//foreach (var candidate in range)
			//{
			//	if (expression != null && IsNumericForAverageIf(GetFirstArgument(candidate.Value)) && 
			//		Evaluate(GetFirstArgument(candidate.Value), expression))
			//	{
			//		retVal += candidate.ValueDouble;
			//		nMatches++;
			//	}
			//}
			//return Divide(retVal, nMatches);
		}


		private bool Evaluate(object obj, string expression)
		{
			if (IsNumeric(obj))
			{
				double? candidate = ConvertUtil.GetValueDouble(obj);
				if (candidate.HasValue)
				{
					return _expressionEvaluator.Evaluate(candidate.Value, expression);
				}
			}
			return _expressionEvaluator.Evaluate(obj, expression);
		}

		private bool IsNumericForAverageIf(object value)
		{
			if (value is bool)
				return false;
			else
				return this.IsNumeric(value);
			//return (value.GetType().IsPrimitive || value is decimal || value is System.DateTime || value is TimeSpan);

		}

		// return true if object matches criteria.
		private bool objectMatchesCriteria(object objectToCompare, string criteria)
		{
			if (objectToCompare.Equals(string.Empty))
				return (objectToCompare == null || objectToCompare.Equals(string.Empty));
			
			string operationFromCriteria = null;
			if (Regex.IsMatch(criteria, @"^[^a-zA-Z0-9]{2}"))
				operationFromCriteria = criteria.Substring(0, 2);
			else if (Regex.IsMatch(criteria, @"^[^a-zA-Z0-9]{1}"))
				operationFromCriteria = criteria.Substring(0, 1);
			else
				return compareAsValue(objectToCompare, criteria);

			// Criteria is an expression.
			var criteriaObject = criteria.Replace(operationFromCriteria, string.Empty);
			IOperator operation;
			if (OperatorsDict.Instance.TryGetValue(operationFromCriteria, out operation))
			{
				switch (operation.Operator)
				{
					case OperatorType.Equals:
						break;
					case OperatorType.GreaterThan:
						break;
					case OperatorType.GreaterThanOrEqual:
						break;
					case OperatorType.LessThan:
						break;
					case OperatorType.LessThanOrEqual:
						break;
					case OperatorType.NotEqualTo:
						break;
					default:
						break;
				}
			}

			return false;
		}

		private bool compareAsValue(object objectToCompare, string criteria)
		{
			string objectAsString = null;
			if (objectToCompare is bool objectAsBool)
				objectAsString = objectAsBool.ToString().ToUpper();
			else
				objectAsString = objectToCompare.ToString();

			if (criteria.Contains("*") || criteria.Contains("?"))
			{
				var regexPattern = Regex.Escape(criteria);
				regexPattern = string.Format("^{0}$", regexPattern);
				regexPattern = regexPattern.Replace(@"\*", ".*");
				regexPattern = regexPattern.Replace(@"\?", ".");
				if (Regex.IsMatch(objectAsString, regexPattern))
				{
					return true;
				}
			}
			return (criteria.CompareTo(objectAsString) == 0);
		}

		private bool compareAsExpression(object objectToCompare, string expression)
		{
			return false;
		}
	}
}
