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
			var rangeArgument = arguments.ElementAt(0).Value as ExcelDataProvider.IRangeInfo;
			//var criteriaArgument = (arguments.ElementAt(1).Value.ToString()).ToUpper();
			string criteriaString = null;
			if (arguments.ElementAt(1).Value is ExcelDataProvider.IRangeInfo criteriaRange && criteriaRange.IsMulti)
				criteriaString = null;
			else
				criteriaString = (GetFirstArgument(arguments.ElementAt(1)).ValueFirst.ToString()).ToUpper();

			if (arguments.Count() > 2)
			{
				var averageRangeArgument = arguments.ElementAt(2).Value as ExcelDataProvider.IRangeInfo;
				return calculateAverageUsingAverageRange(rangeArgument, criteriaString, averageRangeArgument);
			}
			else
			{
				return calculateAverageUsingRange(rangeArgument, criteriaString);
			}
		}

		private CompileResult calculateAverageUsingAverageRange(ExcelDataProvider.IRangeInfo cellsToCompare, string comparisonCriteria, ExcelDataProvider.IRangeInfo potentialCellsToAverage)
		{
			var sumOfValidValues = 0d;
			var numberOfValidValues = 0;
			foreach (var cell in cellsToCompare)
			{
				if (comparisonCriteria != null && objectMatchesCriteria(GetFirstArgument(cell.Value), comparisonCriteria))
				{
					var normalizedRow = cell.Row - cellsToCompare.Address._fromRow;
					var normalizedColumn = cell.Column - cellsToCompare.Address._fromCol;
					var avgRangeFromRow = potentialCellsToAverage.Address._fromRow;
					var avgRangeFromCol = potentialCellsToAverage.Address._fromCol;
					if (potentialCellsToAverage.Address._fromRow + normalizedRow <= potentialCellsToAverage.Address._toRow &&
						potentialCellsToAverage.Address._fromCol + normalizedColumn <= potentialCellsToAverage.Address._toCol)
					{
						var v = potentialCellsToAverage.GetOffset(normalizedRow, normalizedColumn);
						if (v is ExcelErrorValue cellError)
							return new CompileResult(cellError.Type);

						if (v is string || v is bool || v == null)
							continue;
						numberOfValidValues++;
						sumOfValidValues += ConvertUtil.GetValueDouble(v, true);
					}
				}
			}
			if (numberOfValidValues == 0)
				return new CompileResult(eErrorType.Div0);
			else
				return this.CreateResult(sumOfValidValues / numberOfValidValues, DataType.Decimal);
		}

		private CompileResult calculateAverageUsingRange(ExcelDataProvider.IRangeInfo potentialCellsToAverage, string comparisonCriteria)
		{
			var sumOfValidValues = 0d;
			var numberOfValidValues = 0;
			foreach (var cell in potentialCellsToAverage)
			{
				if (comparisonCriteria != null && IsNumericForAverageIf(cell.Value) &&
						objectMatchesCriteria(cell.Value, comparisonCriteria))
				{
					sumOfValidValues += cell.ValueDouble;
					numberOfValidValues++;
				}
				else if (cell.Value is ExcelErrorValue candidateError)
					return new CompileResult(candidateError.Type);
			}
			if (numberOfValidValues == 0)
				return new CompileResult(eErrorType.Div0);
			else
				return this.CreateResult(sumOfValidValues / numberOfValidValues, DataType.Decimal);
		}

		private bool IsNumericForAverageIf(object value)
		{
			// AverageIf does not consider booleans as valid numbers, whereas the normal IsNumeric method does.
			// Ensure that any boolean values that would be passed to IsNumeric() are caught instead.
			if (value is bool)
				return false;
			else
				return this.IsNumeric(value);
		}

		private bool objectMatchesCriteria(object objectToCompare, string criteria)
		{
			// given are the cell's value and the criteria parameter as a string.

			//if (criteria == null) possibly can delete this if condition
			//	return false;
			if (criteria.Equals(string.Empty))
				return (objectToCompare == null || objectToCompare.Equals(string.Empty));

			string operationFromCriteria = null;
			var replaceIndex = -1;
			if (Regex.IsMatch(criteria, @"^(<>|>=|<=){1}"))
			{
				operationFromCriteria = criteria.Substring(0, 2);
				replaceIndex = 2;
			}
			else if (Regex.IsMatch(criteria, @"^(=|<|>){1}"))
			{
				operationFromCriteria = criteria.Substring(0, 1);
				replaceIndex = 1;
			}
			else
				return isMatch(objectToCompare, criteria);

			var criteriaString = criteria.Substring(replaceIndex);
			//if (criteriaString.Equals(string.Empty))
			//	return (objectToCompare == null);
			IOperator operation;
			if (OperatorsDict.Instance.TryGetValue(operationFromCriteria, out operation))
			{
				// left check - objectToCompare
				// constant check - criteriaString
				// operator - operation.Operator
				if (objectToCompare == null && !(operation.Operator == OperatorType.Equals || operation.Operator == OperatorType.NotEqualTo))
					return false;
				switch (operation.Operator)
				{
					case OperatorType.Equals:
						return this.isMatch(objectToCompare, criteriaString);
					case OperatorType.NotEqualTo:

						return (!isMatch(objectToCompare, criteriaString));
						//return (!(criteriaString.Equals(objectToCompare)));
					case OperatorType.GreaterThan:
					case OperatorType.GreaterThanOrEqual:
					case OperatorType.LessThan:
					case OperatorType.LessThanOrEqual:
						return this.compareAsInequalityExpression(objectToCompare, criteriaString, operation.Operator);
					default:
						return this.isMatch(objectToCompare, criteriaString);
				}
			}

			return false;
		}

		// criteria is either a number, boolean, or string. The string can contain a date/time, or 
		// text string that may require wildcard Regex.
		private bool isMatch(object objectToCompare, string criteria)
		{
			if (criteria == null || criteria.Equals(string.Empty))
			{
				return (objectToCompare == null);
			}
			string objectAsString = null;
			var handleAsBool = (criteria.Equals(Boolean.TrueString.ToUpper()) || criteria.Equals(Boolean.FalseString.ToUpper()));
			var handleAsDate = ConvertUtil.TryParseDateObjectToOADate(criteria, out double criteriaAsOADate);
			if (handleAsBool)
			{
				if (objectToCompare is bool objectBool)
				{
					return (criteria.Equals(objectBool.ToString().ToUpper()));
				}
				else
					return false;
			}
			if (handleAsDate)
			{
				criteria = criteriaAsOADate.ToString();
			}
			if (objectToCompare == null)
				return false;
			else if (objectToCompare is bool objectAsBool)
				objectAsString = objectAsBool.ToString();
			else if (ConvertUtil.TryParseDateObjectToOADate(objectToCompare, out double objectAsOADate))
				objectAsString = objectAsOADate.ToString();
			else if (objectToCompare is System.DateTime objectAsDate)
				objectAsString = (objectAsDate.ToOADate()).ToString();
			else
				objectAsString = objectToCompare.ToString();

			objectAsString = objectAsString.ToUpper();

			if (criteria.Contains("*") || criteria.Contains("?"))
			{
				var regexPattern = Regex.Escape(criteria);
				regexPattern = string.Format("^{0}$", regexPattern);
				regexPattern = regexPattern.Replace(@"\*", ".*");
				regexPattern = regexPattern.Replace("~.*", "\\*");
				regexPattern = regexPattern.Replace(@"\?", ".");
				regexPattern = regexPattern.Replace("~.", "\\?");
				if (Regex.IsMatch(objectAsString, regexPattern))
				{
					return true;
				}
			}
			return (criteria.CompareTo(objectAsString) == 0);
		}

		private bool compareAsInequalityExpression(object objectToCompare, string criteria, OperatorType comparisonOperator)
		{
			var handleAsBool = (criteria.Equals(Boolean.TrueString.ToUpper()) || criteria.Equals(Boolean.FalseString.ToUpper()));
			if (ConvertUtil.TryParseDateObjectToOADate(criteria, out double criteriaNumber))
			{
				if (IsNumericForAverageIf(objectToCompare))
				{
					var numberToCompare = ConvertUtil.GetValueDouble(objectToCompare);
					switch (comparisonOperator)
					{
						case OperatorType.LessThan:
							return (numberToCompare < criteriaNumber);
						case OperatorType.LessThanOrEqual:
							return (numberToCompare <= criteriaNumber);
						case OperatorType.GreaterThan:
							return (numberToCompare > criteriaNumber);
						case OperatorType.GreaterThanOrEqual:
							return (numberToCompare >= criteriaNumber);
					}
				}
				else
					return false;
			}
			else
			{
				//if (IsNumericForAverageIf(objectToCompare))
				//	return false;
				var comparisonResult = (objectToCompare.ToString().ToUpper()).CompareTo(criteria);
				if (handleAsBool)
				{
					if (objectToCompare is bool objectBool)
						comparisonResult = (objectToCompare.ToString().ToUpper()).CompareTo(criteria);
					else
						return false;
				}
				else if (IsNumeric(objectToCompare))
					return false;
				
				switch (comparisonOperator)
				{
					case OperatorType.LessThan:
						return (comparisonResult == -1);
					case OperatorType.LessThanOrEqual:
						return (comparisonResult == -1 || comparisonResult == 0);
					case OperatorType.GreaterThan:
						return (comparisonResult == 1);
					case OperatorType.GreaterThanOrEqual:
						return (comparisonResult == 1 || comparisonResult == 0);
					default:
						return false;
				}
			}
			return false;
		}

		public CompileResult aExecute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
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
				if (criteria != null && Evaluate(GetFirstArgument(cell.Value), criteria))
				//if (criteria != null && objectMatchesCriteria(GetFirstArgument(cell.Value), criteria))
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
			var retVal = 0d;
			var nMatches = 0;
			foreach (var candidate in range)
			{
				if (expression != null && IsNumericForAverageIf(GetFirstArgument(candidate.Value)) &&
					Evaluate(GetFirstArgument(candidate.Value), expression))
				{
					retVal += candidate.ValueDouble;
					nMatches++;
				}
			}
			return Divide(retVal, nMatches);
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
	}
}
