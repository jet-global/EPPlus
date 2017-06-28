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
			string criteriaString = null;
			if (arguments.ElementAt(1).Value is ExcelDataProvider.IRangeInfo criteriaRange && criteriaRange.IsMulti)
				return new CompileResult(eErrorType.Div0);
			else
				criteriaString = this.GetFirstArgument(arguments.ElementAt(1)).ValueFirst.ToString().ToUpper();
			if (arguments.Count() > 2)
			{
				var averageRangeArgument = arguments.ElementAt(2).Value as ExcelDataProvider.IRangeInfo;
				return this.calculateAverageUsingAverageRange(rangeArgument, criteriaString, averageRangeArgument);
			}
			else
				return this.calculateAverageUsingRange(rangeArgument, criteriaString);
		}

		private CompileResult calculateAverageUsingAverageRange(ExcelDataProvider.IRangeInfo cellsToCompare, string comparisonCriteria, ExcelDataProvider.IRangeInfo potentialCellsToAverage)
		{
			var sumOfValidValues = 0d;
			var numberOfValidValues = 0;
			foreach (var cell in cellsToCompare)
			{
				if (comparisonCriteria != null && this.objectMatchesCriteria(this.GetFirstArgument(cell.Value), comparisonCriteria))
				{
					var relativeRow = cell.Row - cellsToCompare.Address._fromRow;
					var relativeColumn = cell.Column - cellsToCompare.Address._fromCol;
					if (potentialCellsToAverage.Address._fromRow + relativeRow <= potentialCellsToAverage.Address._toRow &&
						potentialCellsToAverage.Address._fromCol + relativeColumn <= potentialCellsToAverage.Address._toCol)
					{
						var valueOfCellToAverage = potentialCellsToAverage.GetOffset(relativeRow, relativeColumn);
						if (valueOfCellToAverage is ExcelErrorValue cellError)
							return new CompileResult(cellError.Type);
						if (valueOfCellToAverage is string || valueOfCellToAverage is bool || valueOfCellToAverage == null)
							continue;
						sumOfValidValues += ConvertUtil.GetValueDouble(valueOfCellToAverage, true);
						numberOfValidValues++;
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
				if (comparisonCriteria != null && this.IsNumericForAverageIf(this.GetFirstArgument(cell.Value)) &&
						this.objectMatchesCriteria(this.GetFirstArgument(cell.Value), comparisonCriteria))
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

		// This method should be abstracted out of the AVERAGEIF function.
		private bool objectMatchesCriteria(object objectToCompare, string criteria)
		{
			var operatorIndex = -1;
			// Check if the criteria is an expression; i.e. begins with the operators <>, =, >, >=, <, or <=
			if (Regex.IsMatch(criteria, @"^(<>|>=|<=){1}"))
				operatorIndex = 2;
			else if (Regex.IsMatch(criteria, @"^(=|<|>){1}"))
				operatorIndex = 1;
			// If the criteria is an expression, evaluate as such.
			if (operatorIndex != -1)
			{
				var expressionOperatorString = criteria.Substring(0, operatorIndex);
				var criteriaString = criteria.Substring(operatorIndex);
				IOperator expressionOperator;
				if (OperatorsDict.Instance.TryGetValue(expressionOperatorString, out expressionOperator))
				{
					switch (expressionOperator.Operator)
					{
						case OperatorType.Equals:
							return this.isMatch(objectToCompare, criteriaString, true);
						case OperatorType.NotEqualTo:
							return !this.isMatch(objectToCompare, criteriaString, true);
						case OperatorType.GreaterThan:
						case OperatorType.GreaterThanOrEqual:
						case OperatorType.LessThan:
						case OperatorType.LessThanOrEqual:
							return this.compareAsInequalityExpression(objectToCompare, criteriaString, expressionOperator.Operator);
						default:
							return this.isMatch(objectToCompare, criteriaString);
					}
				}
				else
					return false;
			}
			else
				return this.isMatch(objectToCompare, criteria);
		}

		/// <summary>
		/// criteria is either a number, boolean, or string. The string can contain a date/time, or
		/// text string that may require wildcard Regex.
		/// </summary>
		/// <param name="objectToCompare"></param>
		/// <param name="criteria"></param>
		/// <param name="matchAsEqualityExpression"></param>
		/// <returns></returns>
		private bool isMatch(object objectToCompare, string criteria, bool matchAsEqualityExpression = false)
		{
			// Equality related expression evaluation (= or <>) only considers empty cells as equal to empty string criteria.
			if (criteria.Equals(string.Empty))
				return ((matchAsEqualityExpression) ? objectToCompare == null : objectToCompare == null || objectToCompare.Equals(string.Empty));
			var criteriaIsBool = criteria.Equals(Boolean.TrueString.ToUpper()) || criteria.Equals(Boolean.FalseString.ToUpper());
			if (ConvertUtil.TryParseDateObjectToOADate(criteria, out double criteriaAsOADate))
				criteria = criteriaAsOADate.ToString();
			string objectAsString = null;
			if (objectToCompare == null)
				return false;
			else if (objectToCompare is bool && criteriaIsBool)
				return criteria.Equals(objectToCompare.ToString().ToUpper());
			else if (objectToCompare is bool ^ criteriaIsBool)
				return false;
			else if (ConvertUtil.TryParseDateObjectToOADate(objectToCompare, out double objectAsOADate))
				objectAsString = objectAsOADate.ToString();
			else
				objectAsString = objectToCompare.ToString().ToUpper();
			if (criteria.Contains("*") || criteria.Contains("?"))
			{
				var regexPattern = Regex.Escape(criteria);
				regexPattern = string.Format("^{0}$", regexPattern);
				regexPattern = regexPattern.Replace(@"\*", ".*");
				regexPattern = regexPattern.Replace("~.*", "\\*");
				regexPattern = regexPattern.Replace(@"\?", ".");
				regexPattern = regexPattern.Replace("~.", "\\?");
				return Regex.IsMatch(objectAsString, regexPattern);
			}
			else
				return (criteria.CompareTo(objectAsString) == 0);
		}

		private bool compareAsInequalityExpression(object objectToCompare, string criteria, OperatorType comparisonOperator)
		{
			if (objectToCompare == null)
				return false;
			var comparisonResult = int.MinValue;
			if (ConvertUtil.TryParseDateObjectToOADate(criteria, out double criteriaNumber)) // Handle the criteria as a number/date.
			{
				if (this.IsNumericForAverageIf(objectToCompare))
				{
					var numberToCompare = ConvertUtil.GetValueDouble(objectToCompare);
					comparisonResult = numberToCompare.CompareTo(criteriaNumber);
				}
				else
					return false;
			}
			else // Handle the criteria as a non-numeric, non-date text value.
			{
				if (criteria.Equals(Boolean.TrueString.ToUpper()) || criteria.Equals(Boolean.FalseString.ToUpper()))
				{
					if (!(objectToCompare is bool objectBool))
						return false;
				}
				else if (this.IsNumeric(objectToCompare))
					return false;
				comparisonResult = (objectToCompare.ToString().ToUpper()).CompareTo(criteria);
			}
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
	}
}
