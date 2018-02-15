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
* For code change notes, see the source control history.
*******************************************************************************/
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	/// <summary>
	/// This class provides a criteria comparison function to use in any Excel functions
	/// that require comparing cell values against a specific criteria.
	/// This class is currently used in AverageIf.cs, AverageIfs.cs, SumIf.cs, SumIfs.cs, and CountIf.cs.
	/// </summary>
	public static class IfHelper
	{
		private enum ComparisonValue
		{
			LessThan,
			Equal,
			GreaterThan,
			NotEqual
		}

		#region IfHelper Public Static Functions
		/// <summary>
		/// Compares the given <paramref name="testObject"/> against the given <paramref name="rawCriterionObject"/>.
		/// This method is expected to be used with any of the *IF or *IFS Excel functions (ex: the AVERAGEIF function)
		/// for comparing an object against a criterion. See the documentation for the any of the *IF or *IFS functions
		/// for information on acceptable forms of the criterion.
		/// </summary>
		/// <param name="testObject">The object to compare against the given <paramref name="rawCriterionObject"/>.</param>
		/// <param name="rawCriterionObject">The criterion value or expression that dictates whether the given <paramref name="testObject"/> passes or fails.</param>
		/// <returns>Returns true if <paramref name="testObject"/> matches the <paramref name="rawCriterionObject"/>.</returns>
		public static bool ObjectMatchesCriterion(object testObject, object rawCriterionObject)
		{
			object criterion = rawCriterionObject;
			OperatorType criterionOperator = OperatorType.Equals;
			bool criterionIsExpression = false;
			if (rawCriterionObject is string rawCriterionString)
			{
				string criterionString;
				if (IfHelper.TryParseCriterionAsExpression(rawCriterionString, out IOperator expressionOperator, out string expressionCriterion))
				{
					criterionOperator = expressionOperator.Operator;
					criterionString = expressionCriterion;
					criterionIsExpression = true;
				}
				else
					criterionString = rawCriterionString.ToUpper(CultureInfo.CurrentCulture);

				if (IfHelper.TryParseCriterionStringToObject(criterionString, out object criterionObject))
					criterion = criterionObject;
				else
					criterion = criterionString;
			}
			return IfHelper.IsMatch(testObject, criterionOperator, criterion, criterionIsExpression);
		}

		/// <summary>
		/// Returns a list containing the indices of the cells in <paramref name="cellsToCompare"/> that satisfy
		/// the criterion detailed in <paramref name="criterion"/>.
		/// </summary>
		/// <param name="cellsToCompare">The <see cref="ExcelDataProvider.IRangeInfo"/> containing the cells to test against the <paramref name="criterion"/>.</param>
		/// <param name="criterion">The criterion dictating the acceptable contents of a given cell.</param>
		/// <returns>Returns a list of indexes corresponding to each cell that satisfies the given criterion.</returns>
		public static List<int> GetIndicesOfCellsPassingCriterion(ExcelDataProvider.IRangeInfo cellsToCompare, object criterion)
		{
			var passingIndices = new List<int>();
			var cellValues = cellsToCompare.AllValues();
			for (var index = 0; index < cellValues.Count(); index++)
			{
				if (IfHelper.ObjectMatchesCriterion(cellValues.ElementAt(index), criterion))
					passingIndices.Add(index);
			}
			return passingIndices;
		}

		/// <summary>
		/// Checks if the <paramref name="expectedRange"/> is the same width and height as the
		/// <paramref name="actualRange"/>.
		/// </summary>
		/// <param name="expectedRange">The <see cref="ExcelDataProvider.IRangeInfo"/> with the desired cell width and height.</param>
		/// <param name="actualRange">The <see cref="ExcelDataProvider.IRangeInfo"/> with the width and height to be tested.</param>
		/// <returns>Returns true if <paramref name="expectedRange"/> and <paramref name="actualRange"/> have the same width and height values.</returns>
		public static bool RangesAreTheSameShape(ExcelDataProvider.IRangeInfo expectedRange, ExcelDataProvider.IRangeInfo actualRange)
		{
			var expectedRangeWidth = expectedRange.Address._toCol - expectedRange.Address._fromCol;
			var expectedRangeHeight = expectedRange.Address._toRow - expectedRange.Address._fromRow;
			var actualRangeWidth = actualRange.Address._toCol - actualRange.Address._fromCol;
			var actualRangeHeight = actualRange.Address._toRow - actualRange.Address._fromRow;
			return (expectedRangeWidth == actualRangeWidth && expectedRangeHeight == actualRangeHeight);
		}

		/// <summary>
		/// Ensures that the given <paramref name="criterionCandidate"/> is of a form that can be
		/// represented as a criterion.
		/// </summary>
		/// <param name="criterionCandidate">The <see cref="FunctionArgument"/> containing the criterion.</param>
		/// <param name="context">The context from the function calling this function.</param>
		/// <returns>Returns the criterion in <paramref name="criterionCandidate"/> as an object.</returns>
		public static object ExtractCriterionObject(FunctionArgument criterionCandidate, ParsingContext context)
		{
			object criterionObject = null;
			if (criterionCandidate.Value is ExcelDataProvider.IRangeInfo criterionRange)
			{
				if (criterionRange.IsMulti)
				{
					var worksheet = context.ExcelDataProvider.GetRange(context.Scopes.Current.Address.Worksheet, 1, 1, "A1").Worksheet;
					var functionRow = context.Scopes.Current.Address.FromRow;
					var functionColumn = context.Scopes.Current.Address.FromCol;
					criterionObject = IfHelper.ExtractCriterionFromCellRange(criterionCandidate, worksheet, functionRow, functionColumn);
				}
				else
				{
					criterionObject = criterionCandidate.ValueFirst;
					if (criterionObject is List<object> objectList)
						criterionObject = objectList.First();
				}
			}
			else if (criterionCandidate.Value is List<FunctionArgument> argumentList)
				criterionObject = argumentList.First().ValueFirst;
			else
				criterionObject = criterionCandidate.ValueFirst;

			// Note that Excel considers null criterion equivalent to a criterion of 0.
			if (criterionObject == null)
				criterionObject = 0;
			return criterionObject;
		}

		/// <summary>
		/// Takes a cell range and converts it into a single value criterion.
		/// </summary>
		/// <param name="criteriaArgument">The cell range that will be reduced to a single value criteria.</param>
		/// <param name="worksheet">The current worksheet that is being used.</param>
		/// <param name="rowLocation">The row location of the cell of the calling function.</param>
		/// <param name="colLocation">The column location of the cell of the calling function.</param>
		/// <returns>Returns the value of the cell in the given cell range that corresponds to the position of the calling function.</returns>
		public static object ExtractCriterionFromCellRange(FunctionArgument criteriaArgument, ExcelWorksheet worksheet, int rowLocation, int colLocation)
		{
			if (criteriaArgument.Value == null)
				return 0;
			if (criteriaArgument.Value is ExcelErrorValue)
			{
				if (worksheet == null)
					return 0;
			}
			if (rowLocation <= 0 || colLocation <= 0)
				return 0;

			var criterionArgument = criteriaArgument.ValueAsRangeInfo.Address;

			if (criterionArgument.Rows > criterionArgument.Columns)
			{
				var currentAddressRow = rowLocation;
				var startRow = criterionArgument.Start.Row;
				var endRow = criterionArgument.End.Row;

				if (currentAddressRow == startRow)
				{
					var cellColumn = criterionArgument.Start.Column;
					return worksheet.Cells[startRow, cellColumn].Value;
				}
				else if (currentAddressRow == endRow)
				{
					var cellColumn = criterionArgument.Start.Column;
					return worksheet.Cells[endRow, cellColumn].Value;
				}
				else if (currentAddressRow > startRow && currentAddressRow < endRow)
				{
					var cellColumn = criterionArgument.Start.Column;
					return worksheet.Cells[currentAddressRow, cellColumn].Value;
				}
				else
					return 0;
			}
			else if (criterionArgument.Rows < criterionArgument.Columns)
			{
				var currentAddressCol = colLocation;
				var startCol = criterionArgument.Start.Column;
				var endCol = criterionArgument.End.Column;

				if (currentAddressCol == startCol)
				{
					var cellRow = criterionArgument.Start.Row;
					return worksheet.Cells[cellRow, currentAddressCol].Value;
				}
				else if (currentAddressCol == endCol)
				{
					var cellRow = criterionArgument.Start.Row;
					return worksheet.Cells[cellRow, currentAddressCol].Value;
				}
				else if (currentAddressCol > startCol && currentAddressCol < endCol)
				{
					var cellRow = criterionArgument.Start.Row;
					return worksheet.Cells[cellRow, currentAddressCol].Value;
				}
				else
					return 0;
			}
			else
				return 0;
		}
		#endregion

		#region IfHelper Private Static Functions
		/// <summary>
		/// Compare the <paramref name="testObject"/> against the <paramref name="criterionObject"/> using the
		/// <paramref name="criterionOperation"/> to determine what qualifies as a match.
		/// </summary>
		/// <param name="testObject">The object to compare against the <paramref name="criterionObject"/>.</param>
		/// <param name="criterionOperation">The comparison operation to perform between the <paramref name="testObject"/> and the <paramref name="criterionObject"/>.</param>
		/// <param name="criterionObject">The criterion value that determines what value the <paramref name="testObject"/> should be compared against.</param>
		/// <param name="matchCriterionAsExpression">
		///		Optional parameter that indicates whether the given criterion was explicitly parsed as an expression,
		///		or was implied to use the Equals operator.</param>
		/// <returns>Returns true if the <paramref name="testObject"/> matches the <paramref name="criterionObject"/> for the given <paramref name="criterionOperation"/>, and false otherwise.</returns>
		private static bool IsMatch(object testObject, OperatorType criterionOperation, object criterionObject, bool matchCriterionAsExpression = false)
		{
			// Excel compares some data types differently if the operator is equality based (=/<>), as opposed to inequality based (</<=/>/>=).
			bool compareForEquality = (criterionOperation == OperatorType.Equals || criterionOperation == OperatorType.NotEqualTo);
			ComparisonValue compareResult = ComparisonValue.NotEqual;
			if (IfHelper.TryCompareAsStrings(testObject, criterionObject, compareForEquality, matchCriterionAsExpression, out ComparisonValue stringResult))
				compareResult = stringResult;
			else if (IfHelper.TryCompareAsBooleans(testObject, criterionObject, out ComparisonValue boolResult))
				compareResult = boolResult;
			else if (IfHelper.TryCompareAsDates(testObject, criterionObject, compareForEquality, out ComparisonValue dateResult))
				compareResult = dateResult;
			else if (IfHelper.TryCompareAsNumericValues(testObject, criterionObject, compareForEquality, out ComparisonValue numberResult))
				compareResult = numberResult;
			else if (IfHelper.TryCompareAsErrorValues(testObject, criterionObject, out ComparisonValue errorResult))
				compareResult = errorResult;

			switch (criterionOperation)
			{
				case OperatorType.Equals:
					return (compareResult == ComparisonValue.Equal);
				case OperatorType.NotEqualTo:
					return (compareResult != ComparisonValue.Equal);
				case OperatorType.LessThan:
					return (compareResult == ComparisonValue.LessThan);
				case OperatorType.LessThanOrEqual:
					return (compareResult == ComparisonValue.Equal || compareResult == ComparisonValue.LessThan);
				case OperatorType.GreaterThan:
					return (compareResult == ComparisonValue.GreaterThan);
				case OperatorType.GreaterThanOrEqual:
					return (compareResult == ComparisonValue.Equal || compareResult == ComparisonValue.GreaterThan);
				default:
					throw new InvalidOperationException("The criterionOperation is an invalid operator type for this function.");
			}
		}

		/// <summary>
		/// Try to compare the objects as string values.
		/// </summary>
		/// <param name="testObject">The object being compared.</param>
		/// <param name="criterionObject">The criterion used to judge the <paramref name="testObject"/>.</param>
		/// <param name="compareForEquality">Indicates how the two objects should be compared.</param>
		/// <param name="matchCriterionAsExpression">Indicates what is considered a match in the case where the criterion is an empty string.</param>
		/// <param name="result">
		///		Indicates how <paramref name="testObject"/> compares to the <paramref name="criterionObject"/>.
		///		Do not use this value if this function returns false.</param>
		/// <returns>Returns true if <paramref name="testObject"/> and <paramref name="criterionObject"/> are able to be compared, and false otherwise.</returns>
		private static bool TryCompareAsStrings(object testObject, object criterionObject, bool compareForEquality, bool matchCriterionAsExpression, out ComparisonValue result)
		{
			result = ComparisonValue.NotEqual;
			if (criterionObject is string criterionString)
			{
				if (criterionString.Equals(string.Empty))
				{
					if (matchCriterionAsExpression)
						result = (testObject == null) ? ComparisonValue.Equal : ComparisonValue.NotEqual;
					else
						result = (testObject == null || testObject.Equals(string.Empty)) ? ComparisonValue.Equal : ComparisonValue.NotEqual;
				}
				else if (testObject is string testString)
					result = IfHelper.CompareAsStrings(testString, criterionString, compareForEquality);
				else
					return false;
			}
			else
				return false;
			return true;
		}

		/// <summary>
		/// Try to compare the objects as booleans.
		/// </summary>
		/// <param name="testObject">The object being compared.</param>
		/// <param name="criterionObject">The criterion used to judge the <paramref name="testObject"/>.</param>
		/// <param name="result">
		///		Indicates how <paramref name="testObject"/> compares to the <paramref name="criterionObject"/>.
		///		Do not use this value if this function returns false.</param>
		/// <returns>Returns true if <paramref name="testObject"/> and <paramref name="criterionObject"/> are able to be compared, and false otherwise.</returns>
		private static bool TryCompareAsBooleans(object testObject, object criterionObject, out ComparisonValue result)
		{
			result = ComparisonValue.NotEqual;
			if (criterionObject is bool criterionBool && testObject is bool testBool)
				result = IfHelper.ConvertIntToComparisonValue(testBool.CompareTo(criterionBool));
			else
				return false;
			return true;
		}

		/// <summary>
		/// Try to compare the objects as dates.
		/// </summary>
		/// <param name="testObject">The object being compared.</param>
		/// <param name="criterionObject">The criterion used to judge the <paramref name="testObject"/>.</param>
		/// <param name="compareForEquality">Indicates how the two objects should be compared.</param>
		/// <param name="result">
		///		Indicates how <paramref name="testObject"/> compares to the <paramref name="criterionObject"/>.
		///		Do not use this value if this function returns false.</param>
		/// <returns>Returns true if <paramref name="testObject"/> and <paramref name="criterionObject"/> are able to be compared, and false otherwise.</returns>
		private static bool TryCompareAsDates(object testObject, object criterionObject, bool compareForEquality, out ComparisonValue result)
		{
			result = ComparisonValue.NotEqual;
			if (criterionObject is System.DateTime criterionDate)
			{
				if (IfHelper.TryConvertObjectToDouble(testObject, out double testDateDouble))
					result = IfHelper.ConvertIntToComparisonValue(testDateDouble.CompareTo(criterionDate.ToOADate()));
				else if (IfHelper.TryConvertObjectToDate(testObject, out System.DateTime testDate) && compareForEquality)
					result = IfHelper.ConvertIntToComparisonValue(System.DateTime.Compare(testDate, criterionDate));
				else
					return false;
			}
			else
				return false;
			return true;
		}

		/// <summary>
		/// Try to compare the objects as numbers.
		/// </summary>
		/// <param name="testObject">The object being compared.</param>
		/// <param name="criterionObject">The criterion used to judge the <paramref name="testObject"/>.</param>
		/// <param name="compareForEquality">Indicates how the two objects should be compared.</param>
		/// <param name="result">
		///		Indicates how <paramref name="testObject"/> compares to the <paramref name="criterionObject"/>.
		///		Do not use this value if this function returns false.</param>
		/// <returns>Returns true if <paramref name="testObject"/> and <paramref name="criterionObject"/> are able to be compared, and false otherwise.</returns>
		private static bool TryCompareAsNumericValues(object testObject, object criterionObject, bool compareForEquality, out ComparisonValue result)
		{
			result = ComparisonValue.NotEqual;
			if (ConvertUtil.IsNumeric(criterionObject, true))
			{
				if (IfHelper.TryConvertObjectToDouble(testObject, out double testDouble, compareForEquality))
				{
					var criterionDouble = ConvertUtil.GetValueDouble(criterionObject, true);
					result = IfHelper.ConvertIntToComparisonValue(testDouble.CompareTo(criterionDouble));
				}
				else
					return false;
			}
			else
				return false;
			return true;
		}

		/// <summary>
		/// Try to compare the objects as Excel error values.
		/// </summary>
		/// <param name="testObject">The object being compared.</param>
		/// <param name="criterionObject">The criterion used to judge the <paramref name="testObject"/>.</param>
		/// <param name="result">
		///		Indicates how <paramref name="testObject"/> compares to the <paramref name="criterionObject"/>.
		///		Do not use this value if this function returns false.</param>
		/// <returns>Returns true if <paramref name="testObject"/> and <paramref name="criterionObject"/> are able to be compared, and false otherwise.</returns>
		private static bool TryCompareAsErrorValues(object testObject, object criterionObject, out ComparisonValue result)
		{
			result = ComparisonValue.NotEqual;
			if (criterionObject is ExcelErrorValue criterionErrorValue && testObject is ExcelErrorValue testErrorValue)
				result = IfHelper.CompareAsErrorValues(testErrorValue.Type, criterionErrorValue.Type);
			else
				return false;
			return true;
		}

		private static ComparisonValue CompareAsErrorValues(eErrorType testErrorValue, eErrorType criterionErrorValue)
		{
			var testErrorNumericValue = IfHelper.GetErrorNumericValue(testErrorValue);
			var criterionErrorNumericValue = IfHelper.GetErrorNumericValue(criterionErrorValue);
			if (testErrorNumericValue < criterionErrorNumericValue)
				return ComparisonValue.LessThan;
			else if (testErrorNumericValue > criterionErrorNumericValue)
				return ComparisonValue.GreaterThan;
			else
				return ComparisonValue.Equal;
		}

		private static int GetErrorNumericValue(eErrorType error)
		{
			switch (error)
			{
				case eErrorType.Null:
					return 0;
				case eErrorType.Div0:
					return 1;
				case eErrorType.Value:
					return 2;
				case eErrorType.Ref:
					return 3;
				case eErrorType.Name:
					return 4;
				case eErrorType.Num:
					return 5;
				case eErrorType.NA:
					return 6;
				default:
					throw new ArgumentException("Invalid error type.");
			}
		}

		/// <summary>
		/// Try to parse the given <paramref name="rawCriterionString"/> as an expression.
		/// </summary>
		/// <param name="rawCriterionString">The string to parse.</param>
		/// <param name="expressionOperator">The returned <see cref="IOperator"/> indicating what kind of expression was contained in <paramref name="rawCriterionString"/>.</param>
		/// <param name="expressionCriterion">The remainder of <paramref name="rawCriterionString"/> without the leading expression characters.</param>
		/// <returns>
		///		Returns true if <paramref name="rawCriterionString"/> was successfully parsed to a valid expression, and false otherwise.
		///		It is recommended not to use the results contained in <paramref name="expressionOperator"/> or 
		///		<paramref name="expressionCriterion"/> if this function returns false.</returns>
		private static bool TryParseCriterionAsExpression(string rawCriterionString, out IOperator expressionOperator, out string expressionCriterion)
		{
			expressionOperator = null;
			expressionCriterion = null;
			var operatorIndex = -1;
			// The criterion string is an expression if it begins with the operators <>, =, >, >=, <, or <=
			if (Regex.IsMatch(rawCriterionString, @"^(<>|>=|<=){1}"))
				operatorIndex = 2;
			else if (Regex.IsMatch(rawCriterionString, @"^(=|<|>){1}"))
				operatorIndex = 1;
			if (operatorIndex != -1)
			{
				var expressionOperatorString = rawCriterionString.Substring(0, operatorIndex);
				if (OperatorsDict.Instance.TryGetValue(expressionOperatorString, out expressionOperator))
				{
					expressionCriterion = rawCriterionString.Substring(operatorIndex);
					return true;
				}
			}
			return false;
		}

		/// <summary>
		/// Try to parse the given <paramref name="criterionString"/> to a bool, double, <see cref="System.DateTime"/>,
		/// or an <see cref="ExcelErrorValue"/>.
		/// </summary>
		/// <param name="criterionString">The string to be parsed.</param>
		/// <param name="criterionObject">The returned object parsed from the <paramref name="criterionString"/>.</param>
		/// <returns>Returns true if the given <paramref name="criterionString"/> was parsed to an object, and false otherwise.</returns>
		private static bool TryParseCriterionStringToObject(string criterionString, out object criterionObject)
		{
			criterionObject = null;
			if (InternationalizationUtility.TryParseLocalBoolean(criterionString, CultureInfo.CurrentCulture, out bool criterionBool))
				criterionObject = criterionBool;
			else if (IfHelper.TryConvertObjectToDouble(criterionString, out double criterionDouble))
				criterionObject = criterionDouble;
			else if (IfHelper.TryConvertObjectToDate(criterionString, out System.DateTime criterionDate))
				criterionObject = criterionDate;
			else if (InternationalizationUtility.TryParseLocalErrorValue(criterionString, CultureInfo.CurrentCulture, out ExcelErrorValue criterionErrorValue))
				criterionObject = criterionErrorValue;
			else
				return false;
			return true;
		}
		
		/// <summary>
		/// Compares the given <paramref name="testString"/> against the model string <paramref name="criterionString"/>.
		/// </summary>
		/// <param name="testString">The string being tested.</param>
		/// <param name="criterionString">The string that provides the model for the comparison.</param>
		/// <param name="checkWildcardChars">
		///		Optional parameter that indicated whether the <paramref name="criterionString"/> should consider
		///		the characters ? and * as wildcards or literals.</param>
		/// <returns>
		///		Returns an int less than 0 if <paramref name="testString"/> precedes <paramref name="criterionString"/> in the sort order.
		///		Returns the int 0 if <paramref name="testString"/> and <paramref name="criterionString"/> match in content.
		///		Returns an int greater than 0 if <paramref name="testString"/> follows <paramref name="criterionString"/> in the sort order.
		///		Returns <see cref="int.MinValue"/> if <paramref name="criterionString"/> was comparing using wildcard characters, and <paramref name="testString"/> failed to match.</returns>
		private static ComparisonValue CompareAsStrings(string testString, string criterionString, bool checkWildcardChars = false)
		{
			var compareResult = ComparisonValue.NotEqual;
			testString = testString.ToUpper(CultureInfo.CurrentCulture);
			criterionString = criterionString.ToUpper(CultureInfo.CurrentCulture);
			if (checkWildcardChars && (criterionString.Contains("*") || criterionString.Contains("?")))
			{
				var criterionRegexPattern = Regex.Escape(criterionString);
				criterionRegexPattern = string.Format("^{0}$", criterionRegexPattern);
				criterionRegexPattern = criterionRegexPattern.Replace(@"\*", ".*");
				criterionRegexPattern = criterionRegexPattern.Replace("~.*", "\\*");
				criterionRegexPattern = criterionRegexPattern.Replace(@"\?", ".");
				criterionRegexPattern = criterionRegexPattern.Replace("~.", "\\?");
				compareResult = (Regex.IsMatch(testString, criterionRegexPattern)) ? ComparisonValue.Equal : ComparisonValue.NotEqual;
			}
			else
			{
				var compareValue = string.Compare(testString, criterionString, StringComparison.CurrentCultureIgnoreCase);
				compareResult = IfHelper.ConvertIntToComparisonValue(compareValue);
			}
			return compareResult;
		}

		private static ComparisonValue ConvertIntToComparisonValue(int valueToConvert)
		{
			if (valueToConvert < 0)
				return ComparisonValue.LessThan;
			else if (valueToConvert > 0)
				return ComparisonValue.GreaterThan;
			else
				return ComparisonValue.Equal;
		}

		private static bool TryConvertObjectToDouble(object doubleCandidate, out double resultDouble, bool parseNumericStrings = true)
		{
			resultDouble = double.MinValue;
			if (parseNumericStrings && doubleCandidate is string candidateAsString)
			{
				var doubleParsingStyle = NumberStyles.Float | NumberStyles.AllowDecimalPoint;
				if (double.TryParse(candidateAsString, doubleParsingStyle, CultureInfo.CurrentCulture, out double doubleFromString))
					resultDouble = doubleFromString;
				else
					return false;
			}
			else if (ConvertUtil.IsNumeric(doubleCandidate, true))
				resultDouble = ConvertUtil.GetValueDouble(doubleCandidate);
			else
				return false;
			return true;
		}

		private static bool TryConvertObjectToDate(object dateCandidate, out System.DateTime resultDate)
		{
			resultDate = System.DateTime.MinValue;
			if (dateCandidate is System.DateTime candidateAsDate)
				resultDate = candidateAsDate;
			else if (dateCandidate is string candidateAsString)
			{
				var dateParsingStyle = DateTimeStyles.NoCurrentDateDefault;
				var timeStringParsed = System.DateTime.TryParse(candidateAsString, CultureInfo.CurrentCulture.DateTimeFormat, dateParsingStyle, out System.DateTime timeDate);
				var dateStringParsed = System.DateTime.TryParse(candidateAsString, out System.DateTime timeDateFromInput);
				if (timeStringParsed && dateStringParsed)
					resultDate = timeDate;
				else
					return false;
			}
			else
				return false;
			
			return true;
		}
		#endregion
	}
}
