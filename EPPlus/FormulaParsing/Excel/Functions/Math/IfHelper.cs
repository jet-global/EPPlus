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
		public static bool ObjectMatchesCriteria(object testObject, object rawCriterionObject)
		{
			object criterion = rawCriterionObject;
			OperatorType criterionOperator = OperatorType.Equals;
			bool criterionIsExpression = false;
			if (rawCriterionObject is string rawCriterionString)
			{
				string criterionString;
				if (TryParseCriterionAsExpression(rawCriterionString, out IOperator expressionOperator, out string expressionCriterion))
				{
					criterionOperator = expressionOperator.Operator;
					criterionString = expressionCriterion;
					criterionIsExpression = true;
				}
				else
					criterionString = rawCriterionString.ToUpper(CultureInfo.CurrentCulture);

				if (TryParseCriterionStringToObject(criterionString, out object criterionObject))
					criterion = criterionObject;
				else
					criterion = criterionString;
			}

			return IsMatch(testObject, criterionOperator, criterion, criterionIsExpression);
		}

		private static bool IsMatch(object testObject, OperatorType criterionOperation, object criterionObject, bool matchCriterionAsExpression = false)
		{
			/*
			 * 1. Create an if block for each kind of datatype the criterion could be (number, string, error, bool).
			 *	If the criterion is a string, it will always contain text or be empty; parsable objects are already parsed out.
			 * 2. Determine what kind of datatype the testObject is (number, text, error, bool, null).
			 * 3. Enter the criterion if statement for the correct data type, and check if the testObject is a matching data type.
			 * 4. Do a final switch for each operator type that returns the result.
			 */
			var compareResult = int.MinValue;
			bool equalityOperation = (criterionOperation == OperatorType.Equals || criterionOperation == OperatorType.NotEqualTo);
			if (criterionObject is string criterionString)
			{
				if (criterionString.Equals(string.Empty))
				{
					if (matchCriterionAsExpression)
						compareResult = (testObject == null) ? 0 : int.MinValue;
					else
						compareResult = (testObject == null || testObject.Equals(string.Empty)) ? 0 : int.MinValue;
				}
				else if (testObject is string testString)
					compareResult = CompareAsStrings(testString, criterionString, equalityOperation);
				else
					compareResult = int.MinValue;
			}
			else if (criterionObject is bool criterionBool)
			{
				if (testObject is bool testBool)
					compareResult = testBool.CompareTo(criterionBool);
				else
					compareResult = int.MinValue;
			}
			else if (criterionObject is System.DateTime criterionDate)
			{
				if (TryConvertObjectToDouble(testObject, out double testDateDouble))
					compareResult = testDateDouble.CompareTo(criterionDate.ToOADate());
				else if (TryConvertObjectToDate(testObject, out System.DateTime testDate) && equalityOperation)
					compareResult = System.DateTime.Compare(testDate, criterionDate);
				else
					compareResult = int.MinValue;
			}
			else if (IsNumeric(criterionObject, true))
			{
				if (TryConvertObjectToDouble(testObject, out double testDouble, equalityOperation))
				{
					var criterionDouble = ConvertUtil.GetValueDouble(criterionObject, true);
					compareResult = testDouble.CompareTo(criterionDouble);
				}
				else
					compareResult = int.MinValue;
			}
			else if (criterionObject is ExcelErrorValue criterionErrorValue)
			{
				if (testObject is ExcelErrorValue testErrorValue && equalityOperation)
					compareResult = (criterionErrorValue.Type == testErrorValue.Type) ? 0 : int.MinValue;
				else
					compareResult = int.MinValue;
			}

			switch (criterionOperation)
			{
				case OperatorType.Equals:
					return (compareResult == 0);
				case OperatorType.NotEqualTo:
					return (compareResult != 0);
				case OperatorType.LessThan:
					return (compareResult != int.MinValue && compareResult < 0);
				case OperatorType.LessThanOrEqual:
					return (compareResult != int.MinValue && compareResult <= 0);
				case OperatorType.GreaterThan:
					return (compareResult != int.MinValue && compareResult > 0);
				case OperatorType.GreaterThanOrEqual:
					return (compareResult != int.MinValue && compareResult >= 0);
				default:
					throw new InvalidOperationException("The criterionOperation is an invalid operator type for this function.");
			}
		}

		private static bool TryParseCriterionAsExpression(string rawCriterionString, out IOperator expressionOperator, out string expressionCriterion)
		{
			expressionOperator = null;
			expressionCriterion = null;
			var operatorIndex = -1;
			// The criteria string is an expression if it begins with the operators <>, =, >, >=, <, or <=
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

		private static bool TryParseCriterionStringToObject(string criterionString, out object criterionObject)
		{
			criterionObject = null;
			if (InternationalizationUtil.TryParseLocalBoolean(criterionString, CultureInfo.CurrentCulture, out bool criterionBool))
				criterionObject = criterionBool;
			else if (TryConvertObjectToDouble(criterionString, out double criterionDouble))
				criterionObject = criterionDouble;
			else if (TryConvertObjectToDate(criterionString, out System.DateTime criterionDate))
				criterionObject = criterionDate;
			else if (InternationalizationUtil.TryParseLocalErrorValue(criterionString, CultureInfo.CurrentCulture, out ExcelErrorValue criterionErrorValue))
				criterionObject = criterionErrorValue;
			else
				return false;
			return true;
		}

		private static bool TryConvertObjectToDouble(object convertObject, out double objectAsDouble, bool includeNumericStrings = true)
		{
			objectAsDouble = double.MinValue;
			if (IsNumeric(convertObject, true))
				objectAsDouble = ConvertUtil.GetValueDouble(convertObject);
			else if (includeNumericStrings && convertObject is string objectAsString && Double.TryParse(objectAsString, out double doubleParseResult))
				objectAsDouble = doubleParseResult;
			else
				return false;
			return true;
		}

		private static int CompareAsStrings(string testString, string criterionString, bool checkWildcardChars = false)
		{
			var compareResult = int.MinValue;
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
				compareResult = (Regex.IsMatch(testString, criterionRegexPattern)) ? 0 : int.MinValue;
			}
			else
				compareResult = string.Compare(testString, criterionString);
			return compareResult;
		}

		private static bool TryConvertObjectToDouble(object doubleCandidate, out double resultDouble)
		{
			resultDouble = double.MinValue;
			if (doubleCandidate is string candidateAsString)
			{
				var doubleParsingStyle = NumberStyles.Float | NumberStyles.AllowDecimalPoint;
				if (double.TryParse(candidateAsString, doubleParsingStyle, CultureInfo.CurrentCulture, out double doubleFromString))
					resultDouble = doubleFromString;
				else
					return false;
			}
			else if (doubleCandidate is int candidateAsInt)
				resultDouble = candidateAsInt;
			else if (doubleCandidate is double candidateAsDouble)
				resultDouble = candidateAsDouble;
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

		//private static bool CompareAsInequalityExpression(object objectToCompare, ComparisonDataType objectDataType, OperatorType comparisonOperator, string criteriaString)
		//{
		//	var comparisonResult = int.MinValue;
		//	if (objectToCompare is string)
		//		objectDataType = ComparisonDataType.TextValue;
		//	if (TryExtractObjectFromCriteriaString(criteriaString, out object criteriaObject, out ComparisonDataType criteriaDataType))
		//	{
		//		if (objectDataType != criteriaDataType)
		//			return false;
		//		switch (criteriaDataType)
		//		{
		//			case ComparisonDataType.NumericValue:
		//				{
		//					var criteriaValue = ConvertUtil.GetValueDouble(criteriaObject);
		//					var objectToCompareValue = ConvertUtil.GetValueDouble(objectToCompare);
		//					comparisonResult = objectToCompareValue.CompareTo(criteriaValue);
		//					break;
		//				}
		//			case ComparisonDataType.BooleanValue:
		//				{
		//					if (objectToCompare is bool boolToCompare && criteriaObject is bool criteriaBool)
		//						comparisonResult = boolToCompare.CompareTo(criteriaBool);
		//					else
		//						return false;
		//					break;
		//				}
		//		}
		//	}
		//	else if (objectDataType == ComparisonDataType.TextValue)
		//		comparisonResult = string.Compare(objectToCompare.ToString(), criteriaString, StringComparison.CurrentCultureIgnoreCase);
		//	else
		//		return false;

		//	switch (comparisonOperator)
		//	{
		//		case OperatorType.LessThan:
		//			return (comparisonResult == -1);
		//		case OperatorType.LessThanOrEqual:
		//			return (comparisonResult == -1 || comparisonResult == 0);
		//		case OperatorType.GreaterThan:
		//			return (comparisonResult == 1);
		//		case OperatorType.GreaterThanOrEqual:
		//			return (comparisonResult == 1 || comparisonResult == 0);
		//		default:
		//			throw new InvalidOperationException("This function should only be entered if the comparisonOperator is one of the 4 in the switch statement.");
		//	}
		//}

		//private static bool TryExtractObjectFromCriteriaString(string rawCriteriaString, out object objectFromCriteria, out ComparisonDataType criteriaDataType)
		//{
		//	objectFromCriteria = null;
		//	criteriaDataType = ComparisonDataType.TextValue;
		//	if (ConvertUtil.TryParseDateObjectToOADate(rawCriteriaString, out double criteriaDouble))
		//	{
		//		objectFromCriteria = criteriaDouble;
		//		criteriaDataType = ComparisonDataType.NumericValue;
		//	}
		//	else if (InternationalizationUtil.TryParseLocalBoolean(rawCriteriaString, CultureInfo.CurrentCulture, out bool criteriaBool))
		//	{
		//		objectFromCriteria = criteriaBool;
		//		criteriaDataType = ComparisonDataType.BooleanValue;
		//	}
		//	else if (InternationalizationUtil.TryParseLocalErrorValue(rawCriteriaString, CultureInfo.CurrentCulture, out ExcelErrorValue criteriaErrorValue))
		//	{
		//		objectFromCriteria = criteriaErrorValue;
		//		criteriaDataType = ComparisonDataType.ErrorValue;
		//	}
		//	else
		//		return false;

		//	return true;
		//}

		//private static bool CompareAsStrings(object objectToCompare, ComparisonDataType objectDataType, string criteriaString, bool compareAsEqualityExpression = false)
		//{
		//	if (criteriaString.Equals(string.Empty))
		//		return ((compareAsEqualityExpression) ? objectToCompare == null : (objectToCompare == null || objectToCompare.ToString().Equals(string.Empty)));
		//	if (objectDataType != ComparisonDataType.TextValue)
		//		return false;

		//	var stringToCompare = objectToCompare.ToString().ToUpper(CultureInfo.CurrentCulture);
		//	criteriaString = criteriaString.ToUpper(CultureInfo.CurrentCulture);

		//	if (criteriaString.Contains("*") || criteriaString.Contains("?"))
		//	{
		//		var regexPattern = Regex.Escape(criteriaString);
		//regexPattern = string.Format("^{0}$", regexPattern);
		//regexPattern = regexPattern.Replace(@"\*", ".*");
		//		regexPattern = regexPattern.Replace("~.*", "\\*");
		//		regexPattern = regexPattern.Replace(@"\?", ".");
		//		regexPattern = regexPattern.Replace("~.", "\\?");
		//		return Regex.IsMatch(stringToCompare, regexPattern);
		//	}
		//	else
		//		// A return value of 0 from string.Compare() means that the two strings have equivalent content.
		//		return (string.Compare(stringToCompare, criteriaString) == 0);
		//}

		//private static bool ObjectValueEqualsCriteriaValue(object objectToCompare, ComparisonDataType objectDataType,
		//												object criteriaObject, ComparisonDataType criteriaDataType)
		//{
		//	if (objectDataType != criteriaDataType)
		//		return false;

		//	if (criteriaDataType == ComparisonDataType.NumericValue)
		//	{
		//		ConvertUtil.TryParseDateObjectToOADate(criteriaObject, out double criteriaDouble);
		//		criteriaObject = criteriaDouble;
		//	}

		//	var objectString = objectToCompare.ToString();
		//	var criteriaString = criteriaObject.ToString();

		//	return (string.Compare(criteriaString, objectString) == 0);
		//}

		//private static ComparisonDataType GetCriterionDataType(object criterionObject)
		//{
		//	if (criterionObject == null)
		//		return ComparisonDataType.Null;
		//	else if (criterionObject is bool)
		//		return ComparisonDataType.BooleanValue;
		//	else if (IsNumeric(criterionObject, true))
		//		return ComparisonDataType.NumericValue;
		//	else if (criterionObject is string)
		//		return ComparisonDataType.TextValue;
		//	else if (criterionObject is ExcelErrorValue)
		//		return ComparisonDataType.ErrorValue;
		//	else
		//		return ComparisonDataType.InvalidComparisonDataType;
		//}

		public static object ExtractCriteriaObject(FunctionArgument criteriaCandidate, ParsingContext context)
		{
			object criteriaObject = null;
			if (criteriaCandidate.Value is ExcelDataProvider.IRangeInfo criteriaRange)
			{
				if (criteriaRange.IsMulti)
				{
					var worksheet = context.ExcelDataProvider.GetRange(context.Scopes.Current.Address.Worksheet, 1, 1, "A1").Worksheet;
					var functionRow = context.Scopes.Current.Address.FromRow;
					var functionColumn = context.Scopes.Current.Address.FromCol;
					criteriaObject = CalculateCriteria(criteriaCandidate, worksheet, functionRow, functionColumn);
				}
				else
				{
					criteriaObject = criteriaCandidate.ValueFirst;
					if (criteriaObject is List<object> objectList)
						criteriaObject = objectList.First();
				}
			}
			else if (criteriaCandidate.Value is List<FunctionArgument> argumentList)
				criteriaObject = argumentList.First().ValueFirst;
			else
				criteriaObject = criteriaCandidate.ValueFirst;

			// Note that Excel considers null criteria equivalent to a criteria of 0.
			if (criteriaObject == null)
				criteriaObject = 0;

			return criteriaObject;
		}

		// //////////////////////////////////////////////////////////////////

		/// <summary>
		/// Compares the given <paramref name="objectToCompare"/> against the given <paramref name="criteria"/>.
		/// This method is expected to be used with any of the *IF or *IFS Excel functions (ex: the AVERAGEIF function).
		/// </summary>
		/// <param name="objectToCompare">The object to compare against the given <paramref name="criteria"/>.</param>
		/// <param name="criteria">The criteria value or expression that dictates whether the given <paramref name="objectToCompare"/> passes or fails.</param>
		/// <returns>Returns true if <paramref name="objectToCompare"/> matches the <paramref name="criteria"/>.</returns>
		public static bool ObjectMatchesCriteria(object objectToCompare, string criteria)
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
							return IsMatch(objectToCompare, criteriaString, true);
						case OperatorType.NotEqualTo:
							return !IsMatch(objectToCompare, criteriaString, true);
						case OperatorType.GreaterThan:
						case OperatorType.GreaterThanOrEqual:
						case OperatorType.LessThan:
						case OperatorType.LessThanOrEqual:
							return CompareAsInequalityExpression(objectToCompare, criteriaString, expressionOperator.Operator);
						default:
							return IsMatch(objectToCompare, criteriaString);
					}
				}
			}
			return IsMatch(objectToCompare, criteria);
		}

		/// <summary>
		/// Compares the <paramref name="objectToCompare"/> with the given <paramref name="criteria"/>.
		/// <paramref name="criteria"/> is expected to be either a number, boolean, string, or null. The string
		/// can contain a date/time, or a text value that may require wildcard Regex.
		/// The given object is considered a match with the criteria if their content are equivalent in value.
		/// </summary>
		/// <param name="objectToCompare">The object to compare against the given <paramref name="criteria"/>.</param>
		/// <param name="criteria">The criteria value that dictates whether the <paramref name="objectToCompare"/> passes or fails.</param>
		/// <param name="matchAsEqualityExpression">
		///		Indicate if the <paramref name="criteria"/> came from an equality related expression,
		///		which requires slightly different handling.</param>
		/// <returns>Returns true if <paramref name="objectToCompare"/> matches the <paramref name="criteria"/>.</returns>
		private static bool IsMatch(object objectToCompare, string criteria, bool matchAsEqualityExpression = false)
		{
			// Equality related expression evaluation (= or <>) only considers empty cells as equal to empty string criteria.
			// If the given criteria was not originally preceded by an equality operator, then 
			// both empty cells and cells containing the empty string are considered as equal to empty string criteria.
			if (criteria.Equals(string.Empty))
				return ((matchAsEqualityExpression) ? (objectToCompare == null) : (objectToCompare == null || objectToCompare.Equals(string.Empty)));
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
				// A return value of 0 from CompareTo means that the two strings have equivalent content.
				return (criteria.CompareTo(objectAsString) == 0);
		}

		/// <summary>
		/// Compare the given <paramref name="objectToCompare"/> with the given <paramref name="criteria"/> using
		/// the given <paramref name="comparisonOperator"/>.
		/// </summary>
		/// <param name="objectToCompare">The object to compare against the given <paramref name="criteria"/>.</param>
		/// <param name="criteria">The criteria value that dictates whether the <paramref name="objectToCompare"/> passes or fails.</param>
		/// <param name="comparisonOperator">
		///		The inequality operator that dictates how the <paramref name="objectToCompare"/> should
		///		be compared to the <paramref name="criteria"/>.</param>
		/// <returns>Returns true if the <paramref name="objectToCompare"/> passes the comparison with <paramref name="criteria"/>.</returns>
		private static bool CompareAsInequalityExpression(object objectToCompare, string criteria, OperatorType comparisonOperator)
		{
			if (objectToCompare == null || objectToCompare is ExcelErrorValue)
				return false;
			var comparisonResult = int.MinValue;
			if (ConvertUtil.TryParseDateObjectToOADate(criteria, out double criteriaNumber)) // Handle the criteria as a number/date.
			{
				if (IsNumeric(objectToCompare, true))
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
				else if (IsNumeric(objectToCompare))
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
					throw new InvalidOperationException(
						"The default condition is invalid because this function should only be called if the given operator is <,<=,>, or >=.");
			}
		}

		/// <summary>
		/// Returns true if <paramref name="numericCandidate"/> is numeric.
		/// </summary>
		/// <param name="numericCandidate">The object to check for numeric content.</param>
		/// <param name="excludeBool">
		///		An optional parameter to exclude boolean values from the data types that are considered numeric.
		///		This method considers booleans as numeric by default.</param>
		/// <returns>Returns true if <paramref name="numericCandidate"/> is numeric.</returns>
		public static bool IsNumeric(object numericCandidate, bool excludeBool = false)
		{
			if (numericCandidate == null)
				return false;
			if (excludeBool && numericCandidate is bool)
				return false;
			return (numericCandidate.GetType().IsPrimitive || 
				numericCandidate is double || 
				numericCandidate is decimal || 
				numericCandidate is System.DateTime || 
				numericCandidate is TimeSpan);
		}

		/// <summary>
		/// Ensures that the given <paramref name="criteriaCandidate"/> is of a form that can be
		/// represented as a criteria.
		/// </summary>
		/// <param name="criteriaCandidate">The <see cref="FunctionArgument"/> containing the criteria.</param>
		/// <param name="context">The context from the function calling this function.</param>
		/// <returns>Returns the criteria in <paramref name="criteriaCandidate"/> as a string.</returns>
		public static string ExtractCriteriaString(FunctionArgument criteriaCandidate, ParsingContext context)
		{
			object criteriaObject = null;
			if (criteriaCandidate.Value is ExcelDataProvider.IRangeInfo criteriaRange)
			{
				if (criteriaRange.IsMulti)
				{
					var worksheet = context.ExcelDataProvider.GetRange(context.Scopes.Current.Address.Worksheet, 1, 1, "A1").Worksheet;
					var functionRow = context.Scopes.Current.Address.FromRow;
					var functionColumn = context.Scopes.Current.Address.FromCol;
					criteriaObject = CalculateCriteria(criteriaCandidate, worksheet, functionRow, functionColumn);
				}
				else
				{
					criteriaObject = criteriaCandidate.ValueFirst;
					if (criteriaObject is List<object> objectList)
						criteriaObject = objectList.First();
				}
			}
			else if (criteriaCandidate.Value is List<FunctionArgument> argumentList)
				criteriaObject = argumentList.First().ValueFirst;
			else
				criteriaObject = criteriaCandidate.ValueFirst;

			// Note that Excel considers null criteria equivalent to a criteria of 0.
			if (criteriaObject == null)
				return "0";
			else
				return criteriaObject.ToString().ToUpper();
		}

		public static object CalculateCriteria(FunctionArgument criteriaArgument, ExcelWorksheet worksheet, int rowLocation, int colLocation)
		{
			if (criteriaArgument.Value == null)
				return 0;
			if (criteriaArgument.Value is ExcelErrorValue)
				if (worksheet == null)
					return 0;
			if (rowLocation <= 0 || colLocation <= 0)
				return 0;

			var criteriaCandidate = criteriaArgument.ValueAsRangeInfo.Address;

			if (criteriaCandidate.Rows > criteriaCandidate.Columns)
			{
				var currentAddressRow = rowLocation;
				var startRow = criteriaCandidate.Start.Row;
				var endRow = criteriaCandidate.End.Row;

				if (currentAddressRow == startRow)
				{
					var cellColumn = criteriaCandidate.Start.Column;
					return worksheet.Cells[startRow, cellColumn].Value;
				}
				else if (currentAddressRow == endRow)
				{
					var cellColumn = criteriaCandidate.Start.Column;
					return worksheet.Cells[endRow, cellColumn].Value;
				}
				else if (currentAddressRow > startRow && currentAddressRow < endRow)
				{

					var cellColumn = criteriaCandidate.Start.Column;
					return worksheet.Cells[currentAddressRow, cellColumn].Value;
				}
				else
					return 0;
			}
			else if (criteriaCandidate.Rows < criteriaCandidate.Columns)
			{
				var currentAddressCol = colLocation;
				var startCol = criteriaCandidate.Start.Column;
				var endCol = criteriaCandidate.End.Column;

				if (currentAddressCol == startCol)
				{
					var cellRow = criteriaCandidate.Start.Row;
					return worksheet.Cells[cellRow, currentAddressCol].Value;
				}
				else if (currentAddressCol == endCol)
				{
					var cellRow = criteriaCandidate.Start.Row;
					return worksheet.Cells[cellRow, currentAddressCol].Value;
				}
				else if (currentAddressCol > startCol && currentAddressCol < endCol)
				{
					var cellRow = criteriaCandidate.Start.Row;
					return worksheet.Cells[cellRow, currentAddressCol].Value;
				}
				else
					return 0;
			}
			else
				return 0;
		}

		/// <summary>
		/// Takes a cell range and converts it into a single value criteria
		/// </summary>
		/// <param name="arguments">The cell range that will be reduced to a single value criteria.</param>
		/// <param name="worksheet">The current worksheet that is being used.</param>
		/// <param name="rowLocation">The row location of the cell that is calling this function.</param>
		/// <param name="colLocation">The column location of the cell that is calling this function.</param>
		/// <returns>A single value criteria as an integer.</returns>
		public static object CalculateCriteria(IEnumerable<FunctionArgument> arguments, ExcelWorksheet worksheet, int rowLocation, int colLocation)
		{
			if (arguments.ElementAt(1).Value == null)
				return 0;
			if (arguments.ElementAt(1).Value is ExcelErrorValue)
			if (worksheet == null)
				return 0;
			if (rowLocation <= 0 || colLocation <= 0)
				return 0;

			var criteriaCandidate = arguments.ElementAt(1).ValueAsRangeInfo.Address;

			if (criteriaCandidate.Rows > criteriaCandidate.Columns)
			{
				var currentAddressRow = rowLocation;
				var startRow = criteriaCandidate.Start.Row;
				var endRow = criteriaCandidate.End.Row;

				if (currentAddressRow == startRow)
				{
					var cellColumn = criteriaCandidate.Start.Column;
					return worksheet.Cells[startRow, cellColumn].Value;
				}
				else if (currentAddressRow == endRow)
				{
					var cellColumn = criteriaCandidate.Start.Column;
					return worksheet.Cells[endRow, cellColumn].Value;
				}
				else if (currentAddressRow > startRow && currentAddressRow < endRow)
				{

					var cellColumn = criteriaCandidate.Start.Column;
					return worksheet.Cells[currentAddressRow, cellColumn].Value;
				}
				else
					return 0;
			}
			else if (criteriaCandidate.Rows < criteriaCandidate.Columns)
			{
				var currentAddressCol = colLocation;
				var startCol = criteriaCandidate.Start.Column;
				var endCol = criteriaCandidate.End.Column;

				if (currentAddressCol == startCol)
				{
					var cellRow = criteriaCandidate.Start.Row;
					return worksheet.Cells[cellRow, currentAddressCol].Value;
				}
				else if (currentAddressCol == endCol)
				{
					var cellRow = criteriaCandidate.Start.Row;
					return worksheet.Cells[cellRow, currentAddressCol].Value;
				}
				else if (currentAddressCol > startCol && currentAddressCol < endCol)
				{
					var cellRow = criteriaCandidate.Start.Row;
					return worksheet.Cells[cellRow, currentAddressCol].Value;
				}
				else
					return 0;
			}
			else
				return 0;
		}
	}
}
