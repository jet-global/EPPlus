using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	public static class IfHelper
	{
		public static bool objectMatchesCriteria(object objectToCompare, string criteria)
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
							return isMatch(objectToCompare, criteriaString, true);
						case OperatorType.NotEqualTo:
							return !isMatch(objectToCompare, criteriaString, true);
						case OperatorType.GreaterThan:
						case OperatorType.GreaterThanOrEqual:
						case OperatorType.LessThan:
						case OperatorType.LessThanOrEqual:
							return compareAsInequalityExpression(objectToCompare, criteriaString, expressionOperator.Operator);
						default:
							return isMatch(objectToCompare, criteriaString);
					}
				}
				else
					return false;
			}
			else
				return isMatch(objectToCompare, criteria);
		}

		/// criteria is either a number, boolean, or string. The string can contain a date/time, or
		/// text string that may require wildcard Regex.
		private static bool isMatch(object objectToCompare, string criteria, bool matchAsEqualityExpression = false)
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
				// A return value of 0 from CompareTo means that the two strings have equivalent content.
				return (criteria.CompareTo(objectAsString) == 0);
		}

		private static bool compareAsInequalityExpression(object objectToCompare, string criteria, OperatorType comparisonOperator)
		{
			if (objectToCompare == null)
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
					return false;
			}
		}

		public static bool IsNumeric(object val, bool excludeBool = false)
		{
			if (val == null)
				return false;
			if (excludeBool && val is bool)
				return false;
			return (val.GetType().IsPrimitive || val is double || val is decimal || val is System.DateTime || val is TimeSpan);
		}
	}
}
