
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
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeOpenXml.Utils
{
	/// <summary>
	/// Utility to convert values.
	/// </summary>
	public static class ConvertUtil
	{
		#region Static Methods
		internal static bool IsNumeric(object candidate, bool ignoreBool = false)
		{
			if (candidate == null)
				return false;
			else if (ignoreBool && candidate is bool)
				return false;
			return (candidate.GetType().IsPrimitive || candidate is double || candidate is decimal || candidate is long || candidate is DateTime || candidate is TimeSpan);
		}

		/// <summary>
		/// Tries to parse a double from the specified <paramref name="candidate"/> which is expected to be a string value.
		/// </summary>
		/// <param name="candidate">The string value.</param>
		/// <param name="result">The double value parsed from the specified <paramref name="candidate"/>.</param>
		/// <returns>True if <paramref name="candidate"/> could be parsed to a double; otherwise, false.</returns>
		internal static bool TryParseNumericString(object candidate, out double result)
		{
			if (candidate != null)
			{
				// If a number is stored in a string, Excel will not convert it to the invariant format, so assume that it is in the current culture's number format.
				// This may not always be true, but it is a better assumption than assuming it is always in the invariant culture, which will probably never be true
				// for locales outside the United States.
				var style = NumberStyles.Float | NumberStyles.AllowThousands;
				return double.TryParse(candidate.ToString(), style, CultureInfo.CurrentCulture, out result);
			}
			result = 0;
			return false;
		}

		/// <summary>
		/// Tries to parse a boolean value from the specificed <paramref name="candidate"/>.
		/// </summary>
		/// <param name="candidate">The value to check for boolean-ness.</param>
		/// <param name="result">The boolean value parsed from the specified <paramref name="candidate"/>.</param>
		/// <returns>True if <paramref name="candidate"/> could be parsed </returns>
		internal static bool TryParseBooleanString(object candidate, out bool result)
		{
			if (candidate != null)
				return bool.TryParse(candidate.ToString(), out result);
			result = false;
			return false;
		}
		/// <summary>
		/// Tries to parse a <see cref="DateTime"/> from the specified <paramref name="candidate"/> which is expected to be a string value.
		/// </summary>
		/// <param name="candidate">The string value.</param>
		/// <param name="result">The double value parsed from the specified <paramref name="candidate"/>.</param>
		/// <returns>True if <paramref name="candidate"/> could be parsed to a double; otherwise, false.</returns>
		internal static bool TryParseDateString(object candidate, out DateTime result)
		{
			if (candidate != null)
			{
				var style = DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AssumeLocal;
				// If a date is stored in a string, Excel will not convert it to the invariant format, so assume that it is in the current culture's date/time format.
				// This may not always be true, but it is a better assumption than assuming it is always in the invariant culture, which will probably never be true
				// for locales outside the United States.
				return DateTime.TryParse(candidate.ToString(), CultureInfo.CurrentCulture, style, out result);
			}
			result = DateTime.MinValue;
			return false;
		}
		/// <summary>
		/// Tries to parse the given object into the Excel OADate version of that object. Only integers, doubles,
		/// strings, and <see cref="DateTime"/> objects have the possibility of being successfully parsed.
		/// IMPORTANT: This method assumes that any ints or doubles passed into it are already Excel OADates;
		/// Do not pass in System.DateTime OADates as ints/doubles if the OADate is less than 61, because
		/// the result of this method will be incorrect.
		/// </summary>
		/// <param name="dateCandidate">The object to convert into an Excel OADate.</param>
		/// <param name="OADate">The resulting Excel OADate that <paramref name="dateCandidate"/> was converted to.</param>
		/// <returns>Return true if the given object was successfully parsed into an Excel OADate, or false otherwise.</returns>
		public static bool TryParseObjectToDecimal(object dateCandidate, out double OADate)
		{
			OADate = -1.0;
			if (dateCandidate is DateTime dateDateTime)
			{
				OADate = dateDateTime.ToOADate();
				// Note: This if statement is to account for an error from Lotus 1-2-3
				// that Excel implemented which incorrectly includes 2/29/1900 as a valid date;
				// that day does not actually exist: See link for more information.
				// https://support.microsoft.com/en-us/help/214058/days-of-the-week-before-march-1,-1900-are-incorrect-in-excel
				if (OADate < 61)
					OADate--;
				return true;
			}
			else if (dateCandidate is string dateString)
			{
				var doubleParsingStyle = NumberStyles.Float | NumberStyles.AllowDecimalPoint;
				var dateParsingStyle = DateTimeStyles.NoCurrentDateDefault;
				if (double.TryParse(dateString, doubleParsingStyle, CultureInfo.CurrentCulture, out double dateDouble))
				{
					OADate = dateDouble;
					return true;
				}
				var timeStringParsed = DateTime.TryParse(dateString, CultureInfo.CurrentCulture.DateTimeFormat, dateParsingStyle, out DateTime timeDate);
				var dateStringParsed = DateTime.TryParse(dateString, out DateTime timeDateFromInput);
				if(timeStringParsed && dateStringParsed)
				{
					if (timeDate.Equals(timeDateFromInput))
					{
						OADate = timeDate.ToOADate();
						// Note: This if statement is to account for an error from Lotus 1-2-3
						// that Excel implemented which incorrectly includes 2/29/1900 as a valid date;
						// that day does not actually exist: See link for more information.
						// https://support.microsoft.com/en-us/help/214058/days-of-the-week-before-march-1,-1900-are-incorrect-in-excel
						if (OADate < 61 )
							OADate--;
						return true;
					}
					else
					{
						//Note: This if statement is to account for when a pure time string is the input.
						OADate = timeDate.ToOADate();
						return true;
					}
				}
				//else
					//return false;
			}
			else if (dateCandidate is int dateInt)
			{
				OADate = dateInt;
				return true;
			}
			else if (dateCandidate is double dateDouble)
			{
				OADate = dateDouble;
				return true;
			}
			else if (dateCandidate is decimal dateDecimal)
			{
				OADate = (double)dateDecimal;
				return true;
			}
			else if(dateCandidate is bool dateBool)
			{

				OADate = dateBool ? 1 : 0 ;
				return true;
			}
			return false;
		}
		/// <summary>
		/// Tries to parse the given object into a <see cref="DateTime"/>. Only integers, doubles, strings
		/// and <see cref="DateTime"/> objects have the possibility of being successfully parsed.
		/// </summary>
		/// <param name="dateCandidate">The object to convert into a DateTime if valid.</param>
		/// <param name="date">The resulting <see cref="DateTime"/> that dateCandidate was converted to.</param>
		/// <param name="error">Null if the parse was successful, or the <see cref="eErrorType"/> indicating why the parse was unsuccessful.</param>
		/// <returns>True if <paramref name="dateCandidate"/> was successfully parsed into a <see cref="DateTime"/>, and false otherwise.</returns>
		public static bool TryParseDateObject(object dateCandidate, out DateTime date, out eErrorType? error)
		{
			error = null;
			date = DateTime.MinValue;
			if (dateCandidate is DateTime validDate)
			{
				date = validDate;
				return true;
			}
			else if (TryParseObjectToDecimal(dateCandidate, out double OADate))
			{
				// Note: This if statement is to account for an error from Lotus 1-2-3
				// that Excel implemented which incorrectly includes 2/29/1900 as a valid date;
				// that day does not actually exist: See link for more information.
				// https://support.microsoft.com/en-us/help/214058/days-of-the-week-before-march-1,-1900-are-incorrect-in-excel
				if (OADate < 61)
					OADate++;
				if (OADate >= 2)
				{
					date = DateTime.FromOADate(OADate);
					return true;
				}
				else
				{
					error = eErrorType.Num;
					return false;
				}
			}
			error = eErrorType.Value;
			return false;
		}

		/// <summary>
		/// Convert an object value to a double 
		/// </summary>
		/// <param name="v"></param>
		/// <param name="ignoreBool"></param>
		/// <param name="retNaN">Return NaN if invalid double otherwise 0</param>
		/// <returns></returns>
		internal static double GetValueDouble(object v, bool ignoreBool = false, bool retNaN = false)
		{
			double d;
			try
			{
				if (ignoreBool && v is bool)
				{
					return 0;
				}
				if (IsNumeric(v))
				{
					if (v is DateTime)
					{
						d = ((DateTime)v).ToOADate();
					}
					else if (v is TimeSpan)
					{
						d = DateTime.FromOADate(0).Add((TimeSpan)v).ToOADate();
					}
					else
					{
						d = Convert.ToDouble(v, CultureInfo.InvariantCulture);
					}
				}
				else
				{
					d = retNaN ? double.NaN : 0;
				}
			}

			catch
			{
				d = retNaN ? double.NaN : 0;
			}
			return d;
		}
		/// <summary>
		/// OOXML requires that "," , and &amp; be escaped, but ' and " should *not* be escaped, nor should
		/// any extended Unicode characters. This function only encodes the required characters.
		/// System.Security.SecurityElement.Escape() escapes ' and " as  &apos; and &quot;, so it cannot
		/// be used reliably. System.Web.HttpUtility.HtmlEncode overreaches as well and uses the numeric
		/// escape equivalent.
		/// </summary>
		/// <param name="s"></param>
		/// <returns></returns>
		internal static string ExcelEscapeString(string s)
		{
			return s.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");
		}
		/// <summary>
		/// Return true if preserve space attribute is set.
		/// </summary>
		/// <param name="sw"></param>
		/// <param name="t"></param>
		/// <returns></returns>
		internal static void ExcelEncodeString(StreamWriter sw, string t)
		{
			if (Regex.IsMatch(t, "(_x[0-9A-F]{4,4}_)"))
			{
				var match = Regex.Match(t, "(_x[0-9A-F]{4,4}_)");
				int indexAdd = 0;
				while (match.Success)
				{
					t = t.Insert(match.Index + indexAdd, "_x005F");
					indexAdd += 6;
					match = match.NextMatch();
				}
			}
			for (int i = 0; i < t.Length; i++)
			{
				if (t[i] <= 0x1f && t[i] != '\t' && t[i] != '\n' && t[i] != '\r') //Not Tab, CR or LF
				{
					sw.Write("_x00{0}_", (t[i] < 0xf ? "0" : "") + ((int)t[i]).ToString("X"));
				}
				else
				{
					sw.Write(t[i]);
				}
			}

		}
		/// <summary>
		/// Return true if preserve space attribute is set.
		/// </summary>
		/// <param name="sb"></param>
		/// <param name="t"></param>
		/// <param name="encodeTabCRLF"></param>
		/// <returns></returns>
		internal static void ExcelEncodeString(StringBuilder sb, string t, bool encodeTabCRLF = false)
		{
			if (Regex.IsMatch(t, "(_x[0-9A-F]{4,4}_)"))
			{
				var match = Regex.Match(t, "(_x[0-9A-F]{4,4}_)");
				int indexAdd = 0;
				while (match.Success)
				{
					t = t.Insert(match.Index + indexAdd, "_x005F");
					indexAdd += 6;
					match = match.NextMatch();
				}
			}
			for (int i = 0; i < t.Length; i++)
			{
				if (t[i] <= 0x1f && ((t[i] != '\t' && t[i] != '\n' && t[i] != '\r' && encodeTabCRLF == false) || encodeTabCRLF)) //Not Tab, CR or LF
				{
					sb.AppendFormat("_x00{0}_", (t[i] < 0xf ? "0" : "") + ((int)t[i]).ToString("X"));
				}
				else
				{
					sb.Append(t[i]);
				}
			}

		}
		/// <summary>
		/// Return true if preserve space attribute is set.
		/// </summary>
		/// <param name="t"></param>
		/// <returns></returns>
		internal static string ExcelEncodeString(string t)
		{
			StringBuilder sb = new StringBuilder();
			t = t.Replace("\r\n", "\n"); //For some reason can't table name have cr in them. Replace with nl
			ExcelEncodeString(sb, t, true);
			return sb.ToString();
		}

		internal static string ExcelDecodeString(string t)
		{
			var match = Regex.Match(t, "(_x005F|_x[0-9A-F]{4,4}_)");
			if (!match.Success) return t;

			var useNextValue = false;
			var ret = new StringBuilder();
			var prevIndex = 0;
			while (match.Success)
			{
				if (prevIndex < match.Index) ret.Append(t.Substring(prevIndex, match.Index - prevIndex));
				if (!useNextValue && match.Value == "_x005F")
				{
					useNextValue = true;
				}
				else
				{
					if (useNextValue)
					{
						ret.Append(match.Value);
						useNextValue = false;
					}
					else
					{
						ret.Append((char)int.Parse(match.Value.Substring(2, 4), NumberStyles.AllowHexSpecifier));
					}
				}
				prevIndex = match.Index + match.Length;
				match = match.NextMatch();
			}
			ret.Append(t.Substring(prevIndex, t.Length - prevIndex));
			return ret.ToString();
		}

		/// <summary>
		/// Assumes the candidate string can be parsed to a number and validates the number groups sizes. 
		/// </summary>
		/// <param name="candidate">The string to be validated.</param>
		/// <param name="info">The <see cref="NumberFormatInfo"/> of the current culture.</param>
		/// <returns>True if the number group sizes are valid, otherwise false.</returns>
		internal static bool ValidateNumberGroupSizes(string candidate, NumberFormatInfo info)
		{
			if (candidate == null)
				return false;
			if (!candidate.Contains(info.NumberGroupSeparator))
				return true;
			if (!info.NumberGroupSizes.Any())
				return false;
			// Remove decimal point and decimal digits.
			if (candidate.Contains(info.NumberDecimalSeparator))
				candidate = candidate.Remove(candidate.IndexOf(info.NumberDecimalSeparator));
			// Remove scientific notation suffix.
			var eIndex = candidate.IndexOf("e", StringComparison.CurrentCultureIgnoreCase);
			if (eIndex != -1)
				candidate = candidate.Remove(eIndex);
			// Remove leading negative sign.
			candidate = candidate.Replace(info.NegativeSign, string.Empty);
			var groups = candidate.Split(info.NumberGroupSeparator.ToCharArray(), StringSplitOptions.None)
				.ToArray()
				.Reverse();
			int expectedGroupCount = 0;
			for (int i = 0; i < groups.Count(); i++)
			{
				var group = groups.ElementAt(i);
				if (i < info.NumberGroupSizes.Count())
					expectedGroupCount = info.NumberGroupSizes.ElementAt(i);
				// The last group can have fewer than the expected count.
				if (i + 1 == groups.Count())
				{
				// An expected count of 0 indicates no limit on group size.
					if (expectedGroupCount == 0)
						return true;
					var lastGroupCount = groups.Last().Count();
					return 0 < lastGroupCount && lastGroupCount <= expectedGroupCount;
				}
				if (expectedGroupCount != 0 && expectedGroupCount != group.Count())
					return false;
			}
			return true;
		}

		/// <summary>
		/// Converts an object to the string representation that Excel uses in XML attributes.
		/// </summary>
		/// <param name="value">The object value to convert to a string.</param>
		/// <returns>The string representation of the provided parameter.</returns>
		internal static string ConvertObjectToXmlAttributeString(object value)
		{
			if (value == null)
				return null;
			if (value is DateTime dateTimeVal)
				return dateTimeVal.ToString("yyyy’-‘MM’-‘dd’T’HH’:’mm’:’ss");
			else if (ConvertUtil.IsNumeric(value, true) || value is string)
				return value.ToString();
			else if (value is bool boolVal)
				return boolVal ? "1" : "0";
			else if (value is ExcelErrorValue errorValue)
				return errorValue.ToString();
			throw new InvalidOperationException($"Unknown type '{value.GetType()}' in cacheRecord value.");
		}
		#endregion

		#region internal cache objects
		internal static TextInfo _invariantTextInfo = CultureInfo.InvariantCulture.TextInfo;
		internal static CompareInfo _invariantCompareInfo = CompareInfo.GetCompareInfo(CultureInfo.InvariantCulture.LCID);
		#endregion
	}
}
