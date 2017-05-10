using System;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeOpenXml.Utils
{
	internal static class ConvertUtil
	{
		internal static bool IsNumeric(object candidate)
		{
			if (candidate == null)
				return false;
			return candidate is byte || candidate is sbyte || candidate is short || candidate is ushort || candidate is int || candidate is uint ||
					candidate is long || candidate is ulong || candidate is Single || candidate is double || candidate is decimal || candidate is bool ||
					candidate is DateTime || ConvertUtil.TryParseDateString(candidate, out DateTime date) || candidate is TimeSpan;
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
		/// Convert an object value to a double.
		/// </summary>
		/// <param name="value">The value to convert to a double.</param>
		/// <param name="ignoreBool">If the value is a boolean, indicates the boolean value should be ignored and 0 returned instead.</param>
		/// <param name="returnNaN">If true, returns NaN if the double is invalid; otherwise returns 0.</param>
		/// <returns>The value converted to a double, if possible; otherwise, returns NaN or 0 based on the value of <paramref name="returnNaN"/>.</returns>
		internal static double GetValueDouble(object value, bool ignoreBool = false, bool returnNaN = false)
		{
			double result;
			try
			{
				if (ignoreBool && value is bool)
					return 0;
				if (ConvertUtil.IsNumeric(value))
				{
					if (value is DateTime)
						result = ((DateTime)value).ToOADate();
					else if (ConvertUtil.TryParseDateString(value, out DateTime date))
						result = date.ToOADate();
					else if (value is TimeSpan)
						result = DateTime.FromOADate(0).Add((TimeSpan)value).ToOADate();
					else
						result = Convert.ToDouble(value, CultureInfo.InvariantCulture);
				}
				else
					result = returnNaN ? double.NaN : 0;
			}
			catch
			{
				result = returnNaN ? double.NaN : 0;
			}
			return result;
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

		#region Internal Cache Objects
		internal static TextInfo _invariantTextInfo = CultureInfo.InvariantCulture.TextInfo;
		internal static CompareInfo _invariantCompareInfo = CompareInfo.GetCompareInfo(CultureInfo.InvariantCulture.LCID);
		#endregion
	}
}
