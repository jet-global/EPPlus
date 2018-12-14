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
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeOpenXml.FormulaParsing.ExcelUtilities
{
	public class WildCardValueMatcher : ValueMatcher
	{
		#region Public Methods
		/// <summary>
		/// Convert Excel wildcard values to regular expression values in a given string.
		/// </summary>
		/// <param name="searchString">The string that contains Excel wildcard values.</param>
		/// <returns>A string with only regular expression values.</returns>
		public string ExcelWildcardToRegex(string searchString)
		{
			var regexString = new StringBuilder();
			for (int i = 0; i < searchString.Length; i++)
			{
				var currentCharacter = searchString[i];
				// Check if there are any escaped characters.
				if (currentCharacter == '~' && i != searchString.Length - 1)
				{
					var escapeCharacter = searchString[i + 1];
					if (escapeCharacter == '~')
						regexString.Append(escapeCharacter);
					else if (escapeCharacter == '?' || escapeCharacter == '*')
						regexString.Append(Regex.Escape(escapeCharacter.ToString()));
					i++;
				}
				else if (currentCharacter == '?')
					regexString.Append('.');
				else if (currentCharacter == '*')
				{
					regexString.Append('.');
					regexString.Append(currentCharacter);
				}
				else
					regexString.Append(currentCharacter);
			}
			return regexString.ToString();
		}

		protected override int? CompareStringToString(string searchString, string testValue)
		{
			if (searchString.Contains("*") || searchString.Contains("?") || searchString.Contains("~"))
			{
				Regex regex = this.TranslateExcelMatchStringToRegex(searchString);
				if (regex.IsMatch(testValue))
					return 0;
			}
			return base.CompareStringToString(searchString, testValue);
		}
		#endregion

		#region Private Methods
		private Regex TranslateExcelMatchStringToRegex(string searchString)
		{
			// Escape all regex special characters (including * and ?).
			var regexPattern = Regex.Escape(searchString);
			// Un-escape * and replace with ".*" because we want * to act as a wild card.
			regexPattern = regexPattern.Replace(@"\*", ".*");
			// Check for Excel escaped wildcards ("~*" (which is now "~.*" (see above))) and replaced with Regex escaped wildcards.
			regexPattern = regexPattern.Replace(@"~.*", @"\*");
			// Un-escape ? and replace with . because we want ? to match a single character.
			regexPattern = regexPattern.Replace(@"\?", ".");
			// Check for Excel escaped single character match ("~?" (which is now "~." (see above))) and replace with Regex escaped question mark.
			regexPattern = regexPattern.Replace(@"~.", @"\?");
			// Un-escape any escaped Excel escape characters since it's not a special character in regex.
			regexPattern = regexPattern.Replace("~~", "~");
			// Start and end characters for full string match.
			return new Regex(string.Format("^{0}$", regexPattern));
		}
		#endregion
	}
}
