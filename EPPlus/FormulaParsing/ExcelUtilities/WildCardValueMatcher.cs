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
using System.Text.RegularExpressions;

namespace OfficeOpenXml.FormulaParsing.ExcelUtilities
{
	public class WildCardValueMatcher : ValueMatcher
	{
		protected override int CompareStringToString(string searchString, string testValue)
		{
			if (searchString.Contains("*") || searchString.Contains("?"))
			{
				Regex regex = this.TranslateExcelMatchStringToRegex(searchString);
				if (regex.IsMatch(testValue))
					return 0;
			}
			return base.CompareStringToString(searchString, testValue);
		}

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
			// Format as regex.
			return new Regex(string.Format("^{0}$", regexPattern));
		}
	}
}
