/*******************************************************************************
* You may amend and distribute as you like, but don't remove this header!
*
* EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
* See http://www.codeplex.com/EPPlus for details.
*
* Copyright (C) 2011-2018 Michelle Lau and others as noted in the source history.
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
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;
using OfficeOpenXml.Extensions;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Table.PivotTable.Filters
{
	#region Enums
	/// <summary>
	/// All the possible field filter types.
	/// </summary>
	public enum FieldFilter
	{
		Label,
		Value,
		Report
	}

	/// <summary>
	/// All the possible label filter operator options.
	/// </summary>
	public enum LabelFilterType
	{
		CaptionEqual,
		CaptionNotEqual,
		CaptionBeginsWith,
		CaptionNotBeginsWith,
		CaptionEndsWith,
		CaptionNotEndsWith,
		CaptionContains,
		CaptionNotContains,
		CaptionGreaterThan,
		CaptionGreaterThanOrEqual,
		CaptionLessThan,
		CaptionLessThanOrEqual,
		CaptionBetween,
		CaptionNotBetween
	}

	/// <summary>
	/// All the possible value filter operator options.
	/// </summary>
	public enum ValueFilterType
	{
		// TODO: Fill this out for Value Filters (Task #11843).
	}
	#endregion

	/// <summary>
	/// A filter item in the the <see cref="ExcelPivotFieldFiltersCollection"/>.
	/// </summary>
	public class ExcelPivotTableAdvancedFilter : XmlCollectionItemBase
	{
		#region Properties
		/// <summary>
		/// Gets the index of the field this filter applies to.
		/// </summary>
		public int Field
		{
			get { return base.GetXmlNodeInt("@fld"); }
		}

		/// <summary>
		/// Gets the type of the pivot filter (operation type).
		/// </summary>
		public string PivotFilterType
		{
			get { return base.GetXmlNodeString("@type"); }
		}

		/// <summary>
		/// Gets the first string value used by Label Filters only.
		/// </summary>
		public string StringValueOne
		{
			get { return base.GetXmlNodeString("@stringValue1"); }
		}

		/// <summary>
		/// Gets the second string value used by Label Filters only.
		/// Remarks: This is used only if multiple filters/captionBetween are applied to a pivot field.
		/// </summary>
		public string StringValueTwo
		{
			get { return base.GetXmlNodeString("@stringValue2"); }
		}

		/// <summary>
		/// Gets the index of the measure field used by Value Filters only.
		/// </summary>
		public string MeasureFieldIndex
		{
			get { return base.GetXmlNodeString("@iMeasureFld"); }
		}

		/// <summary>
		/// Gets the collection of custom <see cref="ExcelFilter"/>s.
		/// Remarks: This collection is used if the filter value contains a '*' or '?'.
		/// </summary>
		public ExcelCustomFiltersCollection CustomFilters { get; }

		/// <summary>
		/// Gets the collection of <see cref="ExcelFilter"/>s.
		/// </summary>
		public ExcelFilterCriteriaCollection Filters { get; }

		/// <summary>
		/// Gets the type of field filter (Report, Label, or Value).
		/// </summary>
		public FieldFilter FieldFilterType
		{
			get
			{
				if (this.PivotFilterType.Contains("caption"))
					return FieldFilter.Label;
				else
					return FieldFilter.Value;
			}
		}

		/// <summary>
		/// Gets the caption label filter.
		/// </summary>
		public LabelFilterType LabelFilterType
		{
			get
			{
				if (this.FieldFilterType == FieldFilter.Label)
					return (LabelFilterType)Enum.Parse(typeof(LabelFilterType), this.PivotFilterType, true);
				else
					throw new InvalidOperationException("Invalid filter type.");
			}
		}

		public bool HasCustomFilters => this.CustomFilters != null;
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of an <see cref="ExcelPivotTableAdvancedFilter"/>.
		/// </summary>
		/// <param name="namespaceManager">The namespace manager.</param>
		/// <param name="node">The xml top node.</param>
		public ExcelPivotTableAdvancedFilter(XmlNamespaceManager namespaceManager, XmlNode node) : base(namespaceManager, node)
		{
			if (node == null)
				throw new ArgumentNullException(nameof(node));
			var customFiltersNode = node.SelectSingleNode(".//d:customFilters", this.NameSpaceManager);
			if (customFiltersNode != null)
				this.CustomFilters = new ExcelCustomFiltersCollection(this.NameSpaceManager, customFiltersNode);
			var filterCriteriaNode = node.SelectSingleNode(".//d:filters", this.NameSpaceManager);
			if (filterCriteriaNode != null)
				this.Filters = new ExcelFilterCriteriaCollection(this.NameSpaceManager, filterCriteriaNode);
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Gets a value indicating if the given input string matches the regular expression.
		/// </summary>
		/// <param name="input">The string to find the patter in.</param>
		/// <param name="isNumericType">A flag indicating if the input string is a numerical value.</param>
		/// <returns>True if a match is found; otherwise false.</returns>
		public bool MatchesFilterCriteriaResult(string input, bool isNumericType)
		{
			bool match = false;
			string filterValue = this.HasCustomFilters ? this.CustomFilters.First().TopOrBottomValue : this.Filters.First().TopOrBottomValue;
			if (this.FieldFilterType == FieldFilter.Label)
			{
				bool hasWildcard = filterValue.Contains("*") || filterValue.Contains("?");
				filterValue = new WildCardValueMatcher().ExcelWildcardToRegex(filterValue);
				var matches = Regex.Match(input, filterValue, RegexOptions.IgnoreCase);

				if (this.LabelFilterType == LabelFilterType.CaptionEqual 
					|| this.LabelFilterType == LabelFilterType.CaptionBeginsWith 
					|| this.LabelFilterType == LabelFilterType.CaptionEndsWith 
					|| this.LabelFilterType == LabelFilterType.CaptionContains
					|| this.LabelFilterType == LabelFilterType.CaptionLessThanOrEqual)
				{
					match = matches == Match.Empty ? false : true;
				}
				else if (this.LabelFilterType == LabelFilterType.CaptionNotEqual 
					|| this.LabelFilterType == LabelFilterType.CaptionNotBeginsWith
					|| this.LabelFilterType == LabelFilterType.CaptionNotEndsWith 
					|| this.LabelFilterType == LabelFilterType.CaptionNotContains
					|| this.LabelFilterType == LabelFilterType.CaptionGreaterThan)
				{
					match = matches == Match.Empty ? true : false;
				}

				if (this.LabelFilterType == LabelFilterType.CaptionEqual || this.LabelFilterType == LabelFilterType.CaptionBeginsWith)
					match = this.CheckFilterAtStartOfString(filterValue, input, match, false, true);
				else if (this.LabelFilterType == LabelFilterType.CaptionNotEqual || this.LabelFilterType == LabelFilterType.CaptionNotBeginsWith)
					match = this.CheckFilterAtStartOfString(filterValue, input, match, true, true);
				else if (this.LabelFilterType == LabelFilterType.CaptionEndsWith)
					match = this.CheckFilterAtStartOfString(filterValue, input, match, false, false);
				else if (this.LabelFilterType == LabelFilterType.CaptionNotEndsWith)
					match = this.CheckFilterAtStartOfString(filterValue, input, match, true, false);
				else if (this.LabelFilterType == LabelFilterType.CaptionBetween || this.LabelFilterType == LabelFilterType.CaptionNotBetween)
					match = this.SatisfiesBetweenAndNotBetweenLabelFilter(input, filterValue, hasWildcard, isNumericType);
				else
					match = this.SatisfiesInequalityLabelFilter(input, filterValue, match, hasWildcard, isNumericType);
			}
			else
				throw new NotImplementedException($"{this.FieldFilterType} is not supported."); // TODO: Handle Value Filters (Task #11843).
			return match;
		}
		#endregion

		#region Private Methods
		private bool SatisfiesInequalityLabelFilter(string input, string filterValue, bool currentMatch, bool hasWildcard, bool isNumeric)
		{
			bool match = currentMatch;
			if (!hasWildcard)
			{
				if (isNumeric)
				{
					var inputValue = double.Parse(input);
					var numericFilter = double.Parse(filterValue);
					if (this.LabelFilterType == LabelFilterType.CaptionGreaterThan)
						match = inputValue > numericFilter;
					else if (this.LabelFilterType == LabelFilterType.CaptionGreaterThanOrEqual)
						match = inputValue >= numericFilter;
					else if (this.LabelFilterType == LabelFilterType.CaptionLessThan)
						match = inputValue < numericFilter;
					else if (this.LabelFilterType == LabelFilterType.CaptionLessThanOrEqual)
						match = inputValue <= numericFilter;
				}
				else
				{
					StringComparer comparer = StringComparer.CurrentCultureIgnoreCase;
					if (this.LabelFilterType == LabelFilterType.CaptionGreaterThan)
						match = comparer.Compare(input, filterValue) > 0;
					else if (this.LabelFilterType == LabelFilterType.CaptionGreaterThanOrEqual)
						match = comparer.Compare(input, filterValue) >= 0;
					else if (this.LabelFilterType == LabelFilterType.CaptionLessThan)
						match = comparer.Compare(input, filterValue) < 0;
					else if (this.LabelFilterType == LabelFilterType.CaptionLessThanOrEqual)
						match = comparer.Compare(input, filterValue) <= 0;
				}
			}
			else
			{
				// GreaterThan behaves the same as CaptionDoesNotEqual.
				// GreaterThanOrEqual shows everything.
				// LessThan does not have any results.
				// LessThanOrEqual behaves the same as CaptionEquals.
				if (this.LabelFilterType == LabelFilterType.CaptionGreaterThan)
					match = this.CheckFilterAtStartOfString(filterValue, input, match, true, true);
				else if (this.LabelFilterType == LabelFilterType.CaptionGreaterThanOrEqual)
					match = true;
				else if (this.LabelFilterType == LabelFilterType.CaptionLessThan)
					match = false;
				else if (this.LabelFilterType == LabelFilterType.CaptionLessThanOrEqual)
					match = this.CheckFilterAtStartOfString(filterValue, input, match, false, true);
			}
			return match;
		}

		private bool SatisfiesBetweenAndNotBetweenLabelFilter(string input, string filterValue, bool hasWildcard, bool isNumericType)
		{
			bool returnMatch = false;
			bool greaterThanMatch = false;
			bool lessThanMatch = false;
			if (!hasWildcard)
			{
				foreach (var customFilter in this.CustomFilters)
				{
					filterValue = customFilter.TopOrBottomValue;
					string filterOperator = customFilter.FilterComparisonOperator;
					StringComparer comparer = StringComparer.CurrentCultureIgnoreCase;
					if (this.LabelFilterType == LabelFilterType.CaptionBetween)
					{
						if (isNumericType)
						{
							var inputValue = double.Parse(input);
							var numericFilter = double.Parse(filterValue);
							if (filterOperator.IsEquivalentTo("greaterThanOrEqual"))
								greaterThanMatch = inputValue >= numericFilter;
							else if (filterOperator.IsEquivalentTo("lessThanOrEqual"))
								lessThanMatch = inputValue <= numericFilter;
						}
						else
						{
							if (filterOperator.IsEquivalentTo("greaterThanOrEqual"))
								greaterThanMatch = comparer.Compare(input, filterValue) >= 0;
							else if (filterOperator.IsEquivalentTo("lessThanOrEqual"))
								lessThanMatch = comparer.Compare(input, filterValue) <= 0;
						}
					}
					else
					{
						if (isNumericType)
						{
							var inputValue = double.Parse(input);
							var numericFilter = double.Parse(filterValue);
							if (filterOperator.IsEquivalentTo("greaterThan"))
								greaterThanMatch = inputValue > numericFilter;
							else if (filterOperator.IsEquivalentTo("lessThan"))
								lessThanMatch = inputValue < numericFilter;
						}
						else
						{
							if (filterOperator.IsEquivalentTo("greaterThan"))
								greaterThanMatch = comparer.Compare(input, filterValue) > 0;
							else if (filterOperator.IsEquivalentTo("lessThan"))
								lessThanMatch = comparer.Compare(input, filterValue) < 0;
						}
					}
				}
				returnMatch = this.HasCustomFilters && this.CustomFilters.And ? greaterThanMatch && lessThanMatch : greaterThanMatch || lessThanMatch;
			}
			return returnMatch;
		}

		private bool CheckFilterAtStartOfString(string filterValue, string input, bool currentValue, bool negatedCaption, bool startsWith)
		{
			// Filters start at the beginning of the string.
			bool returnVal = currentValue;
			if (filterValue.First() == '.' && filterValue[1] != '*' && filterValue[1] != '.')
			{
				if (!input[1].ToString().IsEquivalentTo(filterValue[1].ToString()))
					returnVal = negatedCaption;
			}
			var character = startsWith ? filterValue.First() : filterValue.Last();
			var checkCharacter = startsWith ? input.First().ToString() : input.Last().ToString();
			var filterValueString = startsWith ? filterValue.First().ToString() : filterValue.Last().ToString();
			if (character != '*' && character != '.')
			{
				if (!checkCharacter.IsEquivalentTo(filterValueString))
					returnVal = negatedCaption;
			}
			return returnVal;
		}
		#endregion
	}
}
