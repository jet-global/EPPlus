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
 * Mats Alm   		                Added		                2015-01-15
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	/// <summary>
	/// Evaluates Excel SUMIFS formulas.
	/// </summary>
	public class SumIfs : MultipleRangeCriteriasFunction
	{
		#region ExcelFunction Overrides
		/// <summary>
		/// Executes the function with the provided <paramref name="arguments"/>.
		/// </summary>
		/// <param name="arguments">Arguments to the function, each argument can contain primitive types, lists or <see cref="ExcelDataProvider.IRangeInfo">Excel ranges</see></param>
		/// <param name="context">The <see cref="ParsingContext"/> contains various data that can be useful in functions.</param>
		/// <returns>A <see cref="CompileResult"/> containing the calculated value</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			var functionArguments = arguments as FunctionArgument[] ?? arguments.ToArray();
			if (functionArguments.Length < 3 || (functionArguments.Length - 1) % 2 != 0)
				return new CompileResult(eErrorType.Value);
			var sumRange = functionArguments[0].ValueAsRangeInfo;
			var sumRangeHeight = sumRange.Address.End.Row - sumRange.Address.Start.Row + 1;
			var sumRangeWidth = sumRange.Address.End.Column - sumRange.Address.Start.Column + 1;
			var criteria = this.GetCriteria(functionArguments);
			if (criteria.Any(c => !c.ValidateDimensionality(sumRangeHeight, sumRangeWidth)))
				return new CompileResult(eErrorType.Value);
			var toSum = new List<object>();
			for (int rowOffset = 0; rowOffset < sumRangeHeight; rowOffset++)
			{
				for (int colOffset = 0; colOffset < sumRangeWidth; colOffset++)
				{
					if (criteria.All(c => c.ValidateCriteriaAtOffset(rowOffset, colOffset, base.Evaluate)))
					{
						var value = sumRange.GetOffset(rowOffset, colOffset);
						var valueAsError = value as ExcelErrorValue;
						if (valueAsError != null)
							return new CompileResult(valueAsError.Type);
						toSum.Add(value);
					}
				}
			}
			var result = toSum.Where(o => this.IsNumeric(o)).Sum(o => this.ArgToDecimal(o));
			return CreateResult(result, DataType.Decimal);
		}
		#endregion

		#region Private Methods
		/// <summary>
		/// Takses the given arguments and gathers the criteria into an enumerable list.
		/// </summary>
		/// <param name="arguments">The given arguments to be turned into a criteria list.</param>
		/// <returns>The criteria needed to evaluate the SUMIF Function.</returns>
		private IEnumerable<Criteria> GetCriteria(FunctionArgument[] arguments)
		{
			var criteria = new List<Criteria>();
			for (int i = 1; i < arguments.Length; i += 2)
			{
				criteria.Add(new Criteria(arguments[i].ValueAsRangeInfo, arguments[i + 1].ValueFirst?.ToString()));
			}
			return criteria;
		}
		#endregion

		#region Nested Classes
		private class Criteria
		{
			#region Properties
			private ExcelDataProvider.IRangeInfo Range { get; }

			private string CriteriaString { get; }
			#endregion

			#region Constructors
			/// <summary>
			/// Creates an instance of a <see cref="Criteria"/> with the provided <paramref name="range"/>
			/// and <paramref name="criteriaString"/>.
			/// </summary>
			/// <param name="range">The range to validate against.</param>
			/// <param name="criteriaString">The criteria used for validation.</param>
			public Criteria(ExcelDataProvider.IRangeInfo range, string criteriaString)
			{
				if (range == null)
					throw new ArgumentNullException(nameof(range));
				this.Range = range;
				this.CriteriaString = criteriaString ?? "0";
			}
			#endregion

			#region Public Methods
			/// <summary>
			/// Validates that the validation range is of a specific dimension.
			/// </summary>
			/// <param name="height">The desired height of the criteria range.</param>
			/// <param name="width">The desired width of the criteria range.</param>
			/// <returns>true if the criteria range is of the correct dimensions; otherwise false.</returns>
			public bool ValidateDimensionality(int height, int width)
			{
				var myAddress = this.Range.Address;
				var myHeight = myAddress.End.Row - myAddress.Start.Row + 1;
				if (myHeight != height)
					return false;
				var myWidth = myAddress.End.Column - myAddress.Start.Column + 1;
				if (myWidth != width)
					return false;
				return true;
			}

			/// <summary>
			/// Validates a specific cell in the criteria range using the provided <paramref name="rowOffset"/>
			/// and <paramref name="colOffset"/>.
			/// </summary>
			/// <param name="rowOffset">The row offset used to locate the desired cell in the validation range.</param>
			/// <param name="colOffset">The column offset used to locate the desired cell in the validation range.</param>
			/// <param name="validate">The validation delegate.</param>
			/// <returns>true if the desired cell's value validates against the criteria string; otherwise false.</returns>
			public bool ValidateCriteriaAtOffset(int rowOffset, int colOffset, Func<object, string, bool> validate)
			{
				var valueToValidate = this.Range.GetOffset(rowOffset, colOffset);
				return validate(valueToValidate, this.CriteriaString);
			}
			#endregion
		}
		#endregion
	}
}
