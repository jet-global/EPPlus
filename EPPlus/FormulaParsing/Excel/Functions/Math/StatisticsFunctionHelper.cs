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
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	/// <summary>
	/// This is a helper class that is used for all standard deviation and variance methods.
	/// </summary>
	class StatisticsFunctionHelper
	{
		#region Statistics Function Helper Methods

		/// <summary>
		///  This is used by the Stdev and Var functions to parse the values it receives from a cell range or refreence.
		/// </summary>
		/// <param name="IgnoreHiddenValues">This controls wether you should skip Hidden Values or not.</param>
		/// <param name="valueToParse">This is the value that will be parsed.</param>
		/// <param name="context">Unused, this is information about where the function is being executed.</param>
		/// <param name="includeLogicals">A flag to control wether you should parse logicals or not.</param>
		/// <param name="parsedValue">This out value is the parse value in double form.</param>
		/// <param name="theInputContainedOnlyStrings">This bool flag shows if only string inputs were given.</param>
		/// <returns>Returns true if the parseing succeded and false if it fails.</returns>
		public static bool TryToParseValuesFromInputArgumentByRefrenceOrRange(bool IgnoreHiddenValues, ExcelDataProvider.ICellInfo valueToParse, ParsingContext context, bool includeLogicals, out double parsedValue, out bool theInputContainedOnlyStrings)
		{
			var shouldIgnore = CellStateHelper.ShouldIgnore(IgnoreHiddenValues, valueToParse, context);
			var isNumeric = ConvertUtil.IsNumeric(valueToParse.Value);
			var isABoolean = valueToParse.Value is bool;

			if (!shouldIgnore && isNumeric && !isABoolean)
			{
				parsedValue = valueToParse.ValueDouble;
				theInputContainedOnlyStrings = false;
				return true;
			}
			if (includeLogicals)
			{
				if (isABoolean)
				{
					if ((bool)valueToParse.Value == true)
						parsedValue = 1;
					else
						parsedValue = 0;
					theInputContainedOnlyStrings = false;
					return true;
				}
				if (ConvertUtil.TryParseObjectToDecimal(valueToParse.ValueDouble, out double numericValue))
				{
					parsedValue = numericValue;
					theInputContainedOnlyStrings = false;
					return true;
				}
				parsedValue = 0.0;
				theInputContainedOnlyStrings = false;
				return false;
			}
			else
			{

				if (isABoolean)
				{
					parsedValue = valueToParse.ValueDoubleLogical;
					theInputContainedOnlyStrings = false;
					return false;
				}
				if (ConvertUtil.TryParseObjectToDecimal(valueToParse.ValueDouble, out double numericValue))
				{
					parsedValue = numericValue;
					theInputContainedOnlyStrings = false;
					return false;
				}
				parsedValue = 0.0;
				theInputContainedOnlyStrings = false;
				return false;
			}
		}

		/// <summary>
		/// This is used by the Stdev and Var functions to parse the values it receives from a direct input.
		/// </summary>
		/// <param name="IgnoreHiddenValues">This controls wether you should skip Hidden Values or not.</param>
		/// <param name="valueToParse">This is the value that will be parsed.</param>
		/// <param name="context">Unused, this is information about where the function is being executed.</param>
		/// <param name="parsedValue">This out value is the parse value in double form.</param>
		/// <param name="theInputContainedOnlyStrings">This bool flag shows if only string inputs were given.</param>
		/// <returns>Returns true if the parseing succeded and false if it fails.</returns>
		public static bool TryToParseValuesFromInputArgument(bool IgnoreHiddenValues, FunctionArgument valueToParse, ParsingContext context, out double parsedValue, out bool theInputContainedOnlyStrings)
		{
			if (ConvertUtil.IsNumeric(valueToParse.Value) && !CellStateHelper.ShouldIgnore(IgnoreHiddenValues, valueToParse, context))
			{
				parsedValue = ConvertUtil.GetValueDouble(valueToParse.Value);
				theInputContainedOnlyStrings = false;
				return true;
			}
			if (valueToParse.Value is string && ConvertUtil.TryParseObjectToDecimal(valueToParse.Value, out double result))
			{
				parsedValue = result;
				theInputContainedOnlyStrings = false;
				return true;
			}
			parsedValue = 0.0;
			theInputContainedOnlyStrings = true;
			return false;
		}

		/// <summary>
		/// Calculates the Variance for a Sample Population.
		/// </summary>
		/// <param name="args">A list of inputs to have their variance calculated.</param>
		/// <returns>The Variance for a Sample population</returns>
		public static double VarianceForASample(List<double> args)
		{
			double avg = args.Average();
			double d = args.Aggregate(0.0, (total, next) => total += System.Math.Pow(next - avg, 2));
			return (d / (args.Count() - 1));
		}

		/// <summary>
		/// Calculates the Variance for a entire Population.
		/// </summary>
		/// <param name="args">A list of inputs to have their variance calculated.</param>
		/// <returns>The Variance for a entire population</returns>
		public static double VarianceForAnEntirePopulation(IEnumerable<double> args)
		{
			double avg = args.Average();
			double d = args.Aggregate(0.0, (total, next) => total += System.Math.Pow(next - avg, 2));
			return (d / (args.Count()));
		}

		/// <summary>
		/// Does the standard deviation on an entire pupulation.
		/// </summary>
		/// <param name="listToDoStandardDeviationOn">This is the list of the entire population to have their standaerd deviation Calulated.</param>
		/// <param name="standardDeviation">This is the calculated standard Deviation.</param>
		/// <returns>Returns true if the standard Deviation succeaded, else false.</returns>
		public static bool TryStandardDeviationEntirePopulation(IEnumerable<double> listToDoStandardDeviationOn, out double standardDeviation)
		{
			standardDeviation = System.Math.Sqrt(StatisticsFunctionHelper.VarianceForAnEntirePopulation(listToDoStandardDeviationOn));
			if (standardDeviation == 0 && listToDoStandardDeviationOn.All(x => x == -1))
				return false;
			return true;
		}

		/// <summary>
		/// Does the standard deviation on an sample pupulation.
		/// </summary>
		/// <param name="listToDoStandardDeviationOn">This is the list of the sample population to have their standaerd deviation Calulated.</param>
		/// <param name="standardDeviation">This is the calculated standard Deviation.</param>
		/// <returns>Returns true if the standard Deviation succeaded, else false.</returns>
		public static bool TryStandardDeviationOnASamplePopulation(List<double> listToDoStandardDeviationOn, out double standardDeviation)
		{
			standardDeviation = System.Math.Sqrt(StatisticsFunctionHelper.VarianceForASample(listToDoStandardDeviationOn));
			if (listToDoStandardDeviationOn.Count() <= 1)
				return false;
			if (standardDeviation == 0 && listToDoStandardDeviationOn.All(x => x == -1))
				return false;
			return true;
		}

		/// <summary>
		/// Checks if the variance for an entire poulation is able to be calcuated.
		/// </summary>
		/// <param name="listOfDoubles">A list of inputs to have their variance calculated.</param>
		/// <param name="variance">The out value of the variance.</param>
		/// <returns>Returns true if it succeads and false if it fails.</returns>
		public static bool TryVarPopulationForAValueErrorCheck(List<double> listOfDoubles, out double variance)
		{
			double avg = listOfDoubles.Average();
			double d = listOfDoubles.Aggregate(0.0, (total, next) => total += System.Math.Pow(next - avg, 2));
			variance = (d / (listOfDoubles.Count()));
			if (variance == 0 && listOfDoubles.All(x => x == -1))
				return false;
			return true;
		}

		/// <summary>
		/// Checks if the variance for an sample poulation is able to be calcuated.
		/// </summary>
		/// <param name="listOfDoubles">A list of inputs to have their variance calculated.</param>
		/// <param name="variance">The out value of the variance.</param>
		/// <returns>Returns true if it succeads and false if it fails.</returns>
		public static bool TryVarSamplePopulationForAValueErrorCheck(IEnumerable<double> listOfDoubles, out double variance)
		{
			double avg = listOfDoubles.Average();
			double d = listOfDoubles.Aggregate(0.0, (total, next) => total += System.Math.Pow(next - avg, 2));
			variance = (d / (listOfDoubles.Count() - 1));
			if (listOfDoubles.Count() <= 1)
				return false;
			if (variance == 0 && listOfDoubles.All(x => x == -1))
				return false;
			return true;
		}

		/// <summary>
		/// Calculates the standard deviation of the specified <paramref name="values"/>.
		/// </summary>
		/// <param name="values">The values to calculate the standard deviation of.</param>
		/// <returns>The standard deviation.</returns>
		public static double Stdev(IEnumerable<double> values)
		{
			// https://stackoverflow.com/questions/3141692/standard-deviation-of-generic-list
			double stdev = 0;
			if (values.Any())
			{
				//Compute the Average
				double avg = values.Average();
				//Perform the Sum of (value-avg)_2_2
				double sum = values.Sum(d => System.Math.Pow(d - avg, 2));
				//Put it all together
				stdev = System.Math.Sqrt((sum) / (values.Count() - 1));
			}
			return stdev;
		}
		#endregion
	}
}
