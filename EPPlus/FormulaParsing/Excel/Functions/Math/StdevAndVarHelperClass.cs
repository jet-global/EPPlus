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
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml.Utils;
using MathObj = System.Math;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	/// <summary>
	/// This is a helper class that is used for all standard deviation and variance methods.
	/// </summary>
	class StdevAndVarHelperClass
	{
		public static bool TryToParseValuesFromInputArgumentByRefrenceOrRange(bool IgnoreHiddenValues, ExcelDataProvider.ICellInfo valueToParse, ParsingContext context, bool includeLogicals, out double parsedValue, out bool theInputContainedOnlyStrings)
		{
			if(includeLogicals)
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
				if (isABoolean)
				{
					if ((bool)valueToParse.Value == true)
						parsedValue = 1;
					else
						parsedValue = 0;
					theInputContainedOnlyStrings = false;
					return true;
				}
				if (ConvertUtil.TryParseDateString(valueToParse.ValueDouble, out System.DateTime dateTime) && ConvertUtil.TryParseDateObjectToOADate(dateTime, out double dateTimeToOADate))
				{
					parsedValue = 0;
					theInputContainedOnlyStrings = false;
					return true;
				}
				parsedValue = 0.0;
				theInputContainedOnlyStrings = false;
				return false;
			}
			else
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
				if (isABoolean)
				{
					parsedValue = valueToParse.ValueDoubleLogical;
					theInputContainedOnlyStrings = false;
					return false;
				}
				if (ConvertUtil.TryParseDateString(valueToParse.ValueDouble, out System.DateTime dateTime) && ConvertUtil.TryParseDateObjectToOADate(dateTime, out double dateTimeToOADAte))
				{
					parsedValue = dateTimeToOADAte;
					theInputContainedOnlyStrings = false;
					return false;
				}
				parsedValue = 0.0;
				theInputContainedOnlyStrings = false;
				return false;
			}
		}

		public static bool TryToParseValuesFromInputArgument(bool ignoreHidden, FunctionArgument valueToParse, ParsingContext context, out double parsedValue, out bool theInputContainedOnlyStrings)
		{
			if (ConvertUtil.IsNumeric(valueToParse.Value) && !CellStateHelper.ShouldIgnore(ignoreHidden, valueToParse, context))
			{
				parsedValue = ConvertUtil.GetValueDouble(valueToParse.Value);
				theInputContainedOnlyStrings = false;
				return true;
			}
			if (valueToParse.Value is string && ConvertUtil.TryParseDateObjectToOADate(valueToParse.Value, out double result))
			{
				parsedValue = result;
				theInputContainedOnlyStrings = false;
				return true;
			}
			parsedValue = 0.0;
			theInputContainedOnlyStrings = true;
			return false;
		}

		public static double VarianceForASample(List<double> args)
		{
			double avg = args.Average();
			double d = args.Aggregate(0.0, (total, next) => total += System.Math.Pow(next - avg, 2));
			return (d / (args.Count() - 1));
		}

		public static double VarianceForAnEntirePopulation(List<double> args)
		{
			double avg = args.Average();
			double d = args.Aggregate(0.0, (total, next) => total += System.Math.Pow(next - avg, 2));
			return (d / (args.Count()));
		}

		public static bool TryStandardDeviationEntirePopulation(List<double> listToDoStandardDeviationOn, out double standardDeviation)
		{
			standardDeviation = MathObj.Sqrt(StdevAndVarHelperClass.VarianceForAnEntirePopulation(listToDoStandardDeviationOn));
			if (standardDeviation == 0 && listToDoStandardDeviationOn.All(x => x == -1))
				return false;
			return true;
		}

		public static bool TryStandardDeviationOnASamplePopulation(List<double> listToDoStandardDeviationOn, out double standardDeviation)
		{
			standardDeviation = MathObj.Sqrt(StdevAndVarHelperClass.VarianceForASample(listToDoStandardDeviationOn));
			if (listToDoStandardDeviationOn.Count() <= 1)
				return false;
			if (standardDeviation == 0 && listToDoStandardDeviationOn.All(x => x == -1))
				return false;
			return true;
		}

		public static bool TryVarPopulationForAValueErrorCheck(List<double> listOfDoubles, out double variance)
		{
			double avg = listOfDoubles.Average();
			double d = listOfDoubles.Aggregate(0.0, (total, next) => total += System.Math.Pow(next - avg, 2));
			variance = (d / (listOfDoubles.Count()));
			if (variance == 0 && listOfDoubles.All(x => x == -1))
				return false;
			return true;
		}

		public static bool TryVarSamplePopulationForAValueErrorCheck(List<double> listOfDoubles, out double variance)
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
	}
}
