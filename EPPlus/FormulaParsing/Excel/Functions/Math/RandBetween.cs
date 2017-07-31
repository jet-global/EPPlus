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
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	/// <summary>
	///Implements the RANDBETWEEN function. 
	/// </summary>
	public class RandBetween : ExcelFunction
	{
		/// <summary>
		/// Get a random number between two given inputs. Both inputs are inclusive.
		/// </summary>
		/// <param name="arguments">This contains upper and lower bounds the user defined.</param>
		/// <param name="context">Unused, this is information about where the function is being executed.</param>
		/// <returns>A random number between the defined upper and lower limits.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentsAreValid(arguments, 2, out eErrorType argumentError) == false)
				return new CompileResult(argumentError);

			var fullOADateOfTheLowInput = ArgToDecimal(arguments, 0);

			var firstArgument = arguments.First();
			var secondArgument = arguments.ElementAt(1);

			ConvertUtil.TryParseObjectToDecimal(firstArgument.Value, out double low);
			ConvertUtil.TryParseObjectToDecimal(secondArgument.Value, out double high);

			if (low > high)
				return CreateResult(eErrorType.Value, DataType.ExcelError);

			var rand = Random.NextDouble();
			var randPart = (this.CalulateDiff(high, low) * rand) + 1;
			randPart = System.Math.Floor(randPart);

			var thisRepresentsTheDateAsAnOADate = System.Math.Truncate(fullOADateOfTheLowInput);
			var todaysOADate = System.Math.Truncate(System.DateTime.Today.ToOADate());

			if (thisRepresentsTheDateAsAnOADate == todaysOADate)
			{
				if (low == 0)
					return CreateResult(0, DataType.Integer);
				else
					return CreateResult(1, DataType.Integer);
			}

			if (firstArgument.Value is bool || secondArgument.Value is bool)
				return CreateResult(eErrorType.Value, DataType.ExcelError);

			return CreateResult(low + randPart, DataType.Integer);
		}

		private double CalulateDiff(double high, double low)
		{
			if (high > 0 && low < 0)
			{
				return high + low * -1;
			}
			else if (high < 0 && low < 0)
			{
				return high * -1 - low * -1;
			}
			return high - low;
		}

		private static System.Random Random { get; } = new System.Random();
	}
}
