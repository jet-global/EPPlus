/*******************************************************************************
* You may amend and distribute as you like, but don't remove this header!
*
* EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
* See http://www.codeplex.com/EPPlus for details.
*
* Copyright (C) 2011-2017 Jan Källman, Matt Delaney, and others as noted in the source history.
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
*  * Author							Change						Date
* *******************************************************************************
* * Mats Alm   		                Added		                2013-12-03
* *******************************************************************************
* For code change notes, see the source control history.
*******************************************************************************/
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	/// <summary>
	/// This class contains the formula for raising one number to the power of another.
	/// </summary>
	public class Power : ExcelFunction
	{
		#region Private Variables
		private object powerCandidate;
		#endregion
		/// <summary>
		/// Takes the user specified arguments and raises the first one to the power of the second one.
		/// </summary>
		/// <param name="arguments">The user specified inputs.</param>
		/// <param name="context">Not used, but needed to override the method.</param>
		/// <returns>The first argument raised to the power of the second argument.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (ValidateArguments(arguments, 1) == false)
				return new CompileResult(eErrorType.Value);
			var number = 0.0;
			var power = 1.0;
			var numberCandidate = arguments.ElementAt(0).Value;

			if (arguments.Count() == 2)
				powerCandidate = arguments.ElementAt(1).Value;
			else
				number = 1;

			if (numberCandidate == null) { }
			else if (powerCandidate == null)
				number = 1;
			else if(numberCandidate is string)
			{
				if (!ConvertUtil.TryParseNumericString(numberCandidate, out double res))
					if (!ConvertUtil.TryParseDateString(numberCandidate, out System.DateTime re))
						return new CompileResult(eErrorType.Value);
					else
					{
						number = this.ArgToDecimal(arguments, 0);
						power = this.ArgToDecimal(arguments, 1);
					}
				else
				{
					number = this.ArgToDecimal(arguments, 0);
					power = this.ArgToDecimal(arguments, 1);
				}
			}
			else if (powerCandidate is string)
			{
				if (!ConvertUtil.TryParseNumericString(powerCandidate, out double res))
					if (!ConvertUtil.TryParseDateString(powerCandidate, out System.DateTime re))
						return new CompileResult(eErrorType.Value);
					else if (ConvertUtil.TryParseDateString(powerCandidate, out System.DateTime ree))
						return new CompileResult(eErrorType.Num);
					else
					{
						number = this.ArgToDecimal(arguments, 0);
						power = this.ArgToDecimal(arguments, 1);
					}
				else
				{
					number = this.ArgToDecimal(arguments, 0);
					power = this.ArgToDecimal(arguments, 1);
				}
			}	
			else if (powerCandidate is System.DateTime)
			{
				return new CompileResult(eErrorType.Num);
			}
			else if (arguments.Count() == 1) { }
			else
			{
				number = this.ArgToDecimal(arguments, 0);
				power = this.ArgToDecimal(arguments, 1);
			}
			
			var resultToReturn = System.Math.Pow(number, power);
			return CreateResult(resultToReturn, DataType.Decimal);
		}
	}
}
