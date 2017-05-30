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
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{
	/// <summary>
	/// This class contains the formula for computing the time based on the user's input. 
	/// </summary>
	public class Time : TimeBaseFunction
	{
		/// <summary>
		/// Execute returns the time as a decimal number. 
		/// </summary>
		/// <param name="arguments">The user's specified hour, minute, and second.</param>
		/// <param name="context">Not used, but needed for overriding the method. </param>
		/// <returns>The time as a double (decimal numer).</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			double dateResult;
			var hour = 0;
			var minute = 0;
			var second = 0;

			if(this.ValidateArguments(arguments, 3))
			{
				if (arguments.ElementAt(0).Value is string && !ConvertUtil.TryParseNumericString(arguments.ElementAt(0).Value, out dateResult))
					return new CompileResult(eErrorType.Value);
				if (arguments.ElementAt(1).Value is string && !ConvertUtil.TryParseNumericString(arguments.ElementAt(1).Value, out dateResult))
					return new CompileResult(eErrorType.Value);
				if (arguments.ElementAt(2).Value is string && !ConvertUtil.TryParseNumericString(arguments.ElementAt(2).Value, out dateResult))
					return new CompileResult(eErrorType.Value);

				hour = this.ArgToInt(arguments, 0);
				minute = this.ArgToInt(arguments, 1);
				second = this.ArgToInt(arguments, 2);
			}
			else
				return new CompileResult(eErrorType.Value);
	
			if (hour < 0)
				return new CompileResult(eErrorType.Num);
			if (hour > 32767 || minute > 32767 || second > 32767)
				return new CompileResult(eErrorType.Num);

			if(hour == 32767 && minute == 32767 && second == 32767)
			{
				//When the maximum input is used in the TIME function it performs all three modifications to the individual
				//parameters, adds them and then performs another calculation if necessary.
				//The link to this information is: https://support.office.com/en-us/article/TIME-function-9a5aff99-8f7d-4611-845e-747d0b8d5457
				//Dealing with the hour being over 23.
				var newHour = hour % 24;
				//Dealing with the minuteute being over 59 and adjusting the hour as such.
				var newMinute = minute % 60;
				var minuteAsHour = minute / 60;
				minuteAsHour = minuteAsHour % 24;
				newHour += minuteAsHour;
				//Dealing with the secondond being over 59 and adjusting the hour and minuteute as such.
				var secondAsHour = (second / 60) / 60;
				var secondAsMinute = second / 60;
				while(secondAsMinute > 59)
					secondAsMinute = secondAsMinute % 60;
				var newSecond = second - ((secondAsHour*60*60) + (secondAsMinute*60));
				//Final calculation to account for the fact that the hour might be over 23.
				hour = (newHour + secondAsHour) % 24;
				minute = newMinute + secondAsMinute;
				second = newSecond;
			}

			if(hour > 23)
				hour = hour % 24;
			if(minute > 59)
			{
				hour = minute / 60;
				minute = minute % 60;
			}
			if(second > 59)
			{
				var newHour = (second / 60) / 60;
				var newMinute = second / 60;
				var hourToSecond = newHour * 60 * 60;
				var minuteToSecond = newMinute * 60;
				second = second - (hourToSecond + minuteToSecond);
				hour = newHour;
				minute = newMinute;
			}

			var secondondsOfThisTime = (double)(hour * 60 * 60 + minute * 60 + second);
			return this.CreateResult(this.GetTimeSerialNumber(secondondsOfThisTime), DataType.Time);
		}
	}
}
