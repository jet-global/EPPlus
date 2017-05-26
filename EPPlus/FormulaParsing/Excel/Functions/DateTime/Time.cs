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
	public class Time : TimeBaseFunction
	{
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (ValidateArguments(arguments, 1) == false)
				return new CompileResult(eErrorType.Value);

			var firstArg = arguments.ElementAt(0).Value.ToString();
			if (arguments.Count() == 1 && TimeStringParser.CanParse(firstArg))
			{
				var result = TimeStringParser.Parse(firstArg);
				return new CompileResult(result, DataType.Time);
			}

			if (arguments.Count() == 3 && arguments.ElementAt(2).Value == null)
				return new CompileResult(eErrorType.Value);

			var hour = 0;
			var sec = 0;
			var min = 0;

			if (ValidateArguments(arguments, 3) == false)
			{
				hour = ArgToInt(arguments, 0);
				min = ArgToInt(arguments, 1);
				sec = 0;
			}
			else
			{
				hour = ArgToInt(arguments, 0);
				min = ArgToInt(arguments, 1);
				sec = ArgToInt(arguments, 2);
			}
			
			

			if (hour < 0)
				return new CompileResult(eErrorType.Num);
			if (hour > 32767 || min > 32767 || sec > 32767)
				return new CompileResult(eErrorType.Num);

			if(hour == 32767 && min == 32767 && sec == 32767)
			{
				//When the maximum input is used in the TIME function it performs all three modifications to the individual
				//parameters, adds them and then performs another calculation if necessary.
				//Dealing with the hour being over 23
				var newHour = hour % 24;
				//Dealing with the minute being over 59 and adjusting the hour as such
				var newMin = min % 60;
				var minAsHour = min / 60;
				minAsHour = minAsHour % 24;
				newHour += minAsHour;
				//Dealing with the second being over 59 and adjusting the hour and minute as such
				var secAsHour = (sec / 60) / 60;
				var secAsMin = sec / 60;
				while(secAsMin > 59)
					secAsMin = secAsMin % 60;
				var newSec = sec - ((secAsHour*60*60) + (secAsMin*60));

				hour = (newHour + secAsHour) % 24;
				min = newMin + secAsMin;
				sec = newSec;
			}

			if(hour > 23)
				hour = hour % 24;
			if(min > 59)
			{
				hour = min / 60;
				min = min % 60;
			}
			if(sec > 59)
			{
				var Newhour = (sec / 60) / 60;
				var Newmin = sec / 60;
				var hourToSec = Newhour * 60 * 60;
				var minToSec = Newmin * 60;
				sec = sec - (hourToSec + minToSec);
				hour = Newhour;
				min = Newmin;
			}

			var secondsOfThisTime = (double)(hour * 60 * 60 + min * 60 + sec);
			return CreateResult(GetTimeSerialNumber(secondsOfThisTime), DataType.Time);
		}
	}
}
