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
using System.Text.RegularExpressions;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime
{   /// <summary>
	/// TimeStringParser is used in TimeValue. It handles parsing the input and returning a double of just the time.
	/// 
	/// Note: The double that this will return is slightly diffrent from from the one you would get from TimeValue in Excel. the diffrence is that in excel if you do timevalue("26:00") you would get a value that is
	/// equivalent to timevalue("02:00"). This is becuase excel seems to mod the hours if they are greater then 24. In this implamentation that does not happen so you will end up with a whole number. To get the same
	/// functionality you will need to truncate it.
	/// </summary>
	public class TimeStringParser
	{
		private const string RegEx24 = @"[0-9]{1,2}(\:[0-9]{1,2}){0,2}$";

		private double GetSerialNumber(int hour, int minute, int second)
		{

			var secondsInADay = 24d * 60d * 60d;
			return ((double)hour * 60 * 60 + (double)minute * 60 + (double)second) / secondsInADay;
		}

		private void ValidateValues(int hour, int minute, int second)
		{

			if (second < 0 || second > 59)
			{
				throw new FormatException("Illegal value for second: " + second);
			}
			if (minute < 0 || minute > 59)
			{
				throw new FormatException("Illegal value for minute: " + minute);
			}
			
		}


		public virtual double Parse(string input)
		{
			return this.InternalParse(input);
		}

		public virtual bool CanParse(string input)
		{
			System.DateTime dt;
			return Regex.IsMatch(input, TimeStringParser.RegEx24)|| System.DateTime.TryParse(input, out dt);
		}
	
		private double InternalParse(string input)
		{	
			var match = Regex.Match(input, RegEx24);
			if (match.Success)
			{
				return this.Parse24HourTimeString(match.Value);
			}
			System.DateTime dateTime;
			if (System.DateTime.TryParse(input, out dateTime))
			{
				return this.GetSerialNumber(dateTime.Hour, dateTime.Minute, dateTime.Second);
			}
			return -1;
		}
		
		private double Parse24HourTimeString(string input)
		{
			int hour;
			int minute;
			int second;
			this.GetValuesFromString(input, out hour, out minute, out second);
			this.ValidateValues(hour, minute, second);
			return GetSerialNumber(hour, minute, second);
		}

		private void GetValuesFromString(string input, out int hour, out int minute, out int second)
		{
			hour = 0;
			minute = 0;
			second = 0;

			var items = input.Split(':');
			items[items.Length-1] = Regex.Replace(items[items.Length-1], "[^0-9]+$", string.Empty);

			hour = int.Parse(items[0]);
			if (items.Length > 1)
			{
				minute = int.Parse(items[1]);
			}
			if (items.Length > 2)
			{
				second = int.Parse(items[2]);
			}
		}
	}
}
