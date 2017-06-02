﻿/*******************************************************************************
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
* For code change notes, see the source control history.
*******************************************************************************/
using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;

namespace EPPlusTest.Excel.Functions
{
	[TestClass]
	public class TimeStringParserTests
	{
		private double GetSerialNumber(int hour, int minute, int second)
		{
			var secondsInADay = 24d * 60d * 60d;
			return ((double)hour * 60 * 60 + (double)minute * 60 + (double)second) / secondsInADay;
		}

		[TestMethod]
		public void CanParseShouldHandleValid24HourPatterns()
		{
			var parser = new TimeStringParser();
			Assert.IsTrue(parser.CanParse("10:12:55"), "Could not parse 10:12:55");
			Assert.IsTrue(parser.CanParse("22:12:55"), "Could not parse 13:12:55");
			Assert.IsTrue(parser.CanParse("13"), "Could not parse 13");
			Assert.IsTrue(parser.CanParse("13:12"), "Could not parse 13:12");
			Assert.IsTrue(parser.CanParse("1/1/2017 13:12"), "Could not parse 1/1/2017 13:12");
			Assert.IsTrue(parser.CanParse("25:00"), "Could not parse 25:00");
			Assert.AreEqual(0.55, Math.Round(parser.Parse("1/1/2017 13:12"), 9));
			Assert.AreEqual(1.041666667, Math.Round(parser.Parse("25:00"),9));
		}

		[TestMethod]
		public void CanParseShouldHandleValid12HourPatterns()
		{
			var parser = new TimeStringParser();
			Assert.IsTrue(parser.CanParse("10:12:55 AM"), "Could not parse 10:12:55 AM");
			Assert.IsTrue(parser.CanParse("9:12:55 PM"), "Could not parse 9:12:55 PM");
			Assert.IsTrue(parser.CanParse("7 AM"), "Could not parse 7 AM");
			Assert.IsTrue(parser.CanParse("4:12 PM"), "Could not parse 4:12 PM");
			Assert.IsTrue(parser.CanParse("1/1/2017 7 AM"), "Could not parse 1/1/2017 7 AM");
			Assert.AreEqual(0.291666667, Math.Round(parser.Parse("1/1/2017 7 AM"), 9));
		}

		[TestMethod]
		public void ParseShouldIdentifyPatternAndReturnCorrectResult()
		{
			var parser = new TimeStringParser();
			var result = parser.Parse("10:12:55");
			Assert.AreEqual(GetSerialNumber(10, 12, 55), result);
		}

		[TestMethod, ExpectedException(typeof(FormatException))]
		public void ParseShouldThrowExceptionIfSecondIsOutOfRange()
		{
			var parser = new TimeStringParser();
			var result = parser.Parse("10:12:60");
		}

		[TestMethod, ExpectedException(typeof(FormatException))]
		public void ParseShouldThrowExceptionIfMinuteIsOutOfRange()
		{
			var parser = new TimeStringParser();
			var result = parser.Parse("10:60:55");
		}

		[TestMethod]
		public void ParseShouldIdentify12HourAMPatternAndReturnCorrectResult()
		{
			var parser = new TimeStringParser();
			var result = parser.Parse("10:12:55 AM");
			Assert.AreEqual(GetSerialNumber(10, 12, 55), result);
		}

		[TestMethod]
		public void ParseShouldIdentify12HourPMPatternAndReturnCorrectResult()
		{
			var parser = new TimeStringParser();
			var result = parser.Parse("10:12:55 PM");
			Assert.AreEqual(GetSerialNumber(22, 12, 55), result);
		}
	}
}
