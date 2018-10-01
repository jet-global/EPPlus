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
* For code change notes, see the source control history.
*******************************************************************************/
using System;
using System.Globalization;
using System.IO;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace EPPlusTest.Excel
{
	[TestClass]
	public class OperatorsTests
	{
		#region Logical Comparison Operator Tests
		[TestMethod]
		public void OperatorLogicalOperatorsShouldCorrectlyCompareIntegersAndText()
		{
			// Numbers are always strictly less than text.
			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult(1000000, DataType.Integer), new CompileResult("Numbers are always less than text.", DataType.String)).Result);
			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult(1000000, DataType.Integer), new CompileResult("1", DataType.String)).Result);
			Assert.AreEqual(false, Operator.GreaterThanOrEqual.Apply(new CompileResult(1000000, DataType.Integer), new CompileResult("1", DataType.String)).Result);
			Assert.AreEqual(true, Operator.LessThan.Apply(new CompileResult(1000000, DataType.Integer), new CompileResult("1", DataType.String)).Result);
			Assert.AreEqual(true, Operator.LessThanOrEqual.Apply(new CompileResult(1000000, DataType.Integer), new CompileResult("1", DataType.String)).Result);
			Assert.AreEqual(false, Operator.EqualsTo.Apply(new CompileResult(1, DataType.Integer), new CompileResult("1", DataType.String)).Result);
			Assert.AreEqual(true, Operator.NotEqualsTo.Apply(new CompileResult(1, DataType.Integer), new CompileResult("1", DataType.String)).Result);

			Assert.AreEqual(true, Operator.GreaterThan.Apply(new CompileResult("Text is greater than numbers.", DataType.String), new CompileResult(1000000, DataType.Integer)).Result);
			Assert.AreEqual(true, Operator.GreaterThan.Apply(new CompileResult("1", DataType.String), new CompileResult(1000000, DataType.Integer)).Result);
			Assert.AreEqual(true, Operator.GreaterThanOrEqual.Apply(new CompileResult("1", DataType.String), new CompileResult(1000000, DataType.Integer)).Result);
			Assert.AreEqual(false, Operator.LessThan.Apply(new CompileResult("1", DataType.String), new CompileResult(1000000, DataType.Integer)).Result);
			Assert.AreEqual(false, Operator.LessThanOrEqual.Apply(new CompileResult("1", DataType.String), new CompileResult(1000000, DataType.Integer)).Result);
		}

		[TestMethod]
		public void OperatorLogicalOperatorsShouldCorrectlyCompareDecimalsAndText()
		{
			// Numbers are always strictly less than text.
			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult(1000000.0d, DataType.Decimal), new CompileResult("Numbers are always less than text.", DataType.String)).Result);
			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult(1000000.0d, DataType.Decimal), new CompileResult("1", DataType.String)).Result);
			Assert.AreEqual(false, Operator.GreaterThanOrEqual.Apply(new CompileResult(1000000.0d, DataType.Decimal), new CompileResult("1", DataType.String)).Result);
			Assert.AreEqual(true, Operator.LessThan.Apply(new CompileResult(1000000.0d, DataType.Decimal), new CompileResult("1", DataType.String)).Result);
			Assert.AreEqual(true, Operator.LessThanOrEqual.Apply(new CompileResult(1000000.0d, DataType.Decimal), new CompileResult("1", DataType.String)).Result);
			Assert.AreEqual(false, Operator.EqualsTo.Apply(new CompileResult(1, DataType.Decimal), new CompileResult("1", DataType.String)).Result);
			Assert.AreEqual(true, Operator.NotEqualsTo.Apply(new CompileResult(1, DataType.Decimal), new CompileResult("1", DataType.String)).Result);

			Assert.AreEqual(true, Operator.GreaterThan.Apply(new CompileResult("Text is greater than numbers.", DataType.String), new CompileResult(1000000.0d, DataType.Decimal)).Result);
			Assert.AreEqual(true, Operator.GreaterThan.Apply(new CompileResult("1", DataType.String), new CompileResult(1000000.0d, DataType.Decimal)).Result);
			Assert.AreEqual(true, Operator.GreaterThanOrEqual.Apply(new CompileResult("1", DataType.String), new CompileResult(1000000.0d, DataType.Decimal)).Result);
			Assert.AreEqual(false, Operator.LessThan.Apply(new CompileResult("1", DataType.String), new CompileResult(1000000.0d, DataType.Decimal)).Result);
			Assert.AreEqual(false, Operator.LessThanOrEqual.Apply(new CompileResult("1", DataType.String), new CompileResult(1000000.0d, DataType.Decimal)).Result);
		}

		[TestMethod]
		public void OperatorLogicalOperatorsShouldCorrectlyCompareNumbersAndLogicalValues()
		{
			// Logical values are strictly larger than all numeric values.
			Assert.AreEqual(true, Operator.GreaterThan.Apply(new CompileResult(true, DataType.Boolean), new CompileResult(int.MaxValue, DataType.Integer)).Result);
			Assert.AreEqual(true, Operator.GreaterThan.Apply(new CompileResult(false, DataType.Boolean), new CompileResult(int.MaxValue, DataType.Integer)).Result);
			Assert.AreEqual(true, Operator.GreaterThan.Apply(new CompileResult(true, DataType.Boolean), new CompileResult(int.MinValue, DataType.Integer)).Result);
			Assert.AreEqual(true, Operator.GreaterThan.Apply(new CompileResult(false, DataType.Boolean), new CompileResult(int.MinValue, DataType.Integer)).Result);
			Assert.AreEqual(true, Operator.GreaterThanOrEqual.Apply(new CompileResult(false, DataType.Boolean), new CompileResult(int.MaxValue, DataType.Integer)).Result);
			Assert.AreEqual(false, Operator.LessThan.Apply(new CompileResult(true, DataType.Boolean), new CompileResult(int.MinValue, DataType.Integer)).Result);
			Assert.AreEqual(false, Operator.LessThan.Apply(new CompileResult(false, DataType.Boolean), new CompileResult(int.MinValue, DataType.Integer)).Result);
			Assert.AreEqual(false, Operator.LessThanOrEqual.Apply(new CompileResult(true, DataType.Boolean), new CompileResult(int.MinValue, DataType.Integer)).Result);
			Assert.AreEqual(false, Operator.EqualsTo.Apply(new CompileResult(false, DataType.Boolean), new CompileResult(0.0, DataType.Decimal)).Result);
			Assert.AreEqual(false, Operator.EqualsTo.Apply(new CompileResult(true, DataType.Boolean), new CompileResult(1.0, DataType.Decimal)).Result);
			Assert.AreEqual(true, Operator.NotEqualsTo.Apply(new CompileResult(false, DataType.Boolean), new CompileResult(0.0, DataType.Decimal)).Result);
			Assert.AreEqual(true, Operator.NotEqualsTo.Apply(new CompileResult(true, DataType.Boolean), new CompileResult(1.0, DataType.Decimal)).Result);

			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult(int.MaxValue, DataType.Integer), new CompileResult(true, DataType.Boolean)).Result);
			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult(int.MaxValue, DataType.Integer), new CompileResult(false, DataType.Boolean)).Result);
			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult(int.MinValue, DataType.Integer), new CompileResult(true, DataType.Boolean)).Result);
			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult(int.MinValue, DataType.Integer), new CompileResult(false, DataType.Boolean)).Result);
			Assert.AreEqual(false, Operator.GreaterThanOrEqual.Apply(new CompileResult(int.MaxValue, DataType.Integer), new CompileResult(false, DataType.Boolean)).Result);
			Assert.AreEqual(true, Operator.LessThan.Apply(new CompileResult(int.MinValue, DataType.Integer), new CompileResult(true, DataType.Boolean)).Result);
			Assert.AreEqual(true, Operator.LessThan.Apply(new CompileResult(int.MinValue, DataType.Integer), new CompileResult(false, DataType.Boolean)).Result);
			Assert.AreEqual(true, Operator.LessThanOrEqual.Apply(new CompileResult(int.MinValue, DataType.Integer), new CompileResult(true, DataType.Boolean)).Result);
		}

		[TestMethod]
		public void OperatorLogicalOperatorsShouldCorrectlyCompareTextAndLogicalValues()
		{
			// Logical values are always strictly greater than text.
			Assert.AreEqual(true, Operator.GreaterThan.Apply(new CompileResult(true, DataType.Boolean), new CompileResult(string.Empty, DataType.String)).Result);
			Assert.AreEqual(true, Operator.GreaterThan.Apply(new CompileResult(false, DataType.Boolean), new CompileResult(string.Empty, DataType.String)).Result);
			Assert.AreEqual(true, Operator.GreaterThan.Apply(new CompileResult(true, DataType.Boolean), new CompileResult("We are confident in this test because W > T", DataType.String)).Result);
			Assert.AreEqual(true, Operator.GreaterThan.Apply(new CompileResult(false, DataType.Boolean), new CompileResult("Gotta start with a letter bigger than F to be confident in this test", DataType.String)).Result);
			Assert.AreEqual(true, Operator.GreaterThanOrEqual.Apply(new CompileResult(false, DataType.Boolean), new CompileResult("ZZZZ huge string", DataType.String)).Result);
			Assert.AreEqual(false, Operator.LessThan.Apply(new CompileResult(true, DataType.Boolean), new CompileResult("arbitrary text", DataType.String)).Result);
			Assert.AreEqual(false, Operator.LessThan.Apply(new CompileResult(false, DataType.Boolean), new CompileResult("arbitrary text", DataType.String)).Result);
			Assert.AreEqual(false, Operator.LessThanOrEqual.Apply(new CompileResult(true, DataType.Boolean), new CompileResult(int.MinValue, DataType.String)).Result);
			Assert.AreEqual(false, Operator.EqualsTo.Apply(new CompileResult(false, DataType.Boolean), new CompileResult("FALSE", DataType.String)).Result);
			Assert.AreEqual(false, Operator.EqualsTo.Apply(new CompileResult(true, DataType.Boolean), new CompileResult("TRUE", DataType.String)).Result);
			Assert.AreEqual(true, Operator.NotEqualsTo.Apply(new CompileResult(false, DataType.Boolean), new CompileResult("FALSE", DataType.String)).Result);
			Assert.AreEqual(true, Operator.NotEqualsTo.Apply(new CompileResult(true, DataType.Boolean), new CompileResult("TRUE", DataType.String)).Result);

			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult(string.Empty, DataType.String), new CompileResult(true, DataType.Boolean)).Result);
			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult(string.Empty, DataType.String), new CompileResult(false, DataType.Boolean)).Result);
			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult("We start with W", DataType.String), new CompileResult(true, DataType.Boolean)).Result);
			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult("Might start with M", DataType.String), new CompileResult(false, DataType.Boolean)).Result);
			Assert.AreEqual(false, Operator.GreaterThanOrEqual.Apply(new CompileResult("Z is a pretty big string", DataType.String), new CompileResult(false, DataType.Boolean)).Result);
			Assert.AreEqual(true, Operator.LessThan.Apply(new CompileResult("Zzzz... Text is always less than logical values", DataType.String), new CompileResult(true, DataType.Boolean)).Result);
			Assert.AreEqual(true, Operator.LessThan.Apply(new CompileResult("Text", DataType.String), new CompileResult(false, DataType.Boolean)).Result);
			Assert.AreEqual(true, Operator.LessThanOrEqual.Apply(new CompileResult("Zzzz", DataType.String), new CompileResult(true, DataType.Boolean)).Result);
		}

		[TestMethod]
		public void OperatorLogicalOperatorsShouldCorrectlyCompareDecimalValues()
		{
			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult(1000000.0, DataType.Decimal), new CompileResult(1000000.01, DataType.Decimal)).Result);
			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult(-1000000, DataType.Decimal), new CompileResult(0, DataType.Decimal)).Result);
			Assert.AreEqual(true, Operator.LessThan.Apply(new CompileResult(100.0, DataType.Decimal), new CompileResult(101.0, DataType.Decimal)).Result);
			Assert.AreEqual(true, Operator.LessThanOrEqual.Apply(new CompileResult(100.0, DataType.Decimal), new CompileResult(101.0, DataType.Decimal)).Result);
			Assert.AreEqual(true, Operator.LessThanOrEqual.Apply(new CompileResult(100.0, DataType.Decimal), new CompileResult(100.0, DataType.Decimal)).Result);
			Assert.AreEqual(false, Operator.EqualsTo.Apply(new CompileResult(1.1, DataType.Decimal), new CompileResult(1.0, DataType.Decimal)).Result);
			Assert.AreEqual(true, Operator.EqualsTo.Apply(new CompileResult(1.0, DataType.Decimal), new CompileResult(1.0, DataType.Decimal)).Result);
			Assert.AreEqual(true, Operator.NotEqualsTo.Apply(new CompileResult(1.0, DataType.Decimal), new CompileResult(1.1, DataType.Decimal)).Result);

			Assert.AreEqual(true, Operator.GreaterThan.Apply(new CompileResult(1000000.01, DataType.Decimal), new CompileResult(1000000.0, DataType.Decimal)).Result);
			Assert.AreEqual(true, Operator.GreaterThan.Apply(new CompileResult(-1.0, DataType.Decimal), new CompileResult(-100000.0, DataType.Decimal)).Result);
			Assert.AreEqual(true, Operator.GreaterThanOrEqual.Apply(new CompileResult(1000000.1, DataType.Decimal), new CompileResult(1000000.0, DataType.Decimal)).Result);
			Assert.AreEqual(true, Operator.GreaterThanOrEqual.Apply(new CompileResult(1000000.0, DataType.Decimal), new CompileResult(1000000.0, DataType.Decimal)).Result);
			Assert.AreEqual(false, Operator.LessThan.Apply(new CompileResult(100000, DataType.Decimal), new CompileResult(1000, DataType.Decimal)).Result);
			Assert.AreEqual(false, Operator.LessThanOrEqual.Apply(new CompileResult(10.0, DataType.Decimal), new CompileResult(9.9, DataType.Decimal)).Result);
		}

		[TestMethod]
		public void OperatorLogicalOperatorsShouldCorrectlyCompareStrings()
		{
			// Text comparison operators are case-insensitive.
			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult("A", DataType.String), new CompileResult("B", DataType.String)).Result);
			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult("a", DataType.String), new CompileResult("B", DataType.String)).Result);
			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult("\"", DataType.String), new CompileResult("a", DataType.String)).Result);
			Assert.AreEqual(true, Operator.LessThan.Apply(new CompileResult("A", DataType.String), new CompileResult("b", DataType.String)).Result);
			Assert.AreEqual(true, Operator.LessThan.Apply(new CompileResult("a", DataType.String), new CompileResult("B", DataType.String)).Result);
			Assert.AreEqual(true, Operator.LessThanOrEqual.Apply(new CompileResult("A", DataType.String), new CompileResult("B", DataType.String)).Result);
			Assert.AreEqual(true, Operator.LessThanOrEqual.Apply(new CompileResult("aaa", DataType.String), new CompileResult("B", DataType.String)).Result);
			Assert.AreEqual(true, Operator.EqualsTo.Apply(new CompileResult("abcde", DataType.String), new CompileResult("AbCdE", DataType.String)).Result);
			Assert.AreEqual(false, Operator.NotEqualsTo.Apply(new CompileResult("abcde", DataType.String), new CompileResult("abcde", DataType.String)).Result);

			Assert.AreEqual(true, Operator.GreaterThan.Apply(new CompileResult("B", DataType.String), new CompileResult("A", DataType.String)).Result);
			Assert.AreEqual(true, Operator.GreaterThan.Apply(new CompileResult("The first character that doesn't match is controlling", DataType.String), new CompileResult("The first character is smaller here", DataType.String)).Result);
			Assert.AreEqual(true, Operator.GreaterThanOrEqual.Apply(new CompileResult("Dogs", DataType.String), new CompileResult("Cats", DataType.String)).Result);
			Assert.AreEqual(true, Operator.GreaterThanOrEqual.Apply(new CompileResult("dogs", DataType.String), new CompileResult("Dogs", DataType.String)).Result);
			Assert.AreEqual(false, Operator.LessThan.Apply(new CompileResult("Dogs", DataType.String), new CompileResult("Cats", DataType.String)).Result);
			Assert.AreEqual(false, Operator.LessThanOrEqual.Apply(new CompileResult("Dogs", DataType.String), new CompileResult("Cheetahs", DataType.String)).Result);
		}

		[TestMethod]
		public void OperatorLogicalOperatorsShouldCorrectlyCompareDatesAndTimes()
		{
			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult(.89, DataType.Time), new CompileResult(43306, DataType.Date)).Result);
			Assert.AreEqual(true, Operator.GreaterThan.Apply(new CompileResult(43306, DataType.Date), new CompileResult(.89, DataType.Time)).Result);
			Assert.AreEqual(true, Operator.GreaterThan.Apply(new CompileResult(.89, DataType.Time), new CompileResult(.74, DataType.Date)).Result);
			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult(.74, DataType.Time), new CompileResult(.89, DataType.Date)).Result);
			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult(.89, DataType.Time), new CompileResult(.89, DataType.Date)).Result);
			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult(43306, DataType.Date), new CompileResult(43399, DataType.Date)).Result);
			Assert.AreEqual(true, Operator.GreaterThan.Apply(new CompileResult(43399, DataType.Date), new CompileResult(43306, DataType.Date)).Result);
			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult(43306, DataType.Date), new CompileResult(43306, DataType.Date)).Result);

			Assert.AreEqual(false, Operator.GreaterThanOrEqual.Apply(new CompileResult(.89, DataType.Time), new CompileResult(43306, DataType.Date)).Result);
			Assert.AreEqual(true, Operator.GreaterThanOrEqual.Apply(new CompileResult(43306, DataType.Date), new CompileResult(.89, DataType.Time)).Result);
			Assert.AreEqual(true, Operator.GreaterThanOrEqual.Apply(new CompileResult(.89, DataType.Time), new CompileResult(.74, DataType.Date)).Result);
			Assert.AreEqual(false, Operator.GreaterThanOrEqual.Apply(new CompileResult(.74, DataType.Time), new CompileResult(.89, DataType.Date)).Result);
			Assert.AreEqual(true, Operator.GreaterThanOrEqual.Apply(new CompileResult(.89, DataType.Time), new CompileResult(.89, DataType.Date)).Result);
			Assert.AreEqual(false, Operator.GreaterThanOrEqual.Apply(new CompileResult(43306, DataType.Date), new CompileResult(43399, DataType.Date)).Result);
			Assert.AreEqual(true, Operator.GreaterThanOrEqual.Apply(new CompileResult(43399, DataType.Date), new CompileResult(43306, DataType.Date)).Result);
			Assert.AreEqual(true, Operator.GreaterThanOrEqual.Apply(new CompileResult(43306, DataType.Date), new CompileResult(43306, DataType.Date)).Result);

			Assert.AreEqual(true, Operator.LessThan.Apply(new CompileResult(.89, DataType.Time), new CompileResult(43306, DataType.Date)).Result);
			Assert.AreEqual(false, Operator.LessThan.Apply(new CompileResult(43306, DataType.Date), new CompileResult(.89, DataType.Time)).Result);
			Assert.AreEqual(false, Operator.LessThan.Apply(new CompileResult(.89, DataType.Time), new CompileResult(.74, DataType.Date)).Result);
			Assert.AreEqual(true, Operator.LessThan.Apply(new CompileResult(.74, DataType.Time), new CompileResult(.89, DataType.Date)).Result);
			Assert.AreEqual(false, Operator.LessThan.Apply(new CompileResult(.89, DataType.Time), new CompileResult(.89, DataType.Date)).Result);
			Assert.AreEqual(true, Operator.LessThan.Apply(new CompileResult(43306, DataType.Date), new CompileResult(43399, DataType.Date)).Result);
			Assert.AreEqual(false, Operator.LessThan.Apply(new CompileResult(43399, DataType.Date), new CompileResult(43306, DataType.Date)).Result);
			Assert.AreEqual(false, Operator.LessThan.Apply(new CompileResult(43306, DataType.Date), new CompileResult(43306, DataType.Date)).Result);

			Assert.AreEqual(true, Operator.LessThanOrEqual.Apply(new CompileResult(.89, DataType.Time), new CompileResult(43306, DataType.Date)).Result);
			Assert.AreEqual(false, Operator.LessThanOrEqual.Apply(new CompileResult(43306, DataType.Date), new CompileResult(.89, DataType.Time)).Result);
			Assert.AreEqual(false, Operator.LessThanOrEqual.Apply(new CompileResult(.89, DataType.Time), new CompileResult(.74, DataType.Date)).Result);
			Assert.AreEqual(true, Operator.LessThanOrEqual.Apply(new CompileResult(.74, DataType.Time), new CompileResult(.89, DataType.Date)).Result);
			Assert.AreEqual(true, Operator.LessThanOrEqual.Apply(new CompileResult(.89, DataType.Time), new CompileResult(.89, DataType.Date)).Result);
			Assert.AreEqual(true, Operator.LessThanOrEqual.Apply(new CompileResult(43306, DataType.Date), new CompileResult(43399, DataType.Date)).Result);
			Assert.AreEqual(false, Operator.LessThanOrEqual.Apply(new CompileResult(43399, DataType.Date), new CompileResult(43306, DataType.Date)).Result);
			Assert.AreEqual(true, Operator.LessThanOrEqual.Apply(new CompileResult(43306, DataType.Date), new CompileResult(43306, DataType.Date)).Result);

			Assert.AreEqual(false, Operator.EqualsTo.Apply(new CompileResult(.89, DataType.Time), new CompileResult(43306, DataType.Date)).Result);
			Assert.AreEqual(false, Operator.EqualsTo.Apply(new CompileResult(43306, DataType.Date), new CompileResult(.89, DataType.Time)).Result);
			Assert.AreEqual(false, Operator.EqualsTo.Apply(new CompileResult(.89, DataType.Time), new CompileResult(.74, DataType.Date)).Result);
			Assert.AreEqual(false, Operator.EqualsTo.Apply(new CompileResult(.74, DataType.Time), new CompileResult(.89, DataType.Date)).Result);
			Assert.AreEqual(true, Operator.EqualsTo.Apply(new CompileResult(.89, DataType.Time), new CompileResult(.89, DataType.Date)).Result);
			Assert.AreEqual(false, Operator.EqualsTo.Apply(new CompileResult(43306, DataType.Date), new CompileResult(43399, DataType.Date)).Result);
			Assert.AreEqual(false, Operator.EqualsTo.Apply(new CompileResult(43399, DataType.Date), new CompileResult(43306, DataType.Date)).Result);
			Assert.AreEqual(true, Operator.EqualsTo.Apply(new CompileResult(43306, DataType.Date), new CompileResult(43306, DataType.Date)).Result);

			Assert.AreEqual(true, Operator.NotEqualsTo.Apply(new CompileResult(.89, DataType.Time), new CompileResult(43306, DataType.Date)).Result);
			Assert.AreEqual(true, Operator.NotEqualsTo.Apply(new CompileResult(43306, DataType.Date), new CompileResult(.89, DataType.Time)).Result);
			Assert.AreEqual(true, Operator.NotEqualsTo.Apply(new CompileResult(.89, DataType.Time), new CompileResult(.74, DataType.Date)).Result);
			Assert.AreEqual(true, Operator.NotEqualsTo.Apply(new CompileResult(.74, DataType.Time), new CompileResult(.89, DataType.Date)).Result);
			Assert.AreEqual(false, Operator.NotEqualsTo.Apply(new CompileResult(.89, DataType.Time), new CompileResult(.89, DataType.Date)).Result);
			Assert.AreEqual(true, Operator.NotEqualsTo.Apply(new CompileResult(43306, DataType.Date), new CompileResult(43399, DataType.Date)).Result);
			Assert.AreEqual(true, Operator.NotEqualsTo.Apply(new CompileResult(43399, DataType.Date), new CompileResult(43306, DataType.Date)).Result);
			Assert.AreEqual(false, Operator.NotEqualsTo.Apply(new CompileResult(43306, DataType.Date), new CompileResult(43306, DataType.Date)).Result);
		}

		[TestMethod]
		public void OperatorLogicalOperatorsShouldCorrectlyCompareLogicalValues()
		{
			// TRUE is strictly greater than FALSE.
			Assert.AreEqual(true, Operator.GreaterThan.Apply(new CompileResult(true, DataType.Boolean), new CompileResult(false, DataType.Boolean)).Result);
			Assert.AreEqual(true, Operator.GreaterThanOrEqual.Apply(new CompileResult(true, DataType.Boolean), new CompileResult(false, DataType.Boolean)).Result);
			Assert.AreEqual(false, Operator.EqualsTo.Apply(new CompileResult(true, DataType.Boolean), new CompileResult(false, DataType.Boolean)).Result);
			Assert.AreEqual(true, Operator.NotEqualsTo.Apply(new CompileResult(true, DataType.Boolean), new CompileResult(false, DataType.Boolean)).Result);
			Assert.AreEqual(true, Operator.EqualsTo.Apply(new CompileResult(true, DataType.Boolean), new CompileResult(true, DataType.Boolean)).Result);
			Assert.AreEqual(true, Operator.GreaterThanOrEqual.Apply(new CompileResult(true, DataType.Boolean), new CompileResult(true, DataType.Boolean)).Result);
			Assert.AreEqual(true, Operator.LessThanOrEqual.Apply(new CompileResult(true, DataType.Boolean), new CompileResult(true, DataType.Boolean)).Result);
			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult(true, DataType.Boolean), new CompileResult(true, DataType.Boolean)).Result);
			Assert.AreEqual(false, Operator.LessThan.Apply(new CompileResult(true, DataType.Boolean), new CompileResult(true, DataType.Boolean)).Result);
			Assert.AreEqual(true, Operator.EqualsTo.Apply(new CompileResult(false, DataType.Boolean), new CompileResult(false, DataType.Boolean)).Result);
			Assert.AreEqual(true, Operator.GreaterThanOrEqual.Apply(new CompileResult(false, DataType.Boolean), new CompileResult(false, DataType.Boolean)).Result);
			Assert.AreEqual(true, Operator.LessThanOrEqual.Apply(new CompileResult(false, DataType.Boolean), new CompileResult(false, DataType.Boolean)).Result);
			Assert.AreEqual(false, Operator.LessThan.Apply(new CompileResult(false, DataType.Boolean), new CompileResult(false, DataType.Boolean)).Result);
			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult(false, DataType.Boolean), new CompileResult(false, DataType.Boolean)).Result);
			Assert.AreEqual(false, Operator.LessThan.Apply(new CompileResult(true, DataType.Boolean), new CompileResult(false, DataType.Boolean)).Result);
			Assert.AreEqual(false, Operator.LessThanOrEqual.Apply(new CompileResult(true, DataType.Boolean), new CompileResult(false, DataType.Boolean)).Result);
			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult(false, DataType.Boolean), new CompileResult(true, DataType.Boolean)).Result);
			Assert.AreEqual(false, Operator.GreaterThanOrEqual.Apply(new CompileResult(false, DataType.Boolean), new CompileResult(true, DataType.Boolean)).Result);
			Assert.AreEqual(true, Operator.LessThan.Apply(new CompileResult(false, DataType.Boolean), new CompileResult(true, DataType.Boolean)).Result);
			Assert.AreEqual(true, Operator.LessThanOrEqual.Apply(new CompileResult(false, DataType.Boolean), new CompileResult(true, DataType.Boolean)).Result);
		}

		[TestMethod]
		public void OperatorLogicalOperatorsShouldCorrectlyCompareLogicalAndEmptyValues()
		{
			Assert.AreEqual(true, Operator.GreaterThan.Apply(new CompileResult(true, DataType.Boolean), CompileResult.Empty).Result);
			Assert.AreEqual(false, Operator.GreaterThan.Apply(CompileResult.Empty, new CompileResult(true, DataType.Boolean)).Result);
			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult(false, DataType.Boolean), CompileResult.Empty).Result);
			Assert.AreEqual(false, Operator.GreaterThan.Apply(CompileResult.Empty, new CompileResult(false, DataType.Boolean)).Result);

			Assert.AreEqual(true, Operator.GreaterThanOrEqual.Apply(new CompileResult(true, DataType.Boolean), CompileResult.Empty).Result);
			Assert.AreEqual(false, Operator.GreaterThanOrEqual.Apply(CompileResult.Empty, new CompileResult(true, DataType.Boolean)).Result);
			Assert.AreEqual(true, Operator.GreaterThanOrEqual.Apply(new CompileResult(false, DataType.Boolean), CompileResult.Empty).Result);
			Assert.AreEqual(true, Operator.GreaterThanOrEqual.Apply(CompileResult.Empty, new CompileResult(false, DataType.Boolean)).Result);

			Assert.AreEqual(false, Operator.EqualsTo.Apply(new CompileResult(true, DataType.Boolean), CompileResult.Empty).Result);
			Assert.AreEqual(false, Operator.EqualsTo.Apply(CompileResult.Empty, new CompileResult(true, DataType.Boolean)).Result);
			Assert.AreEqual(true, Operator.EqualsTo.Apply(new CompileResult(false, DataType.Boolean), CompileResult.Empty).Result);
			Assert.AreEqual(true, Operator.EqualsTo.Apply(CompileResult.Empty, new CompileResult(false, DataType.Boolean)).Result);

			Assert.AreEqual(true, Operator.NotEqualsTo.Apply(new CompileResult(true, DataType.Boolean), CompileResult.Empty).Result);
			Assert.AreEqual(true, Operator.NotEqualsTo.Apply(CompileResult.Empty, new CompileResult(true, DataType.Boolean)).Result);
			Assert.AreEqual(false, Operator.NotEqualsTo.Apply(new CompileResult(false, DataType.Boolean), CompileResult.Empty).Result);
			Assert.AreEqual(false, Operator.NotEqualsTo.Apply(CompileResult.Empty, new CompileResult(false, DataType.Boolean)).Result);

			Assert.AreEqual(false, Operator.LessThan.Apply(new CompileResult(true, DataType.Boolean), CompileResult.Empty).Result);
			Assert.AreEqual(true, Operator.LessThan.Apply(CompileResult.Empty, new CompileResult(true, DataType.Boolean)).Result);
			Assert.AreEqual(false, Operator.LessThan.Apply(new CompileResult(false, DataType.Boolean), CompileResult.Empty).Result);
			Assert.AreEqual(false, Operator.LessThan.Apply(CompileResult.Empty, new CompileResult(false, DataType.Boolean)).Result);

			Assert.AreEqual(false, Operator.LessThanOrEqual.Apply(new CompileResult(true, DataType.Boolean), CompileResult.Empty).Result);
			Assert.AreEqual(true, Operator.LessThanOrEqual.Apply(CompileResult.Empty, new CompileResult(true, DataType.Boolean)).Result);
			Assert.AreEqual(true, Operator.LessThanOrEqual.Apply(new CompileResult(false, DataType.Boolean), CompileResult.Empty).Result);
			Assert.AreEqual(true, Operator.LessThanOrEqual.Apply(CompileResult.Empty, new CompileResult(false, DataType.Boolean)).Result);
		}

		[TestMethod]
		public void OperatorLogicalOperatorsShouldCorrectlyCompareNumericAndEmptyValues()
		{
			Assert.AreEqual(true, Operator.GreaterThan.Apply(new CompileResult(1d, DataType.Decimal), CompileResult.Empty).Result);
			Assert.AreEqual(false, Operator.GreaterThan.Apply(CompileResult.Empty, new CompileResult(1d, DataType.Decimal)).Result);
			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult(-1d, DataType.Decimal), CompileResult.Empty).Result);
			Assert.AreEqual(true, Operator.GreaterThan.Apply(CompileResult.Empty, new CompileResult(-1d, DataType.Decimal)).Result);

			Assert.AreEqual(true, Operator.GreaterThanOrEqual.Apply(new CompileResult(1d, DataType.Decimal), CompileResult.Empty).Result);
			Assert.AreEqual(false, Operator.GreaterThanOrEqual.Apply(CompileResult.Empty, new CompileResult(1d, DataType.Decimal)).Result);
			Assert.AreEqual(false, Operator.GreaterThanOrEqual.Apply(new CompileResult(-1d, DataType.Decimal), CompileResult.Empty).Result);
			Assert.AreEqual(true, Operator.GreaterThanOrEqual.Apply(CompileResult.Empty, new CompileResult(-1d, DataType.Decimal)).Result);

			Assert.AreEqual(false, Operator.EqualsTo.Apply(new CompileResult(1d, DataType.Decimal), CompileResult.Empty).Result);
			Assert.AreEqual(false, Operator.EqualsTo.Apply(CompileResult.Empty, new CompileResult(1d, DataType.Decimal)).Result);
			Assert.AreEqual(false, Operator.EqualsTo.Apply(new CompileResult(-1d, DataType.Decimal), CompileResult.Empty).Result);
			Assert.AreEqual(false, Operator.EqualsTo.Apply(CompileResult.Empty, new CompileResult(-1d, DataType.Decimal)).Result);

			Assert.AreEqual(true, Operator.NotEqualsTo.Apply(new CompileResult(1d, DataType.Decimal), CompileResult.Empty).Result);
			Assert.AreEqual(true, Operator.NotEqualsTo.Apply(CompileResult.Empty, new CompileResult(1d, DataType.Decimal)).Result);
			Assert.AreEqual(true, Operator.NotEqualsTo.Apply(new CompileResult(-1d, DataType.Decimal), CompileResult.Empty).Result);
			Assert.AreEqual(true, Operator.NotEqualsTo.Apply(CompileResult.Empty, new CompileResult(-1d, DataType.Decimal)).Result);

			Assert.AreEqual(false, Operator.LessThan.Apply(new CompileResult(1d, DataType.Decimal), CompileResult.Empty).Result);
			Assert.AreEqual(true, Operator.LessThan.Apply(CompileResult.Empty, new CompileResult(1d, DataType.Decimal)).Result);
			Assert.AreEqual(true, Operator.LessThan.Apply(new CompileResult(-1d, DataType.Decimal), CompileResult.Empty).Result);
			Assert.AreEqual(false, Operator.LessThan.Apply(CompileResult.Empty, new CompileResult(-1d, DataType.Decimal)).Result);

			Assert.AreEqual(false, Operator.LessThanOrEqual.Apply(new CompileResult(1d, DataType.Decimal), CompileResult.Empty).Result);
			Assert.AreEqual(true, Operator.LessThanOrEqual.Apply(CompileResult.Empty, new CompileResult(1d, DataType.Decimal)).Result);
			Assert.AreEqual(true, Operator.LessThanOrEqual.Apply(new CompileResult(-1d, DataType.Decimal), CompileResult.Empty).Result);
			Assert.AreEqual(false, Operator.LessThanOrEqual.Apply(CompileResult.Empty, new CompileResult(-1d, DataType.Decimal)).Result);
		}

		[TestMethod]
		public void OperatorLogicalOperatorsShouldCorrectlyCompareStringAndEmptyValues()
		{
			const string value = "sdflkjsdlfk";
			Assert.AreEqual(true, Operator.GreaterThan.Apply(new CompileResult(value, DataType.String), CompileResult.Empty).Result);
			Assert.AreEqual(false, Operator.GreaterThan.Apply(CompileResult.Empty, new CompileResult(value, DataType.String)).Result);
			Assert.AreEqual(false, Operator.GreaterThan.Apply(new CompileResult(string.Empty, DataType.String), CompileResult.Empty).Result);
			Assert.AreEqual(false, Operator.GreaterThan.Apply(CompileResult.Empty, new CompileResult(string.Empty, DataType.String)).Result);

			Assert.AreEqual(true, Operator.GreaterThanOrEqual.Apply(new CompileResult(value, DataType.String), CompileResult.Empty).Result);
			Assert.AreEqual(false, Operator.GreaterThanOrEqual.Apply(CompileResult.Empty, new CompileResult(value, DataType.String)).Result);
			Assert.AreEqual(true, Operator.GreaterThanOrEqual.Apply(new CompileResult(string.Empty, DataType.String), CompileResult.Empty).Result);
			Assert.AreEqual(true, Operator.GreaterThanOrEqual.Apply(CompileResult.Empty, new CompileResult(string.Empty, DataType.String)).Result);

			Assert.AreEqual(false, Operator.EqualsTo.Apply(new CompileResult(value, DataType.String), CompileResult.Empty).Result);
			Assert.AreEqual(false, Operator.EqualsTo.Apply(CompileResult.Empty, new CompileResult(value, DataType.String)).Result);
			Assert.AreEqual(true, Operator.EqualsTo.Apply(new CompileResult(string.Empty, DataType.String), CompileResult.Empty).Result);
			Assert.AreEqual(true, Operator.EqualsTo.Apply(CompileResult.Empty, new CompileResult(string.Empty, DataType.String)).Result);

			Assert.AreEqual(true, Operator.NotEqualsTo.Apply(new CompileResult(value, DataType.String), CompileResult.Empty).Result);
			Assert.AreEqual(true, Operator.NotEqualsTo.Apply(CompileResult.Empty, new CompileResult(value, DataType.String)).Result);
			Assert.AreEqual(false, Operator.NotEqualsTo.Apply(new CompileResult(string.Empty, DataType.String), CompileResult.Empty).Result);
			Assert.AreEqual(false, Operator.NotEqualsTo.Apply(CompileResult.Empty, new CompileResult(string.Empty, DataType.String)).Result);

			Assert.AreEqual(false, Operator.LessThan.Apply(new CompileResult(value, DataType.String), CompileResult.Empty).Result);
			Assert.AreEqual(true, Operator.LessThan.Apply(CompileResult.Empty, new CompileResult(value, DataType.String)).Result);
			Assert.AreEqual(false, Operator.LessThan.Apply(new CompileResult(string.Empty, DataType.String), CompileResult.Empty).Result);
			Assert.AreEqual(false, Operator.LessThan.Apply(CompileResult.Empty, new CompileResult(string.Empty, DataType.String)).Result);

			Assert.AreEqual(false, Operator.LessThanOrEqual.Apply(new CompileResult(value, DataType.String), CompileResult.Empty).Result);
			Assert.AreEqual(true, Operator.LessThanOrEqual.Apply(CompileResult.Empty, new CompileResult(value, DataType.String)).Result);
			Assert.AreEqual(true, Operator.LessThanOrEqual.Apply(new CompileResult(string.Empty, DataType.String), CompileResult.Empty).Result);
			Assert.AreEqual(true, Operator.LessThanOrEqual.Apply(CompileResult.Empty, new CompileResult(string.Empty, DataType.String)).Result);
		}
		#endregion

		#region Operator Plus Tests
		[TestMethod]
		public void OperatorPlusShouldThrowExceptionIfNonNumericOperand()
		{
			var result = Operator.Plus.Apply(new CompileResult(1, DataType.Integer), new CompileResult("a", DataType.String));
			Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), result.Result);
		}

		[TestMethod]
		public void OperatorPlusShouldAddNumericStringAndNumber()
		{
			var result = Operator.Plus.Apply(new CompileResult(1, DataType.Integer), new CompileResult("2", DataType.String));
			Assert.AreEqual(3d, result.Result);
		}

		[TestMethod]
		public void OperatorPlusTimeDataTypes()
		{
			var result = Operator.Plus.Apply(new CompileResult(.6, DataType.Time), new CompileResult(.8, DataType.Time));
			Assert.AreEqual(.6d + .8d, result.Result);
			Assert.AreEqual(DataType.Time, result.DataType);
		}

		[TestMethod]
		public void OperatorPlusDateAndTime()
		{
			var result = Operator.Plus.Apply(new CompileResult(43306, DataType.Date), new CompileResult(.8, DataType.Time));
			Assert.AreEqual(43306 + .8d, result.Result);
			Assert.AreEqual(DataType.Date, result.DataType);
		}

		[TestMethod]
		public void OperatorPlusTimeAndDate()
		{
			var result = Operator.Plus.Apply(new CompileResult(.8, DataType.Time), new CompileResult(43306, DataType.Date));
			Assert.AreEqual(43306 + .8d, result.Result);
			Assert.AreEqual(DataType.Time, result.DataType);
		}

		[TestMethod]
		public void OperatorPlusDateDataTypes()
		{
			var result = Operator.Plus.Apply(new CompileResult(43309, DataType.Date), new CompileResult(43306, DataType.Date));
			Assert.AreEqual(43309d + 43306d, result.Result);
			Assert.AreEqual(DataType.Date, result.DataType);
		}

		[TestMethod]
		public void OperatorPlusErrorTypeArguments()
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet1");
				sheet.Cells[2, 2].Formula = @"2+2";
				sheet.Cells[2, 3].Formula = @"2+""text""";
				sheet.Cells[2, 4].Formula = @"2+notavalidname";
				sheet.Cells[2, 5].Formula = @"""text""+2";
				sheet.Cells[2, 6].Formula = @"""text""+""other text""";
				sheet.Cells[2, 7].Formula = @"""text""+notavalidname";
				sheet.Cells[2, 8].Formula = @"notavalidname+2";
				sheet.Cells[2, 9].Formula = @"notavalidname+""other text""";
				sheet.Cells[2, 10].Formula = @"notavalidname+notavalidname";
				sheet.Calculate();
				Assert.AreEqual(4d, sheet.Cells[2, 2].Value);
				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells[2, 3].Value);
				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Name), sheet.Cells[2, 4].Value);
				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells[2, 5].Value);
				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells[2, 6].Value);
				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells[2, 7].Value);
				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Name), sheet.Cells[2, 8].Value);
				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Name), sheet.Cells[2, 9].Value);
				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Name), sheet.Cells[2, 10].Value);
			}
		}
		#endregion

		#region Operator Minus Tests
		[TestMethod]
		public void OperatorMinusShouldThrowExceptionIfNonNumericOperand()
		{
			var result = Operator.Minus.Apply(new CompileResult(1, DataType.Integer), new CompileResult("a", DataType.String));
			Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), result.Result);
		}

		[TestMethod]
		public void OperatorMinusShouldSubtractNumericStringAndNumber()
		{
			var result = Operator.Minus.Apply(new CompileResult(5, DataType.Integer), new CompileResult("2", DataType.String));
			Assert.AreEqual(3d, result.Result);
		}

		[TestMethod]
		public void OperatorMinusTimeDataTypes()
		{
			var result = Operator.Minus.Apply(new CompileResult(.8, DataType.Time), new CompileResult(.6, DataType.Time));
			Assert.AreEqual(.8d - .6d, result.Result);
			Assert.AreEqual(DataType.Time, result.DataType);
		}

		[TestMethod]
		public void OperatorMinusDateAndTime()
		{
			var result = Operator.Minus.Apply(new CompileResult(43306, DataType.Date), new CompileResult(.8, DataType.Time));
			Assert.AreEqual(43306 - .8d, result.Result);
			Assert.AreEqual(DataType.Date, result.DataType);
		}

		[TestMethod]
		public void OperatorMinusTimeAndDate()
		{
			var result = Operator.Minus.Apply(new CompileResult(.8, DataType.Time), new CompileResult(43306, DataType.Date));
			Assert.AreEqual(.8d - 43306, result.Result);
			Assert.AreEqual(DataType.Time, result.DataType);
		}

		[TestMethod]
		public void OperatorMinusDateDataTypes()
		{
			var result = Operator.Minus.Apply(new CompileResult(43399, DataType.Date), new CompileResult(43306, DataType.Date));
			Assert.AreEqual(43399d - 43306d, result.Result);
			Assert.AreEqual(DataType.Date, result.DataType);
		}

		[TestMethod]
		public void OperatorMinusErrorTypeArguments()
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet1");
				sheet.Cells[2, 2].Formula = @"5-2";
				sheet.Cells[2, 3].Formula = @"2-""text""";
				sheet.Cells[2, 4].Formula = @"2-notavalidname";
				sheet.Cells[2, 5].Formula = @"""text""-2";
				sheet.Cells[2, 6].Formula = @"""text""-""other text""";
				sheet.Cells[2, 7].Formula = @"""text""-notavalidname";
				sheet.Cells[2, 8].Formula = @"notavalidname-2";
				sheet.Cells[2, 9].Formula = @"notavalidname-""other text""";
				sheet.Cells[2, 10].Formula = @"notavalidname-notavalidname";
				sheet.Calculate();
				Assert.AreEqual(3d, sheet.Cells[2, 2].Value);
				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells[2, 3].Value);
				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Name), sheet.Cells[2, 4].Value);
				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells[2, 5].Value);
				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells[2, 6].Value);
				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells[2, 7].Value);
				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Name), sheet.Cells[2, 8].Value);
				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Name), sheet.Cells[2, 9].Value);
				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Name), sheet.Cells[2, 10].Value);
			}
		}
		#endregion

		#region Operator Divide Tests
		[TestMethod]
		public void OperatorDivideShouldReturnDivideByZeroIfRightOperandIsZero()
		{
			var result = Operator.Divide.Apply(new CompileResult(1d, DataType.Decimal), new CompileResult(0d, DataType.Decimal));
			Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Div0), result.Result);
		}

		[TestMethod]
		public void OperatorDivideShouldDivideCorrectly()
		{
			var result = Operator.Divide.Apply(new CompileResult(9d, DataType.Decimal), new CompileResult(3d, DataType.Decimal));
			Assert.AreEqual(3d, result.Result);
		}

		[TestMethod]
		public void OperatorDivideShouldReturnValueErrorIfNonNumericOperand()
		{
			var result = Operator.Divide.Apply(new CompileResult(1, DataType.Integer), new CompileResult("a", DataType.String));
			Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), result.Result);
		}

		[TestMethod]
		public void OperatorDivideShouldDivideNumericStringAndNumber()
		{
			var result = Operator.Divide.Apply(new CompileResult(9, DataType.Integer), new CompileResult("3", DataType.String));
			Assert.AreEqual(3d, result.Result);
		}

		[TestMethod]
		public void OperatorDivideWithTimes()
		{
			var result = Operator.Divide.Apply(new CompileResult(.8, DataType.Time), new CompileResult(.5, DataType.Time));
			Assert.AreEqual(.8 / .5, (double)result.Result, .0000001);
			Assert.AreEqual(DataType.Time, result.DataType);
		}

		[TestMethod]
		public void OperatorDivideDateWithTime()
		{
			var result = Operator.Divide.Apply(new CompileResult(43306, DataType.Date), new CompileResult(.4, DataType.Time));
			Assert.AreEqual(43306d / .4, (double)result.Result, .0000001);
			Assert.AreEqual(DataType.Date, result.DataType);
		}

		[TestMethod]
		public void OperatorDivideTimeWithDate()
		{
			var result = Operator.Divide.Apply(new CompileResult(.4, DataType.Time), new CompileResult(5, DataType.Date));
			Assert.AreEqual(.4 / 5d, (double)result.Result, .0000001);
			Assert.AreEqual(DataType.Time, result.DataType);
		}

		[TestMethod]
		public void OperatorDivideWithDates()
		{
			var result = Operator.Divide.Apply(new CompileResult(40336, DataType.Date), new CompileResult(40339, DataType.Date));
			Assert.AreEqual(40336d / 40339d, (double)result.Result, .0000001);
			Assert.AreEqual(DataType.Decimal, result.DataType);
		}

		[TestMethod]
		public void OperatorDivideErrorTypeArguments()
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet1");
				sheet.Cells[2, 2].Formula = @"10/2";
				sheet.Cells[2, 3].Formula = @"2/""text""";
				sheet.Cells[2, 4].Formula = @"2/notavalidname";
				sheet.Cells[2, 5].Formula = @"""text""/2";
				sheet.Cells[2, 6].Formula = @"""text""/""other text""";
				sheet.Cells[2, 7].Formula = @"""text""/notavalidname";
				sheet.Cells[2, 8].Formula = @"notavalidname/2";
				sheet.Cells[2, 9].Formula = @"notavalidname/""other text""";
				sheet.Cells[2, 10].Formula = @"notavalidname/notavalidname";
				sheet.Calculate();
				Assert.AreEqual(5d, sheet.Cells[2, 2].Value);
				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells[2, 3].Value);
				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Name), sheet.Cells[2, 4].Value);
				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells[2, 5].Value);
				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells[2, 6].Value);
				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells[2, 7].Value);
				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Name), sheet.Cells[2, 8].Value);
				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Name), sheet.Cells[2, 9].Value);
				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Name), sheet.Cells[2, 10].Value);
			}
		}
		#endregion

		#region Operator Multiply Tests
		[TestMethod]
		public void OperatorMultiplyShouldThrowExceptionIfNonNumericOperand()
		{
			Operator.Multiply.Apply(new CompileResult(1, DataType.Integer), new CompileResult("a", DataType.String));
		}

		[TestMethod]
		public void OperatorMultiplyShouldMultiplyNumericStringAndNumber()
		{
			var result = Operator.Multiply.Apply(new CompileResult(1, DataType.Integer), new CompileResult("3", DataType.String));
			Assert.AreEqual(3d, result.Result);
		}

		[TestMethod]
		public void OperatorMultiplyWithNonZeroIntegersReturnsCorrectResult()
		{
			var result = Operator.Multiply.Apply(new CompileResult(5, DataType.Integer), new CompileResult(8, DataType.Integer));
			Assert.AreEqual(40d, result.Result);
		}

		[TestMethod]
		public void OperatorMultiplyWithZeroIntegerReturnsZero()
		{
			var result = Operator.Multiply.Apply(new CompileResult(5, DataType.Integer), new CompileResult(0, DataType.Integer));
			Assert.AreEqual(0d, result.Result);
		}

		[TestMethod]
		public void OperatorMultiplyWithTwoNegativeIntegersReturnsCorrectValue()
		{
			var result = Operator.Multiply.Apply(new CompileResult(-5, DataType.Integer), new CompileResult(-8, DataType.Integer));
			Assert.AreEqual(40d, result.Result);
		}

		[TestMethod]
		public void OperatorMultiplyWithOneNegativeIntegerReturnsCorrectValue()
		{
			var result = Operator.Multiply.Apply(new CompileResult(-5, DataType.Integer), new CompileResult(8, DataType.Integer));
			Assert.AreEqual(-40d, result.Result);
		}

		[TestMethod]
		public void OperatorMultiplyWithDoublesReturnsCorrectValue()
		{
			var result = Operator.Multiply.Apply(new CompileResult(3.3, DataType.Decimal), new CompileResult(-5.6, DataType.Decimal));
			Assert.AreEqual(-18.48d, (double)result.Result, 0.000001);
		}

		[TestMethod]
		public void OperatorMultiplyWithTimes()
		{
			var result = Operator.Multiply.Apply(new CompileResult(.89, DataType.Time), new CompileResult(.7, DataType.Time));
			Assert.AreEqual(.89d * .7d, (double)result.Result, 0.000001);
			Assert.AreEqual(DataType.Time, result.DataType);
		}

		[TestMethod]
		public void OperatorMultiplyWithDateAndTime()
		{
			var result = Operator.Multiply.Apply(new CompileResult(43306, DataType.Date), new CompileResult(.7, DataType.Time));
			Assert.AreEqual(43306d * .7d, (double)result.Result, 0.000001);
			Assert.AreEqual(DataType.Date, result.DataType);
		}

		[TestMethod]
		public void OperatorMultiplyWithTimeAndDate()
		{
			var result = Operator.Multiply.Apply(new CompileResult(.7, DataType.Time), new CompileResult(43306, DataType.Date));
			Assert.AreEqual(43306d * .7d, (double)result.Result, 0.000001);
			Assert.AreEqual(DataType.Time, result.DataType);
		}

		[TestMethod]
		public void OperatorMultiplyWithDates()
		{
			var result = Operator.Multiply.Apply(new CompileResult(4, DataType.Date), new CompileResult(6, DataType.Date));
			Assert.AreEqual(4d * 6d, result.Result);
			Assert.AreEqual(DataType.Decimal, result.DataType);
		}

		[TestMethod]
		public void OperatorMultiplyWithFractionsReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["B1"].Formula = "(2/3) * (5/4)";
				ws.Calculate();
				Assert.AreEqual(0.83333333, (double)ws.Cells["B1"].Value, 0.000001);
			}
		}

		[TestMethod]
		public void OperatorMultiplyWithDateFunctionResultReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["B2"].Formula = "2 * DATE(2017,5,1)";
				ws.Calculate();
				Assert.AreEqual(85712d, ws.Cells["B2"].Value);
			}
		}

		[TestMethod]
		public void OperatorMultiplyWithDateAsStringReturnsCorrectValue()
		{
			var result = Operator.Multiply.Apply(new CompileResult(2, DataType.Integer), new CompileResult("5/1/2017", DataType.String));
			Assert.AreEqual(85712d, result.Result);
		}

		[TestMethod]
		public void OperatorMultiplyWithNoArgumentsReturnsPoundValue()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["B1"].Value = 1;
				ws.Cells["B2"].Value = 1;
				ws.Cells["B3"].Value = 1;
				ws.Cells["B4"].Value = 1;
				ws.Cells["B5"].Value = 1;
				ws.Cells["B6"].Formula = "*";
				ws.Calculate();
				Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)ws.Cells["B6"].Value).Type);
			}
		}

		[TestMethod]
		public void OperatorMultiplyWithMaxInputsReturnsCorrectValue()
		{
			// The maximum number of inputs the function takes is 264.
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				for (int i = 1; i < 270; i++)
				{
					for (int j = 2; j < 3; j++)
					{
						ws.Cells[i, j].Value = 2;
					}
				}
				ws.Cells["C1"].Formula = "B1* B2* B3* B4* B5* B6* B7* B8* B9* B10* B11* B12* B13* B14* B15* B16* B17* B18* B19* B20* B21* B22* B23" +
					"* B24* B25* B26* B27* B28* B29* B30* B31* B32* B33* B34* B35* B36* B37* B38* B39* B40* B41* B42* B43* B44* B45* B46* B47* B48*" +
					" B49* B50* B51* B52* B53* B54* B55* B56* B57* B58* B59* B60* B61* B62* B63* B64* B65* B66* B67* B68* B69* B70* B71* B72* B73*" +
					" B74* B75* B76* B77* B78* B79* B80* B81* B82* B83* B84* B85* B86* B87* B88* B89* B90* B91* B92* B93* B94* B95* B96* B97* B98*" +
					" B99* B100* B101* B102* B103* B104* B105* B106* B107* B108* B109* B110* B111* B112* B113* B114* B115* B116* B117* B118* B119* " +
					"B120* B121* B122* B123* B124* B125* B126* B127* B128* B129* B130* B131* B132* B133* B134* B135* B136* B137* B138* B139* B140* " +
					"B141* B142* B143* B144* B145* B146* B147* B148* B149* B150* B151* B152* B153* B154* B155* B156* B157* B158* B159* B160* B170* " +
					"B171* B172* B173* B174* B175* B176* B187* B179* B180* B181* B182* B183* B184* B185* B186* B187* B188* B189* B190* B191* B192* " +
					"B193* B194* B194* B195* B196* B197* B198* B199* B200* B201* B202* B203* B204* B205* B206* B207* B208* B209* B210* B211* B212* " +
					"B213* B214* B215* B216* B217* B218* B219* B220* B221* B222* B223* B224* B225* B226* B227* B228* B229* B230* B231* B232* B233*" +
					"B234* B235* B236* B237* B238* B239* B240* B241* B242* B243* B244* B245* B246* B247* B248* B249* B250* B251* B252* B253* B254* " +
					"B255* B256* B257* B258* B259* B260* B261* B262* B263* B264";
				ws.Calculate();
				Assert.AreEqual(System.Math.Pow(2, 255), ws.Cells["C1"].Value);
			}
		}

		[TestMethod]
		public void OperatorMultiplyErrorTypeArguments()
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet1");
				sheet.Cells[2, 2].Formula = @"2*2";
				sheet.Cells[2, 3].Formula = @"2*""text""";
				sheet.Cells[2, 4].Formula = @"2*notavalidname";
				sheet.Cells[2, 5].Formula = @"""text""*2";
				sheet.Cells[2, 6].Formula = @"""text""*""other text""";
				sheet.Cells[2, 7].Formula = @"""text""*notavalidname";
				sheet.Cells[2, 8].Formula = @"notavalidname*2";
				sheet.Cells[2, 9].Formula = @"notavalidname*""other text""";
				sheet.Cells[2, 10].Formula = @"notavalidname*notavalidname";
				sheet.Calculate();
				Assert.AreEqual(4d, sheet.Cells[2, 2].Value);
				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells[2, 3].Value);
				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Name), sheet.Cells[2, 4].Value);
				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells[2, 5].Value);
				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells[2, 6].Value);
				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells[2, 7].Value);
				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Name), sheet.Cells[2, 8].Value);
				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Name), sheet.Cells[2, 9].Value);
				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Name), sheet.Cells[2, 10].Value);
			}
		}
		#endregion

		#region Operator Percent Tests
		[TestMethod]
		public void OperatorPercentPropagatesErrors()
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet1");
				sheet.Cells[2, 2].Formula = @"2%";
				sheet.Cells[2, 3].Formula = @"""text""%";
				sheet.Cells[2, 4].Formula = @"notavalidname%";
				sheet.Calculate();
				Assert.AreEqual(.02, sheet.Cells[2, 2].Value);
				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells[2, 3].Value);
				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Name), sheet.Cells[2, 4].Value);
			}
		}

		[TestMethod]
		public void OperatorPercentCompounds()
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet1");
				sheet.Cells[2, 2].Formula = @"24%%%";
				sheet.Calculate();
				Assert.AreEqual(.000024d, (double)sheet.Cells[2, 2].Value, .0000000000000001);
			}
		}

		[TestMethod]
		public void OperatorPercentOfDate()
		{
			var result = Operator.Percent.Apply(new CompileResult(40663, DataType.Date), new CompileResult(0.01, DataType.Decimal));
			Assert.AreEqual(40663d * .01d, result.Result);
			Assert.AreEqual(DataType.Decimal, result.DataType);
		}

		[TestMethod]
		public void OperatorPercentOfTime()
		{
			var result = Operator.Percent.Apply(new CompileResult(.87, DataType.Date), new CompileResult(0.01, DataType.Decimal));
			Assert.AreEqual(.87d * .01d, result.Result);
			Assert.AreEqual(DataType.Decimal, result.DataType);
		}
		#endregion

		#region Operator Concat Tests
		[TestMethod]
		public void OperatorConcatShouldConcatTwoStrings()
		{
			var result = Operator.Concat.Apply(new CompileResult("a", DataType.String), new CompileResult("b", DataType.String));
			Assert.AreEqual("ab", result.Result);
		}

		[TestMethod]
		public void OperatorConcatShouldConcatANumberAndAString()
		{
			var result = Operator.Concat.Apply(new CompileResult(12, DataType.Integer), new CompileResult("b", DataType.String));
			Assert.AreEqual("12b", result.Result);
		}

		[TestMethod]
		public void OperatorConcatShouldConcatAnEmptyRange()
		{
			var file = new FileInfo("filename.xlsx");
			using (var package = new ExcelPackage(file))
			using (var sheet = package.Workbook.Worksheets.Add("NewSheet"))
			using (var excelDataProvider = new EpplusExcelDataProvider(package))
			{
				var emptyRange = excelDataProvider.GetRange("NewSheet", 2, 2, "B2");
				var result = Operator.Concat.Apply(new CompileResult(emptyRange, DataType.ExcelAddress), new CompileResult("b", DataType.String));
				Assert.AreEqual("b", result.Result);
				result = Operator.Concat.Apply(new CompileResult("b", DataType.String), new CompileResult(emptyRange, DataType.ExcelAddress));
				Assert.AreEqual("b", result.Result);
			}
		}

		[TestMethod]
		public void OperatorConcatShouldPropagateErrors()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet1");
				sheet.Cells[2, 2].Formula = @"""text""&notaname";
				sheet.Cells[2, 3].Formula = @"notaname&""text""";
				sheet.Cells[2, 4].Formula = @"notaname&ROUND(""AA"", 1)";
				sheet.Cells[2, 5].Formula = @"ROUND(""AA"", 1)&notaname";
				sheet.Calculate();
				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Name), sheet.Cells[2, 2].Value);
				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Name), sheet.Cells[2, 3].Value);
				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Name), sheet.Cells[2, 4].Value);
				Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), sheet.Cells[2, 5].Value);
			}
		}
		#endregion

		#region Operator Equals Tests
		[TestMethod]
		public void OperatorEqShouldReturnTruefSuppliedValuesAreEqual()
		{
			var result = Operator.EqualsTo.Apply(new CompileResult(12, DataType.Integer), new CompileResult(12, DataType.Integer));
			Assert.IsTrue((bool)result.Result);
		}

		[TestMethod]
		public void OperatorEqShouldReturnFalsefSuppliedValuesDiffer()
		{
			var result = Operator.EqualsTo.Apply(new CompileResult(11, DataType.Integer), new CompileResult(12, DataType.Integer));
			Assert.IsFalse((bool)result.Result);
		}
		#endregion

		#region Operator NotEqualsTo Tests
		[TestMethod]
		public void OperatorNotEqualToShouldReturnTruefSuppliedValuesDiffer()
		{
			var result = Operator.NotEqualsTo.Apply(new CompileResult(11, DataType.Integer), new CompileResult(12, DataType.Integer));
			Assert.IsTrue((bool)result.Result);
		}

		[TestMethod]
		public void OperatorNotEqualToShouldReturnFalsefSuppliedValuesAreEqual()
		{
			var result = Operator.NotEqualsTo.Apply(new CompileResult(11, DataType.Integer), new CompileResult(11, DataType.Integer));
			Assert.IsFalse((bool)result.Result);
		}
		#endregion

		#region Operator GreaterThan Tests
		[TestMethod]
		public void OperatorGreaterThanToShouldReturnTrueIfLeftIsSetAndRightIsNull()
		{
			var result = Operator.GreaterThan.Apply(new CompileResult(11, DataType.Integer), new CompileResult(null, DataType.Empty));
			Assert.IsTrue((bool)result.Result);
		}

		[TestMethod]
		public void OperatorGreaterThanToShouldReturnTrueIfLeftIs11AndRightIs10()
		{
			var result = Operator.GreaterThan.Apply(new CompileResult(11, DataType.Integer), new CompileResult(10, DataType.Integer));
			Assert.IsTrue((bool)result.Result);
		}
		#endregion

		#region Operator Exp Tests
		[TestMethod]
		public void OperatorExpShouldReturnCorrectResult()
		{
			var result = Operator.Exp.Apply(new CompileResult(2, DataType.Integer), new CompileResult(3, DataType.Integer));
			Assert.AreEqual(8d, result.Result);
		}

		[TestMethod]
		public void OperatorExpWithDates()
		{
			var result = Operator.Exp.Apply(new CompileResult(40663, DataType.Date), new CompileResult(2, DataType.Date));
			Assert.AreEqual(Math.Pow(40663d, 2d), result.Result);
			Assert.AreEqual(DataType.Decimal, result.DataType);
		}

		[TestMethod]
		public void OperatorExpWithTimes()
		{
			var result = Operator.Exp.Apply(new CompileResult(.6, DataType.Time), new CompileResult(.7, DataType.Time));
			Assert.AreEqual(Math.Pow(.6, .7), (double)result.Result, .0000001);
			Assert.AreEqual(DataType.Decimal, result.DataType);
		}

		[TestMethod]
		public void OperatorExpWithDateAndTime()
		{
			var result = Operator.Exp.Apply(new CompileResult(40663, DataType.Date), new CompileResult(.7, DataType.Time));
			Assert.AreEqual(Math.Pow(40663d, .7d), (double)result.Result, .0000001);
			Assert.AreEqual(DataType.Decimal, result.DataType);
		}
		#endregion

		#region Numeric and Date String Comparison Tests
		[TestMethod]
		public void OperatorsActingOnNumericStrings()
		{
			double number1 = 42.0;
			double number2 = -143.75;
			CompileResult result1 = new CompileResult(number1.ToString("n"), DataType.String);
			CompileResult result2 = new CompileResult(number2.ToString("n"), DataType.String);
			var operatorResult = Operator.Concat.Apply(result1, result2);
			Assert.AreEqual($"{number1.ToString("n")}{number2.ToString("n")}", operatorResult.Result);
			operatorResult = Operator.Divide.Apply(result1, result2);
			Assert.AreEqual(number1 / number2, operatorResult.Result);
			operatorResult = Operator.Exp.Apply(result1, result2);
			Assert.AreEqual(Math.Pow(number1, number2), operatorResult.Result);
			operatorResult = Operator.Minus.Apply(result1, result2);
			Assert.AreEqual(number1 - number2, operatorResult.Result);
			operatorResult = Operator.Multiply.Apply(result1, result2);
			Assert.AreEqual(number1 * number2, operatorResult.Result);
			operatorResult = Operator.Percent.Apply(result1, result2);
			Assert.AreEqual(number1 * number2, operatorResult.Result);
			operatorResult = Operator.Plus.Apply(result1, result2);
			Assert.AreEqual(number1 + number2, operatorResult.Result);
			// Comparison operators always compare string-wise and don't parse out the actual numbers.
			operatorResult = Operator.NotEqualsTo.Apply(result1, new CompileResult(number1.ToString("n0"), DataType.String));
			Assert.IsTrue((bool)operatorResult.Result);
			operatorResult = Operator.EqualsTo.Apply(result1, new CompileResult(number1.ToString("n0"), DataType.String));
			Assert.IsFalse((bool)operatorResult.Result);
			operatorResult = Operator.GreaterThan.Apply(result1, result2);
			Assert.IsTrue((bool)operatorResult.Result);
			operatorResult = Operator.GreaterThanOrEqual.Apply(result1, result2);
			Assert.IsTrue((bool)operatorResult.Result);
			operatorResult = Operator.LessThan.Apply(result1, result2);
			Assert.IsFalse((bool)operatorResult.Result);
			operatorResult = Operator.LessThanOrEqual.Apply(result1, result2);
			Assert.IsFalse((bool)operatorResult.Result);
		}

		[TestMethod]
		public void OperatorsActingOnDateStrings()
		{
			const string dateFormat = "M-dd-yyyy";
			DateTime date1 = new DateTime(2015, 2, 20);
			DateTime date2 = new DateTime(2015, 12, 1);
			var numericDate1 = date1.ToOADate();
			var numericDate2 = date2.ToOADate();
			CompileResult result1 = new CompileResult(date1.ToString(dateFormat), DataType.String); // 2/20/2015
			CompileResult result2 = new CompileResult(date2.ToString(dateFormat), DataType.String); // 12/1/2015
			var operatorResult = Operator.Concat.Apply(result1, result2);
			Assert.AreEqual($"{date1.ToString(dateFormat)}{date2.ToString(dateFormat)}", operatorResult.Result);
			operatorResult = Operator.Divide.Apply(result1, result2);
			Assert.AreEqual(numericDate1 / numericDate2, operatorResult.Result);
			operatorResult = Operator.Exp.Apply(result1, result2);
			Assert.AreEqual(Math.Pow(numericDate1, numericDate2), operatorResult.Result);
			operatorResult = Operator.Minus.Apply(result1, result2);
			Assert.AreEqual(numericDate1 - numericDate2, operatorResult.Result);
			operatorResult = Operator.Multiply.Apply(result1, result2);
			Assert.AreEqual(numericDate1 * numericDate2, operatorResult.Result);
			operatorResult = Operator.Percent.Apply(result1, result2);
			Assert.AreEqual(numericDate1 * numericDate2, operatorResult.Result);
			operatorResult = Operator.Plus.Apply(result1, result2);
			Assert.AreEqual(numericDate1 + numericDate2, operatorResult.Result);
			// Comparison operators always compare string-wise and don't parse out the actual numbers.
			operatorResult = Operator.EqualsTo.Apply(result1, new CompileResult(date1.ToString("f"), DataType.String));
			Assert.IsFalse((bool)operatorResult.Result);
			operatorResult = Operator.NotEqualsTo.Apply(result1, new CompileResult(date1.ToString("f"), DataType.String));
			Assert.IsTrue((bool)operatorResult.Result);
			operatorResult = Operator.GreaterThan.Apply(result1, result2);
			Assert.IsTrue((bool)operatorResult.Result);
			operatorResult = Operator.GreaterThanOrEqual.Apply(result1, result2);
			Assert.IsTrue((bool)operatorResult.Result);
			operatorResult = Operator.LessThan.Apply(result1, result2);
			Assert.IsFalse((bool)operatorResult.Result);
			operatorResult = Operator.LessThanOrEqual.Apply(result1, result2);
			Assert.IsFalse((bool)operatorResult.Result);
		}

		[TestMethod]
		public void OperatorsActingOnGermanDateStrings()
		{
			var currentCulture = CultureInfo.CurrentCulture;
			var culture = new CultureInfo("de-DE");
			Thread.CurrentThread.CurrentCulture = culture;
			try
			{
				string dateFormat = culture.DateTimeFormat.ShortDatePattern;
				DateTime date1 = new DateTime(2015, 2, 20);
				DateTime date2 = new DateTime(2015, 12, 1);
				var numericDate1 = date1.ToOADate();
				var numericDate2 = date2.ToOADate();
				CompileResult result1 = new CompileResult(date1.ToString(dateFormat), DataType.String); // 20.02.2015
				CompileResult result2 = new CompileResult(date2.ToString(dateFormat), DataType.String); // 01.12.2015
				var operatorResult = Operator.Concat.Apply(result1, result2);
				Assert.AreEqual($"{date1.ToString(dateFormat)}{date2.ToString(dateFormat)}", operatorResult.Result);
				operatorResult = Operator.Divide.Apply(result1, result2);
				Assert.AreEqual(numericDate1 / numericDate2, operatorResult.Result);
				operatorResult = Operator.Exp.Apply(result1, result2);
				Assert.AreEqual(Math.Pow(numericDate1, numericDate2), operatorResult.Result);
				operatorResult = Operator.Minus.Apply(result1, result2);
				Assert.AreEqual(numericDate1 - numericDate2, operatorResult.Result);
				operatorResult = Operator.Multiply.Apply(result1, result2);
				Assert.AreEqual(numericDate1 * numericDate2, operatorResult.Result);
				operatorResult = Operator.Percent.Apply(result1, result2);
				Assert.AreEqual(numericDate1 * numericDate2, operatorResult.Result);
				operatorResult = Operator.Plus.Apply(result1, result2);
				Assert.AreEqual(numericDate1 + numericDate2, operatorResult.Result);
				// Comparison operators always compare strings string-wise and don't parse out the actual numbers.
				operatorResult = Operator.EqualsTo.Apply(result1, new CompileResult(date1.ToString("f"), DataType.String));
				Assert.IsFalse((bool)operatorResult.Result);
				operatorResult = Operator.NotEqualsTo.Apply(result1, new CompileResult(date1.ToString("f"), DataType.String));
				Assert.IsTrue((bool)operatorResult.Result);
				operatorResult = Operator.GreaterThan.Apply(result1, result2);
				Assert.IsTrue((bool)operatorResult.Result);
				operatorResult = Operator.GreaterThanOrEqual.Apply(result1, result2);
				Assert.IsTrue((bool)operatorResult.Result);
				operatorResult = Operator.LessThan.Apply(result1, result2);
				Assert.IsFalse((bool)operatorResult.Result);
				operatorResult = Operator.LessThanOrEqual.Apply(result1, result2);
				Assert.IsFalse((bool)operatorResult.Result);
			}
			finally
			{
				Thread.CurrentThread.CurrentCulture = currentCulture;
			}
		}
		#endregion
	}
}
