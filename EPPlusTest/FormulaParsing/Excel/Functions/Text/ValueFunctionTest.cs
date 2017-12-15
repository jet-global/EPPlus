/*******************************************************************************
* You may amend and distribute as you like, but don't remove this header!
*
* EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
* See http://www.codeplex.com/EPPlus for details.
*
* Copyright (C) 2011-2017 Jan Källman, Evan Schallerer, and others as noted in the source history.
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
using System.Threading;
using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Text
{
	[TestClass]
	public class ValueFunctionTest
	{
		#region Properties
		private ParsingContext ParsingContext { get; } = ParsingContext.Create();
		#endregion

		#region Test Methods
		[TestMethod]
		public void ValueWithIntegerArgumentTest()
		{
			var function = new Value();
			var args = FunctionsHelper.CreateArgs("543");
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(543d, result.Result);
		}

		[TestMethod]
		public void ValueWithDoubleArgumentTest()
		{
			var function = new Value();
			var args = FunctionsHelper.CreateArgs("543.432");
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(543.432, result.Result);
		}

		[TestMethod]
		public void ValueWithGroupSeparatorArgumentTest()
		{
			var function = new Value();
			var args = FunctionsHelper.CreateArgs("543,432.124");
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(543432.124, result.Result);
		}

		[TestMethod]
		public void ValueWithScientificFormatArgumentTest()
		{
			var function = new Value();
			var args = FunctionsHelper.CreateArgs("1.2345E-03");
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(0.0012345d, result.Result);
		}

		[TestMethod]
		public void ValueWithDateFormatUSArgumentTest()
		{
			var currentCulture = Thread.CurrentThread.CurrentCulture;
			Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
			try
			{
				var function = new Value();
				var date = new DateTime(2015, 12, 31);
				var args = FunctionsHelper.CreateArgs(date.ToString(CultureInfo.CurrentCulture));
				var result = function.Execute(args, this.ParsingContext);
				Assert.AreEqual(date.ToOADate(), result.Result);
			}
			finally
			{
				Thread.CurrentThread.CurrentCulture = currentCulture;
			}
		}

		[TestMethod]
		public void ValueWithTimeFormatUSArgumentTest()
		{
			var currentCulture = Thread.CurrentThread.CurrentCulture;
			Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
			try
			{
				var function = new Value();
				var time = new DateTime(2015, 12, 31, 12, 00, 00).Subtract(new DateTime(2015, 12, 31));
				var args = FunctionsHelper.CreateArgs(time.ToString());
				var result = function.Execute(args, this.ParsingContext);
				Assert.AreEqual(0.5, result.Result);
			}
			finally
			{
				Thread.CurrentThread.CurrentCulture = currentCulture;
			}
		}

		[TestMethod]
		public void ValueWithDateFormatGermanArgumentTest()
		{
			var currentCulture = Thread.CurrentThread.CurrentCulture;
			Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE");
			try
			{
				var function = new Value();
				var date = new DateTime(2017, 11, 13);
				var args = FunctionsHelper.CreateArgs(date.ToString(CultureInfo.CurrentCulture));
				var result = function.Execute(args, this.ParsingContext);
				Assert.AreEqual(date.ToOADate(), result.Result);
			}
			finally
			{
				Thread.CurrentThread.CurrentCulture = currentCulture;
			}
		}

		[TestMethod]
		public void ValueWithValidPercentSignNumberArgumentDividesTest()
		{
			var function = new Value();
			var args = FunctionsHelper.CreateArgs("1234%");
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(12.34, result.Result);
		}

		[TestMethod]
		public void ValueWithNumberArgumentTest()
		{
			var function = new Value();
			var args = FunctionsHelper.CreateArgs(15);
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(15d, result.Result);
		}

		[TestMethod]
		public void ValueWithTooManyPercentSignsArgumentPoundValuesTest()
		{
			var function = new Value();
			var args = FunctionsHelper.CreateArgs("1234%%");
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void ValueWithAlphaStringArgumentPoundValuesTest()
		{
			var function = new Value();
			var args = FunctionsHelper.CreateArgs("asd");
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void ValueWithEmptyStringArgumentPoundValuesTest()
		{
			var function = new Value();
			var args = FunctionsHelper.CreateArgs(string.Empty);
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void ValueWithNoArgumentPoundValuesTest()
		{
			var function = new Value();
			var args = FunctionsHelper.CreateArgs();
			var result = function.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}
		#endregion
	}
}
