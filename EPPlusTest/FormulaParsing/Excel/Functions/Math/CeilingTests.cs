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
using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math
{
	[TestClass]
	public class CeilingTests : MathFunctionsTestBase
	{
		#region Ceiling Function (Execute) Tests
		[TestMethod]
		public void CeilingWithNoInputsReturnsPoundValue()
		{
			var function = new Ceiling();
			var result = function.Execute(FunctionsHelper.CreateArgs(), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void CeilingWithPositiveInputsReturnsCorrectValue()
		{
			var function = new Ceiling();
			var result = function.Execute(FunctionsHelper.CreateArgs(3.7, 2), this.ParsingContext);
			Assert.AreEqual(4d, result.Result);
		}

		[TestMethod]
		public void CeilingWithNegativeInputsReturnsCorrectValue()
		{
			var function = new Ceiling();
			var result = function.Execute(FunctionsHelper.CreateArgs(-2.5, -2), this.ParsingContext);
			Assert.AreEqual(-4d, result.Result);
		}

		[TestMethod]
		public void CeilingWithPositiveAndNegativeInputReturnsPoundNum()
		{
			var function = new Ceiling();
			var result = function.Execute(FunctionsHelper.CreateArgs(52.3, -9), this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void CeilingWithIntegerFirstInputReturnsCorrectValue()
		{
			var function = new Ceiling();
			var result = function.Execute(FunctionsHelper.CreateArgs(45, 2), this.ParsingContext);
			Assert.AreEqual(46d, result.Result);
		}

		[TestMethod]
		public void CeilingWithDoubleFirstInputReturnsCorrectValue()
		{
			var function = new Ceiling();
			var result = function.Execute(FunctionsHelper.CreateArgs(12.63, 3), this.ParsingContext);
			Assert.AreEqual(15d, result.Result);
		}

		[TestMethod]
		public void CeilingWithGeneralStringFirstInputReturnsCorrectValue()
		{
			var function = new Ceiling();
			var result = function.Execute(FunctionsHelper.CreateArgs("string", 3), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void CeilingWithNumbericStringFirstInputReturnsCorrectValue()
		{
			var function = new Ceiling();
			var result = function.Execute(FunctionsHelper.CreateArgs("45.6", 4), this.ParsingContext);
			Assert.AreEqual(48d, result.Result);
		}

		[TestMethod]
		public void CeilingWithDateFunctionFirstInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells["B1"].Formula = "CEILING(DATE(2017, 6, 5), 5)";
				worksheet.Calculate();
				Assert.AreEqual(42895d, worksheet.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void CeilingWithDateAsStringFirstInuptReturnsCorrectValue()
		{
			var function = new Ceiling();
			var result = function.Execute(FunctionsHelper.CreateArgs("5/5/2017", 4), this.ParsingContext);
			Assert.AreEqual(42860d, result.Result);
		}
		#endregion
	}
}
