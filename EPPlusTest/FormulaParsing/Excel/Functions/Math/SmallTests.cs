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
	public class SmallTests: MathFunctionsTestBase
	{
		#region Small Function (Execute) Tests
		[TestMethod]
		public void SmallWithCorrectInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["A2"].Value = 45;
				ws.Cells["A3"].Value = 2;
				ws.Cells["A4"].Value = 789;
				ws.Cells["A5"].Value = 3;
				ws.Cells["B1"].Formula = "SMALL(A2:A5, 2)";
				ws.Calculate();
				Assert.AreEqual(3d, ws.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void SmallWithNullFirstParameterReturnsPoundNum()
		{
			var function = new Small();
			var result = function.Execute(FunctionsHelper.CreateArgs(null, 5), this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SmallWithNullSecondParameterReturnsPoundNum()
		{
			var function = new Small();
			var result = function.Execute(FunctionsHelper.CreateArgs(5, null), this.ParsingContext);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void SmallWithNegativeSecondInputReturnsPoundNum()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["A2"].Value = 45;
				ws.Cells["A3"].Value = 2;
				ws.Cells["A4"].Value = 789;
				ws.Cells["B1"].Formula = "SMALL(A2:A4, -78)";
				ws.Calculate();
				Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)ws.Cells["B1"].Value).Type);
			}
		}

		[TestMethod]
		public void SmallWithSecondInputOutOfArrayRangeReturnsPoundNum()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["A2"].Value = 45;
				ws.Cells["A3"].Value = 2;
				ws.Cells["A4"].Value = 789;
				ws.Cells["B1"].Formula = "SMALL(A2:A4, 67)";
				ws.Calculate();
				Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)ws.Cells["B1"].Value).Type);
			}
		}

		[TestMethod]
		public void SmallWithNumericStringSecondInputReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["A2"].Value = 45;
				ws.Cells["A3"].Value = 2;
				ws.Cells["A4"].Value = 789;
				ws.Cells["A5"].Value = 3;
				ws.Cells["B1"].Formula = "SMALL(A2:A5, \"3\")";
				ws.Calculate();
				Assert.AreEqual(45d, ws.Cells["B1"].Value);
			}
		}

		[TestMethod]
		public void SmallWithStringSecondInputReturnsPoundValue()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["A2"].Value = 45;
				ws.Cells["A3"].Value = 2;
				ws.Cells["A4"].Value = 789;
				ws.Cells["A5"].Value = 3;
				ws.Cells["B1"].Formula = "SMALL(A2:A5, \"string\")";
				ws.Calculate();
				Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)ws.Cells["B1"].Value).Type);
			}
		}

		[TestMethod]
		public void SmallWithStringArrayReturnsPoundNum()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["A2"].Value = "string";
				ws.Cells["A3"].Value = "string";
				ws.Cells["B1"].Formula = "SMALL(A2:A3, 2)";
				ws.Calculate();
				Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)ws.Cells["B1"].Value).Type);
			}
		}


	}
		#endregion
}
