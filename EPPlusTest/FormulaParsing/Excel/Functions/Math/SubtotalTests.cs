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
using System.IO;
using System.Linq;
using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math
{

	[TestClass]
	public class SubtotalTests : MathFunctionsTestBase
	{
	#region Subtotal Function(Execute) Tests

		#region Subtotal Tests Built By Jet Interns - 2017
		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum1()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["B4"].Value = 5;
				worksheet.Cells["B5"].Value = 10;
				worksheet.Cells["A1"].Formula = "=subtotal(1,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(4.2, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubTotalCountAEmptySingleCell()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["A1"].Formula = "=subtotal(103,B1)";
				worksheet.Calculate();
				Assert.AreEqual(0, worksheet.Cells["A1"].Value);
			}
		}

		[TestMethod]
		public void SubTotalCountAEmptyArrayOfCells()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["A1"].Formula = "=subtotal(103,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(0d, worksheet.Cells["A1"].Value);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum6TestTwo()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = -1;
				worksheet.Cells["B2"].Value = -2;
				worksheet.Cells["B3"].Value = -3;
				worksheet.Cells["A1"].Formula = "=subtotal(6,B1:B3)";
				worksheet.Calculate();
				Assert.AreEqual(-6, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum101()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["B4"].Value = 5;
				worksheet.Cells["B5"].Value = 10;
				worksheet.Cells["A1"].Formula = "=subtotal(101,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(4.2, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum2()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["B4"].Value = 5;
				worksheet.Cells["B5"].Value = 10;
				worksheet.Cells["A1"].Formula = "=subtotal(2,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(5, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum102()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["B4"].Value = 5;
				worksheet.Cells["B5"].Value = 10;
				worksheet.Cells["A1"].Formula = "=subtotal(102,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(5, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum3()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["B4"].Value = 5;
				worksheet.Cells["B5"].Value = 10;
				worksheet.Cells["A1"].Formula = "=subtotal(3,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(5, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum103()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["B4"].Value = 5;
				worksheet.Cells["B5"].Value = 10;
				worksheet.Cells["A1"].Formula = "=subtotal(103,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(5, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum4()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["B4"].Value = 5;
				worksheet.Cells["B5"].Value = 10;
				worksheet.Cells["A1"].Formula = "=subtotal(4,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(10, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum104()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["B4"].Value = 5;
				worksheet.Cells["B5"].Value = 10;
				worksheet.Cells["A1"].Formula = "=subtotal(104,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(10, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum5()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["B4"].Value = 5;
				worksheet.Cells["B5"].Value = 10;
				worksheet.Cells["A1"].Formula = "=subtotal(5,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(1, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum105()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["B4"].Value = 5;
				worksheet.Cells["B5"].Value = 10;
				worksheet.Cells["A1"].Formula = "=subtotal(105,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(1, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum6()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["B4"].Value = 5;
				worksheet.Cells["B5"].Value = 10;
				worksheet.Cells["A1"].Formula = "=subtotal(6,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(300, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum106()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["B4"].Value = 5;
				worksheet.Cells["B5"].Value = 10;
				worksheet.Cells["A1"].Formula = "=subtotal(106,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(300, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum7()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["B4"].Value = 5;
				worksheet.Cells["B5"].Value = 10;
				worksheet.Cells["A1"].Formula = "=subtotal(7,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(3.563705936, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum107()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["B4"].Value = 5;
				worksheet.Cells["B5"].Value = 10;
				worksheet.Cells["A1"].Formula = "=subtotal(107,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(3.563705936, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum8()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["B4"].Value = 5;
				worksheet.Cells["B5"].Value = 10;
				worksheet.Cells["A1"].Formula = "=subtotal(8,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(3.18747549, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum108()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["B4"].Value = 5;
				worksheet.Cells["B5"].Value = 10;
				worksheet.Cells["A1"].Formula = "=subtotal(108,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(3.18747549, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum9()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["B4"].Value = 5;
				worksheet.Cells["B5"].Value = 10;
				worksheet.Cells["A1"].Formula = "=subtotal(9,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(21, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum109()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["B4"].Value = 5;
				worksheet.Cells["B5"].Value = 10;
				worksheet.Cells["A1"].Formula = "=subtotal(109,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(21, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalSumWithDecimal()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = (decimal)7;
				worksheet.Cells["A1"].Formula = "=subtotal(109,B1)";
				worksheet.Calculate();
				Assert.AreEqual(7d, worksheet.Cells["A1"].Value);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum10()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["B4"].Value = 5;
				worksheet.Cells["B5"].Value = 10;
				worksheet.Cells["A1"].Formula = "=subtotal(10,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(12.7, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum110()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["B4"].Value = 5;
				worksheet.Cells["B5"].Value = 10;
				worksheet.Cells["A1"].Formula = "=subtotal(110,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(12.7, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum11()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["B4"].Value = 5;
				worksheet.Cells["B5"].Value = 10;
				worksheet.Cells["A1"].Formula = "=subtotal(11,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(10.16, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum111()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["B4"].Value = 5;
				worksheet.Cells["B5"].Value = 10;
				worksheet.Cells["A1"].Formula = "=subtotal(111,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(10.16, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum1TestingBools()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = true;
				worksheet.Cells["B2"].Value = false;
				worksheet.Cells["B3"].Value = true;
				worksheet.Cells["B4"].Value = 1;
				worksheet.Cells["B5"].Value = 1;
				worksheet.Cells["A1"].Formula = "=subtotal(1,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(1, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum101TestingBools()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = true;
				worksheet.Cells["B2"].Value = false;
				worksheet.Cells["B3"].Value = true;
				worksheet.Cells["B4"].Value = 1;
				worksheet.Cells["B5"].Value = 1;
				worksheet.Cells["A1"].Formula = "=subtotal(101,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(1, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum2TestingBools()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = true;
				worksheet.Cells["B2"].Value = false;
				worksheet.Cells["B3"].Value = true;
				worksheet.Cells["B4"].Value = 1;
				worksheet.Cells["B5"].Value = 1;
				worksheet.Cells["A1"].Formula = "=subtotal(2,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(2, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum102TestingBools()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = true;
				worksheet.Cells["B2"].Value = false;
				worksheet.Cells["B3"].Value = true;
				worksheet.Cells["B4"].Value = 1;
				worksheet.Cells["B5"].Value = 1;
				worksheet.Cells["A1"].Formula = "=subtotal(102,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(2, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum3TestingBools()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = true;
				worksheet.Cells["B2"].Value = false;
				worksheet.Cells["B3"].Value = true;
				worksheet.Cells["B4"].Value = 1;
				worksheet.Cells["B5"].Value = 1;
				worksheet.Cells["A1"].Formula = "=subtotal(3,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(5, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum103TestingBools()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = true;
				worksheet.Cells["B2"].Value = false;
				worksheet.Cells["B3"].Value = true;
				worksheet.Cells["B4"].Value = 1;
				worksheet.Cells["B5"].Value = 1;
				worksheet.Cells["A1"].Formula = "=subtotal(103,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(5, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum4TestingBools()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = true;
				worksheet.Cells["B2"].Value = false;
				worksheet.Cells["B3"].Value = true;
				worksheet.Cells["B4"].Value = 1;
				worksheet.Cells["B5"].Value = 1;
				worksheet.Cells["A1"].Formula = "=subtotal(4,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(1, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum104TestingBools()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = true;
				worksheet.Cells["B2"].Value = false;
				worksheet.Cells["B3"].Value = true;
				worksheet.Cells["B4"].Value = 1;
				worksheet.Cells["B5"].Value = 1;
				worksheet.Cells["A1"].Formula = "=subtotal(104,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(1, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum5TestingBools()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = true;
				worksheet.Cells["B2"].Value = false;
				worksheet.Cells["B3"].Value = true;
				worksheet.Cells["B4"].Value = 1;
				worksheet.Cells["B5"].Value = 1;
				worksheet.Cells["A1"].Formula = "=subtotal(5,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(1, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum105TestingBools()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = true;
				worksheet.Cells["B2"].Value = false;
				worksheet.Cells["B3"].Value = true;
				worksheet.Cells["B4"].Value = 1;
				worksheet.Cells["B5"].Value = 1;
				worksheet.Cells["A1"].Formula = "=subtotal(105,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(1, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum6TestingBools()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = true;
				worksheet.Cells["B2"].Value = false;
				worksheet.Cells["B3"].Value = true;
				worksheet.Cells["B4"].Value = 1;
				worksheet.Cells["B5"].Value = 1;
				worksheet.Cells["A1"].Formula = "=subtotal(6,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(1, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum106TestingBools()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = true;
				worksheet.Cells["B2"].Value = false;
				worksheet.Cells["B3"].Value = true;
				worksheet.Cells["B4"].Value = 1;
				worksheet.Cells["B5"].Value = 1;
				worksheet.Cells["A1"].Formula = "=subtotal(106,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(1, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum7TestingBools()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = true;
				worksheet.Cells["B2"].Value = false;
				worksheet.Cells["B3"].Value = true;
				worksheet.Cells["B4"].Value = 1;
				worksheet.Cells["B5"].Value = 1;
				worksheet.Cells["A1"].Formula = "=subtotal(7,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(0, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum107TestingBools()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = true;
				worksheet.Cells["B2"].Value = false;
				worksheet.Cells["B3"].Value = true;
				worksheet.Cells["B4"].Value = 1;
				worksheet.Cells["B5"].Value = 1;
				worksheet.Cells["A1"].Formula = "=subtotal(107,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(0, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum8TestingBools()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = true;
				worksheet.Cells["B2"].Value = false;
				worksheet.Cells["B3"].Value = true;
				worksheet.Cells["B4"].Value = 1;
				worksheet.Cells["B5"].Value = 1;
				worksheet.Cells["A1"].Formula = "=subtotal(8,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(0, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum108TestingBools()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = true;
				worksheet.Cells["B2"].Value = false;
				worksheet.Cells["B3"].Value = true;
				worksheet.Cells["B4"].Value = 1;
				worksheet.Cells["B5"].Value = 1;
				worksheet.Cells["A1"].Formula = "=subtotal(108,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(0, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum9TestingBools()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = true;
				worksheet.Cells["B2"].Value = false;
				worksheet.Cells["B3"].Value = true;
				worksheet.Cells["B4"].Value = 1;
				worksheet.Cells["B5"].Value = 1;
				worksheet.Cells["B6"].Formula = "=subtotal(9,B1:B5)";
				worksheet.Cells["B7"].Formula = "=subtotal(9,B1:B6)";
				worksheet.Calculate();
				Assert.AreEqual(2, (double)worksheet.Cells["B6"].Value, .00001);
				Assert.AreEqual(2, (double)worksheet.Cells["B7"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum109TestingBools()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = true;
				worksheet.Cells["B2"].Value = false;
				worksheet.Cells["B3"].Value = true;
				worksheet.Cells["B4"].Value = 1;
				worksheet.Cells["B5"].Value = 1;
				worksheet.Cells["A1"].Formula = "=subtotal(109,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(2, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum10TestingBools()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = true;
				worksheet.Cells["B2"].Value = false;
				worksheet.Cells["B3"].Value = true;
				worksheet.Cells["B4"].Value = 1;
				worksheet.Cells["B5"].Value = 1;
				worksheet.Cells["A1"].Formula = "=subtotal(10,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(0, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum110TestingBools()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = true;
				worksheet.Cells["B2"].Value = false;
				worksheet.Cells["B3"].Value = true;
				worksheet.Cells["B4"].Value = 1;
				worksheet.Cells["B5"].Value = 1;
				worksheet.Cells["A1"].Formula = "=subtotal(110,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(0, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum11TestingBools()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = true;
				worksheet.Cells["B2"].Value = false;
				worksheet.Cells["B3"].Value = true;
				worksheet.Cells["B4"].Value = 1;
				worksheet.Cells["B5"].Value = 1;
				worksheet.Cells["A1"].Formula = "=subtotal(11,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(0, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum111TestingBools()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = true;
				worksheet.Cells["B2"].Value = false;
				worksheet.Cells["B3"].Value = true;
				worksheet.Cells["B4"].Value = 1;
				worksheet.Cells["B5"].Value = 1;
				worksheet.Cells["A1"].Formula = "=subtotal(111,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(0, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum1TimeInputs()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = "00:00";
				worksheet.Cells["B2"].Value = "12:00";
				worksheet.Cells["B3"].Value = "13:15";
				worksheet.Cells["B4"].Value = "2:00";
				worksheet.Cells["B5"].Value = "13:00";
				worksheet.Cells["A1"].Formula = "=subtotal(1,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["A1"].Value).Type);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum101TimeInputs()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = "00:00";
				worksheet.Cells["B2"].Value = "12:00";
				worksheet.Cells["B3"].Value = "13:15";
				worksheet.Cells["B4"].Value = "2:00";
				worksheet.Cells["B5"].Value = "13:00";
				worksheet.Cells["A1"].Formula = "=subtotal(101,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["A1"].Value).Type);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum2TimeInputs()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = "00:00";
				worksheet.Cells["B2"].Value = "12:00";
				worksheet.Cells["B3"].Value = "13:15";
				worksheet.Cells["B4"].Value = "2:00";
				worksheet.Cells["B5"].Value = "13:00";
				worksheet.Cells["A1"].Formula = "=subtotal(2,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(0, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum102TimeInputs()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = "00:00";
				worksheet.Cells["B2"].Value = "12:00";
				worksheet.Cells["B3"].Value = "13:15";
				worksheet.Cells["B4"].Value = "2:00";
				worksheet.Cells["B5"].Value = "13:00";
				worksheet.Cells["A1"].Formula = "=subtotal(102,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(0, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum3TimeInputs()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = "00:00";
				worksheet.Cells["B2"].Value = "12:00";
				worksheet.Cells["B3"].Value = "13:15";
				worksheet.Cells["B4"].Value = "2:00";
				worksheet.Cells["B5"].Value = "13:00";
				worksheet.Cells["A1"].Formula = "=subtotal(3,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(5, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum103TimeInputs()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = "00:00";
				worksheet.Cells["B2"].Value = "12:00";
				worksheet.Cells["B3"].Value = "13:15";
				worksheet.Cells["B4"].Value = "2:00";
				worksheet.Cells["B5"].Value = "13:00";
				worksheet.Cells["A1"].Formula = "=subtotal(103,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(5, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum4TimeInputs()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = "00:00";
				worksheet.Cells["B2"].Value = "12:00";
				worksheet.Cells["B3"].Value = "13:15";
				worksheet.Cells["B4"].Value = "2:00";
				worksheet.Cells["B5"].Value = "13:00";
				worksheet.Cells["A1"].Formula = "=subtotal(4,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(0, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum104TimeInputs()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = "00:00";
				worksheet.Cells["B2"].Value = "12:00";
				worksheet.Cells["B3"].Value = "13:15";
				worksheet.Cells["B4"].Value = "2:00";
				worksheet.Cells["B5"].Value = "13:00";
				worksheet.Cells["A1"].Formula = "=subtotal(104,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(0, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum5TimeInputs()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = "00:00";
				worksheet.Cells["B2"].Value = "12:00";
				worksheet.Cells["B3"].Value = "13:15";
				worksheet.Cells["B4"].Value = "2:00";
				worksheet.Cells["B5"].Value = "13:00";
				worksheet.Cells["A1"].Formula = "=subtotal(5,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(0, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum105TimeInputs()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = "00:00";
				worksheet.Cells["B2"].Value = "12:00";
				worksheet.Cells["B3"].Value = "13:15";
				worksheet.Cells["B4"].Value = "2:00";
				worksheet.Cells["B5"].Value = "13:00";
				worksheet.Cells["A1"].Formula = "=subtotal(105,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(0, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum6TimeInputs()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = "00:00";
				worksheet.Cells["B2"].Value = "12:00";
				worksheet.Cells["B3"].Value = "13:15";
				worksheet.Cells["B4"].Value = "2:00";
				worksheet.Cells["B5"].Value = "13:00";
				worksheet.Cells["A1"].Formula = "=subtotal(6,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(0, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum106TimeInputs()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = "00:00";
				worksheet.Cells["B2"].Value = "12:00";
				worksheet.Cells["B3"].Value = "13:15";
				worksheet.Cells["B4"].Value = "2:00";
				worksheet.Cells["B5"].Value = "13:00";
				worksheet.Cells["A1"].Formula = "=subtotal(106,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(0, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum7TimeInputs()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = "00:00";
				worksheet.Cells["B2"].Value = "12:00";
				worksheet.Cells["B3"].Value = "13:15";
				worksheet.Cells["B4"].Value = "2:00";
				worksheet.Cells["B5"].Value = "13:00";
				worksheet.Cells["A1"].Formula = "=subtotal(7,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["A1"].Value).Type);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum107TimeInputs()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = "00:00";
				worksheet.Cells["B2"].Value = "12:00";
				worksheet.Cells["B3"].Value = "13:15";
				worksheet.Cells["B4"].Value = "2:00";
				worksheet.Cells["B5"].Value = "13:00";
				worksheet.Cells["A1"].Formula = "=subtotal(107,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["A1"].Value).Type);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum8TimeInputs()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = "00:00";
				worksheet.Cells["B2"].Value = "12:00";
				worksheet.Cells["B3"].Value = "13:15";
				worksheet.Cells["B4"].Value = "2:00";
				worksheet.Cells["B5"].Value = "13:00";
				worksheet.Cells["A1"].Formula = "=subtotal(8,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["A1"].Value).Type);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum108TimeInputs()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = "00:00";
				worksheet.Cells["B2"].Value = "12:00";
				worksheet.Cells["B3"].Value = "13:15";
				worksheet.Cells["B4"].Value = "2:00";
				worksheet.Cells["B5"].Value = "13:00";
				worksheet.Cells["A1"].Formula = "=subtotal(108,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["A1"].Value).Type);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum9TimeInputs()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = "00:00";
				worksheet.Cells["B2"].Value = "12:00";
				worksheet.Cells["B3"].Value = "13:15";
				worksheet.Cells["B4"].Value = "2:00";
				worksheet.Cells["B5"].Value = "13:00";
				worksheet.Cells["A1"].Formula = "=subtotal(9,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(0, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum109TimeInputs()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = "00:00";
				worksheet.Cells["B2"].Value = "12:00";
				worksheet.Cells["B3"].Value = "13:15";
				worksheet.Cells["B4"].Value = "2:00";
				worksheet.Cells["B5"].Value = "13:00";
				worksheet.Cells["A1"].Formula = "=subtotal(109,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(0, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum10TimeInputs()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = "00:00";
				worksheet.Cells["B2"].Value = "12:00";
				worksheet.Cells["B3"].Value = "13:15";
				worksheet.Cells["B4"].Value = "2:00";
				worksheet.Cells["B5"].Value = "13:00";
				worksheet.Cells["A1"].Formula = "=subtotal(10,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["A1"].Value).Type);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum110TimeInputs()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = "00:00";
				worksheet.Cells["B2"].Value = "12:00";
				worksheet.Cells["B3"].Value = "13:15";
				worksheet.Cells["B4"].Value = "2:00";
				worksheet.Cells["B5"].Value = "13:00";
				worksheet.Cells["A1"].Formula = "=subtotal(110,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["A1"].Value).Type);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum11TimeInputs()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = "00:00";
				worksheet.Cells["B2"].Value = "12:00";
				worksheet.Cells["B3"].Value = "13:15";
				worksheet.Cells["B4"].Value = "2:00";
				worksheet.Cells["B5"].Value = "13:00";
				worksheet.Cells["A1"].Formula = "=subtotal(11,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["A1"].Value).Type);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum111TimeInputs()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = "00:00";
				worksheet.Cells["B2"].Value = "12:00";
				worksheet.Cells["B3"].Value = "13:15";
				worksheet.Cells["B4"].Value = "2:00";
				worksheet.Cells["B5"].Value = "13:00";
				worksheet.Cells["A1"].Formula = "=subtotal(111,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["A1"].Value).Type);
			}
		}
		
		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum1DateTime()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = "1/12/2017  16:00";
				worksheet.Cells["B2"].Value = "5/12/2015  2:00:00";
				worksheet.Cells["B3"].Value = "3/20/2012  12:00";
				worksheet.Cells["B4"].Value = "12/12/2010  11:00";
				worksheet.Cells["B5"].Value = "1/11/2011  10:00";
				worksheet.Cells["A1"].Formula = "=subtotal(1,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["A1"].Value).Type);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum101DateTime()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = "1/12/2017  16:00";
				worksheet.Cells["B2"].Value = "5/12/2015  2:00:00";
				worksheet.Cells["B3"].Value = "3/20/2012  12:00";
				worksheet.Cells["B4"].Value = "12/12/2010  11:00";
				worksheet.Cells["B5"].Value = "1/11/2011  10:00";
				worksheet.Cells["A1"].Formula = "=subtotal(101,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["A1"].Value).Type);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum2DateTime()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = "1/12/2017  16:00";
				worksheet.Cells["B2"].Value = "5/12/2015  2:00:00";
				worksheet.Cells["B3"].Value = "3/20/2012  12:00";
				worksheet.Cells["B4"].Value = "12/12/2010  11:00";
				worksheet.Cells["B5"].Value = "1/11/2011  10:00";
				worksheet.Cells["A1"].Formula = "=subtotal(2,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(0, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum102DateTime()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = "1/12/2017  16:00";
				worksheet.Cells["B2"].Value = "5/12/2015  2:00:00";
				worksheet.Cells["B3"].Value = "3/20/2012  12:00";
				worksheet.Cells["B4"].Value = "12/12/2010  11:00";
				worksheet.Cells["B5"].Value = "1/11/2011  10:00";
				worksheet.Cells["A1"].Formula = "=subtotal(102,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(0, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum3DateTime()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = "1/12/2017  16:00";
				worksheet.Cells["B2"].Value = "5/12/2015  2:00:00";
				worksheet.Cells["B3"].Value = "3/20/2012  12:00";
				worksheet.Cells["B4"].Value = "12/12/2010  11:00";
				worksheet.Cells["B5"].Value = "1/11/2011  10:00";
				worksheet.Cells["A1"].Formula = "=subtotal(3,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(5, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum103DateTime()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = "1/12/2017  16:00";
				worksheet.Cells["B2"].Value = "5/12/2015  2:00:00";
				worksheet.Cells["B3"].Value = "3/20/2012  12:00";
				worksheet.Cells["B4"].Value = "12/12/2010  11:00";
				worksheet.Cells["B5"].Value = "1/11/2011  10:00";
				worksheet.Cells["A1"].Formula = "=subtotal(103,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(5, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum4DateTime()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = "1/12/2017  16:00";
				worksheet.Cells["B2"].Value = "5/12/2015  2:00:00";
				worksheet.Cells["B3"].Value = "3/20/2012  12:00";
				worksheet.Cells["B4"].Value = "12/12/2010  11:00";
				worksheet.Cells["B5"].Value = "1/11/2011  10:00";
				worksheet.Cells["A1"].Formula = "=subtotal(4,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(0, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum104DateTime()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = "1/12/2017  16:00";
				worksheet.Cells["B2"].Value = "5/12/2015  2:00:00";
				worksheet.Cells["B3"].Value = "3/20/2012  12:00";
				worksheet.Cells["B4"].Value = "12/12/2010  11:00";
				worksheet.Cells["B5"].Value = "1/11/2011  10:00";
				worksheet.Cells["A1"].Formula = "=subtotal(104,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(0, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum5DateTime()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = "1/12/2017  16:00";
				worksheet.Cells["B2"].Value = "5/12/2015  2:00:00";
				worksheet.Cells["B3"].Value = "3/20/2012  12:00";
				worksheet.Cells["B4"].Value = "12/12/2010  11:00";
				worksheet.Cells["B5"].Value = "1/11/2011  10:00";
				worksheet.Cells["A1"].Formula = "=subtotal(5,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(0, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum105DateTime()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = "1/12/2017  16:00";
				worksheet.Cells["B2"].Value = "5/12/2015  2:00:00";
				worksheet.Cells["B3"].Value = "3/20/2012  12:00";
				worksheet.Cells["B4"].Value = "12/12/2010  11:00";
				worksheet.Cells["B5"].Value = "1/11/2011  10:00";
				worksheet.Cells["A1"].Formula = "=subtotal(105,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(0, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum6DateTime()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = "1/12/2017  16:00";
				worksheet.Cells["B2"].Value = "5/12/2015  2:00:00";
				worksheet.Cells["B3"].Value = "3/20/2012  12:00";
				worksheet.Cells["B4"].Value = "12/12/2010  11:00";
				worksheet.Cells["B5"].Value = "1/11/2011  10:00";
				worksheet.Cells["A1"].Formula = "=subtotal(6,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(0, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum106DateTime()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = "1/12/2017  16:00";
				worksheet.Cells["B2"].Value = "5/12/2015  2:00:00";
				worksheet.Cells["B3"].Value = "3/20/2012  12:00";
				worksheet.Cells["B4"].Value = "12/12/2010  11:00";
				worksheet.Cells["B5"].Value = "1/11/2011  10:00";
				worksheet.Cells["A1"].Formula = "=subtotal(106,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(0, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum7DateTime()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = "1/12/2017  16:00";
				worksheet.Cells["B2"].Value = "5/12/2015  2:00:00";
				worksheet.Cells["B3"].Value = "3/20/2012  12:00";
				worksheet.Cells["B4"].Value = "12/12/2010  11:00";
				worksheet.Cells["B5"].Value = "1/11/2011  10:00";
				worksheet.Cells["A1"].Formula = "=subtotal(7,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["A1"].Value).Type);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum107DateTime()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = "1/12/2017  16:00";
				worksheet.Cells["B2"].Value = "5/12/2015  2:00:00";
				worksheet.Cells["B3"].Value = "3/20/2012  12:00";
				worksheet.Cells["B4"].Value = "12/12/2010  11:00";
				worksheet.Cells["B5"].Value = "1/11/2011  10:00";
				worksheet.Cells["A1"].Formula = "=subtotal(107,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["A1"].Value).Type);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum8DateTime()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = "1/12/2017  16:00";
				worksheet.Cells["B2"].Value = "5/12/2015  2:00:00";
				worksheet.Cells["B3"].Value = "3/20/2012  12:00";
				worksheet.Cells["B4"].Value = "12/12/2010  11:00";
				worksheet.Cells["B5"].Value = "1/11/2011  10:00";
				worksheet.Cells["A1"].Formula = "=subtotal(8,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["A1"].Value).Type);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum108DateTime()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = "1/12/2017  16:00";
				worksheet.Cells["B2"].Value = "5/12/2015  2:00:00";
				worksheet.Cells["B3"].Value = "3/20/2012  12:00";
				worksheet.Cells["B4"].Value = "12/12/2010  11:00";
				worksheet.Cells["B5"].Value = "1/11/2011  10:00";
				worksheet.Cells["A1"].Formula = "=subtotal(108,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["A1"].Value).Type); ;
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum9DateTime()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = "1/12/2017  16:00";
				worksheet.Cells["B2"].Value = "5/12/2015  2:00:00";
				worksheet.Cells["B3"].Value = "3/20/2012  12:00";
				worksheet.Cells["B4"].Value = "12/12/2010  11:00";
				worksheet.Cells["B5"].Value = "1/11/2011  10:00";
				worksheet.Cells["A1"].Formula = "=subtotal(9,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(0, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum109DateTime()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = "1/12/2017  16:00";
				worksheet.Cells["B2"].Value = "5/12/2015  2:00:00";
				worksheet.Cells["B3"].Value = "3/20/2012  12:00";
				worksheet.Cells["B4"].Value = "12/12/2010  11:00";
				worksheet.Cells["B5"].Value = "1/11/2011  10:00";
				worksheet.Cells["A1"].Formula = "=subtotal(109,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(0, (double)worksheet.Cells["A1"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum10DateTime()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = "1/12/2017  16:00";
				worksheet.Cells["B2"].Value = "5/12/2015  2:00:00";
				worksheet.Cells["B3"].Value = "3/20/2012  12:00";
				worksheet.Cells["B4"].Value = "12/12/2010  11:00";
				worksheet.Cells["B5"].Value = "1/11/2011  10:00";
				worksheet.Cells["A1"].Formula = "=subtotal(10,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["A1"].Value).Type);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum110DateTime()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = "1/12/2017  16:00";
				worksheet.Cells["B2"].Value = "5/12/2015  2:00:00";
				worksheet.Cells["B3"].Value = "3/20/2012  12:00";
				worksheet.Cells["B4"].Value = "12/12/2010  11:00";
				worksheet.Cells["B5"].Value = "1/11/2011  10:00";
				worksheet.Cells["A1"].Formula = "=subtotal(110,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["A1"].Value).Type);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum11DateTime()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = "1/12/2017  16:00";
				worksheet.Cells["B2"].Value = "5/12/2015  2:00:00";
				worksheet.Cells["B3"].Value = "3/20/2012  12:00";
				worksheet.Cells["B4"].Value = "12/12/2010  11:00";
				worksheet.Cells["B5"].Value = "1/11/2011  10:00";
				worksheet.Cells["A1"].Formula = "=subtotal(11,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["A1"].Value).Type);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum111DateTime()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = "1/12/2017  16:00";
				worksheet.Cells["B2"].Value = "5/12/2015  2:00:00";
				worksheet.Cells["B3"].Value = "3/20/2012  12:00";
				worksheet.Cells["B4"].Value = "12/12/2010  11:00";
				worksheet.Cells["B5"].Value = "1/11/2011  10:00";
				worksheet.Cells["A1"].Formula = "=subtotal(111,B1:B5)";
				worksheet.Calculate();
				Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)worksheet.Cells["A1"].Value).Type);
			}
		}

			[TestMethod]
			public void SubtotalIsGivenAListOfInputsFuntionNum1EmptyInputInFrontMiddleAndBack()
			{
				var function = new Subtotal();
				using (var package = new ExcelPackage())
				{
					var worksheet = package.Workbook.Worksheets.Add("Sheet1");

					worksheet.Cells["B2"].Value = 1;
					worksheet.Cells["B4"].Value = 7;
					worksheet.Cells["A1"].Formula = "=subtotal(1,B1:B5)";
					worksheet.Calculate();
					Assert.AreEqual(4, (double)worksheet.Cells["A1"].Value, .00001);
				}
			}

			[TestMethod]
			public void SubtotalIsGivenAListOfInputsFuntionNum101EmptyInputInFrontMiddleAndBack()
			{
				var function = new Subtotal();
				using (var package = new ExcelPackage())
				{
					var worksheet = package.Workbook.Worksheets.Add("Sheet1");

					worksheet.Cells["B2"].Value = 1;
					worksheet.Cells["B4"].Value = 7;
					worksheet.Cells["A1"].Formula = "=subtotal(101,B1:B5)";
					worksheet.Calculate();
					Assert.AreEqual(4, (double)worksheet.Cells["A1"].Value, .00001);
				}
			}

			[TestMethod]
			public void SubtotalIsGivenAListOfInputsFuntionNum2EmptyInputInFrontMiddleAndBack()
			{
				var function = new Subtotal();
				using (var package = new ExcelPackage())
				{
					var worksheet = package.Workbook.Worksheets.Add("Sheet1");

					worksheet.Cells["B2"].Value = 1;
					worksheet.Cells["B4"].Value = 7;	
					worksheet.Cells["A1"].Formula = "=subtotal(2,B1:B5)";
					worksheet.Calculate();
					Assert.AreEqual(2, (double)worksheet.Cells["A1"].Value, .00001);
				}
			}

			[TestMethod]
			public void SubtotalIsGivenAListOfInputsFuntionNum102EmptyInputInFrontMiddleAndBack()
			{
				var function = new Subtotal();
				using (var package = new ExcelPackage())
				{
					var worksheet = package.Workbook.Worksheets.Add("Sheet1");

					worksheet.Cells["B2"].Value = 1;
					worksheet.Cells["B4"].Value = 7;
					worksheet.Cells["A1"].Formula = "=subtotal(102,B1:B5)";
					worksheet.Calculate();
					Assert.AreEqual(2, (double)worksheet.Cells["A1"].Value, .00001);
				}
			}

			[TestMethod]
			public void SubtotalIsGivenAListOfInputsFuntionNum3EmptyInputInFrontMiddleAndBack()
			{
				var function = new Subtotal();
				using (var package = new ExcelPackage())
				{
					var worksheet = package.Workbook.Worksheets.Add("Sheet1");

					worksheet.Cells["B2"].Value = 1;
					worksheet.Cells["B4"].Value = 7;
					worksheet.Cells["A1"].Formula = "=subtotal(3,B1:B5)";
					worksheet.Calculate();
					Assert.AreEqual(2, (double)worksheet.Cells["A1"].Value, .00001);
				}
			}

			[TestMethod]
			public void SubtotalIsGivenAListOfInputsFuntionNum103EmptyInputInFrontMiddleAndBack()
			{
				var function = new Subtotal();
				using (var package = new ExcelPackage())
				{
					var worksheet = package.Workbook.Worksheets.Add("Sheet1");

					worksheet.Cells["B2"].Value = 1;
					worksheet.Cells["B4"].Value = 7;
					worksheet.Cells["A1"].Formula = "=subtotal(103,B1:B5)";
					worksheet.Calculate();
					Assert.AreEqual(2, (double)worksheet.Cells["A1"].Value, .00001);
				}
			}

			[TestMethod]
			public void SubtotalIsGivenAListOfInputsFuntionNum4EmptyInputInFrontMiddleAndBack()
			{
				var function = new Subtotal();
				using (var package = new ExcelPackage())
				{
					var worksheet = package.Workbook.Worksheets.Add("Sheet1");

					worksheet.Cells["B2"].Value = 1;
					worksheet.Cells["B4"].Value = 7;
					worksheet.Cells["A1"].Formula = "=subtotal(4,B1:B5)";
					worksheet.Calculate();
					Assert.AreEqual(7, (double)worksheet.Cells["A1"].Value, .00001);
				}
			}

			[TestMethod]
			public void SubtotalIsGivenAListOfInputsFuntionNum104EmptyInputInFrontMiddleAndBack()
			{
				var function = new Subtotal();
				using (var package = new ExcelPackage())
				{
					var worksheet = package.Workbook.Worksheets.Add("Sheet1");

					worksheet.Cells["B2"].Value = 1;
					worksheet.Cells["B4"].Value = 7;
					worksheet.Cells["A1"].Formula = "=subtotal(104,B1:B5)";
					worksheet.Calculate();
					Assert.AreEqual(7, (double)worksheet.Cells["A1"].Value, .00001);
				}
			}

			[TestMethod]
			public void SubtotalIsGivenAListOfInputsFuntionNum5EmptyInputInFrontMiddleAndBack()
			{
				var function = new Subtotal();
				using (var package = new ExcelPackage())
				{
					var worksheet = package.Workbook.Worksheets.Add("Sheet1");

					worksheet.Cells["B2"].Value = 1;
					worksheet.Cells["B4"].Value = 7;
					worksheet.Cells["A1"].Formula = "=subtotal(5,B1:B5)";
					worksheet.Calculate();
					Assert.AreEqual(1, (double)worksheet.Cells["A1"].Value, .00001);
				}
			}

			[TestMethod]
			public void SubtotalIsGivenAListOfInputsFuntionNum105EmptyInputInFrontMiddleAndBack()
			{
				var function = new Subtotal();
				using (var package = new ExcelPackage())
				{
					var worksheet = package.Workbook.Worksheets.Add("Sheet1");

					worksheet.Cells["B2"].Value = 1;
					worksheet.Cells["B4"].Value = 7;
					worksheet.Cells["A1"].Formula = "=subtotal(105,B1:B5)";
					worksheet.Calculate();
					Assert.AreEqual(1, (double)worksheet.Cells["A1"].Value, .00001);
				}
			}

			[TestMethod]
			public void SubtotalIsGivenAListOfInputsFuntionNum6EmptyInputInFrontMiddleAndBack()
			{
				var function = new Subtotal();
				using (var package = new ExcelPackage())
				{
					var worksheet = package.Workbook.Worksheets.Add("Sheet1");

					worksheet.Cells["B2"].Value = 1;
					worksheet.Cells["B4"].Value = 7;
					worksheet.Cells["A1"].Formula = "=subtotal(6,B1:B5)";
					worksheet.Calculate();
					Assert.AreEqual(7, (double)worksheet.Cells["A1"].Value, .00001);
				}
			}

			[TestMethod]
			public void SubtotalIsGivenAListOfInputsFuntionNum106EmptyInputInFrontMiddleAndBack()
			{
				var function = new Subtotal();
				using (var package = new ExcelPackage())
				{
					var worksheet = package.Workbook.Worksheets.Add("Sheet1");

					worksheet.Cells["B2"].Value = 1;
					worksheet.Cells["B4"].Value = 7;
					worksheet.Cells["A1"].Formula = "=subtotal(106,B1:B5)";
					worksheet.Calculate();
					Assert.AreEqual(7, (double)worksheet.Cells["A1"].Value, .00001);
				}
			}

			[TestMethod]
			public void SubtotalIsGivenAListOfInputsFuntionNum7EmptyInputInFrontMiddleAndBack()
			{
				var function = new Subtotal();
				using (var package = new ExcelPackage())
				{
					var worksheet = package.Workbook.Worksheets.Add("Sheet1");

					worksheet.Cells["B2"].Value = 1;
					worksheet.Cells["B4"].Value = 7;
					worksheet.Cells["A1"].Formula = "=subtotal(7,B1:B5)";
					worksheet.Calculate();
					Assert.AreEqual(4.242640687, (double)worksheet.Cells["A1"].Value, .00001);
				}
			}

			[TestMethod]
			public void SubtotalIsGivenAListOfInputsFuntionNum107EmptyInputInFrontMiddleAndBack()
			{
				var function = new Subtotal();
				using (var package = new ExcelPackage())
				{
					var worksheet = package.Workbook.Worksheets.Add("Sheet1");

					worksheet.Cells["B2"].Value = 1;
					worksheet.Cells["B4"].Value = 7;
					worksheet.Cells["A1"].Formula = "=subtotal(107,B1:B5)";
					worksheet.Calculate();
					Assert.AreEqual(4.242640687, (double)worksheet.Cells["A1"].Value, .00001);
				}
			}

			[TestMethod]
			public void SubtotalIsGivenAListOfInputsFuntionNum8EmptyInputInFrontMiddleAndBack()
			{
				var function = new Subtotal();
				using (var package = new ExcelPackage())
				{
					var worksheet = package.Workbook.Worksheets.Add("Sheet1");

					worksheet.Cells["B2"].Value = 1;
					worksheet.Cells["B4"].Value = 7;
					worksheet.Cells["A1"].Formula = "=subtotal(8,B1:B5)";
					worksheet.Calculate();
					Assert.AreEqual(3, (double)worksheet.Cells["A1"].Value, .00001);
				}
			}

			[TestMethod]
			public void SubtotalIsGivenAListOfInputsFuntionNum108EmptyInputInFrontMiddleAndBack()
			{
				var function = new Subtotal();
				using (var package = new ExcelPackage())
				{
					var worksheet = package.Workbook.Worksheets.Add("Sheet1");

					worksheet.Cells["B2"].Value = 1;
					worksheet.Cells["B4"].Value = 7;
					worksheet.Cells["A1"].Formula = "=subtotal(108,B1:B5)";
					worksheet.Calculate();
					Assert.AreEqual(3, (double)worksheet.Cells["A1"].Value, .00001);
				}
			}

			[TestMethod]
			public void SubtotalIsGivenAListOfInputsFuntionNum9EmptyInputInFrontMiddleAndBack()
			{
				var function = new Subtotal();
				using (var package = new ExcelPackage())
				{
					var worksheet = package.Workbook.Worksheets.Add("Sheet1");

					worksheet.Cells["B2"].Value = 1;
					worksheet.Cells["B4"].Value = 7;
					worksheet.Cells["A1"].Formula = "=subtotal(9,B1:B5)";
					worksheet.Calculate();
					Assert.AreEqual(8, (double)worksheet.Cells["A1"].Value, .00001);
				}
			}

			[TestMethod]
			public void SubtotalIsGivenAListOfInputsFuntionNum109EmptyInputInFrontMiddleAndBack()
			{
				var function = new Subtotal();
				using (var package = new ExcelPackage())
				{
					var worksheet = package.Workbook.Worksheets.Add("Sheet1");

					worksheet.Cells["B2"].Value = 1;
					worksheet.Cells["B4"].Value = 7;
					worksheet.Cells["A1"].Formula = "=subtotal(109,B1:B5)";
					worksheet.Calculate();
					Assert.AreEqual(8, (double)worksheet.Cells["A1"].Value, .00001);
				}
			}

			[TestMethod]
			public void SubtotalIsGivenAListOfInputsFuntionNum10EmptyInputInFrontMiddleAndBack()
			{
				var function = new Subtotal();
				using (var package = new ExcelPackage())
				{
					var worksheet = package.Workbook.Worksheets.Add("Sheet1");

					worksheet.Cells["B2"].Value = 1;
					worksheet.Cells["B4"].Value = 7;
					worksheet.Cells["A1"].Formula = "=subtotal(10,B1:B5)";
					worksheet.Calculate();
					Assert.AreEqual(18, (double)worksheet.Cells["A1"].Value, .00001);
				}
			}

			[TestMethod]
			public void SubtotalIsGivenAListOfInputsFuntionNum110EmptyInputInFrontMiddleAndBack()
			{
				var function = new Subtotal();
				using (var package = new ExcelPackage())
				{
					var worksheet = package.Workbook.Worksheets.Add("Sheet1");

					worksheet.Cells["B2"].Value = 1;
					worksheet.Cells["B4"].Value = 7;
					worksheet.Cells["A1"].Formula = "=subtotal(110,B1:B5)";
					worksheet.Calculate();
					Assert.AreEqual(18, (double)worksheet.Cells["A1"].Value, .00001);
				}
			}

			[TestMethod]
			public void SubtotalIsGivenAListOfInputsFuntionNum11EmptyInputInFrontMiddleAndBack()
			{
				var function = new Subtotal();
				using (var package = new ExcelPackage())
				{
					var worksheet = package.Workbook.Worksheets.Add("Sheet1");

					worksheet.Cells["B2"].Value = 1;
					worksheet.Cells["B4"].Value = 7;
					worksheet.Cells["A1"].Formula = "=subtotal(11,B1:B5)";
					worksheet.Calculate();
					Assert.AreEqual(9, (double)worksheet.Cells["A1"].Value, .00001);
				}
			}

			[TestMethod]
			public void SubtotalIsGivenAListOfInputsFuntionNum111EmptyInputInFrontMiddleAndBack()
			{
				var function = new Subtotal();
				using (var package = new ExcelPackage())
				{
					var worksheet = package.Workbook.Worksheets.Add("Sheet1");

					worksheet.Cells["B2"].Value = 1;
					worksheet.Cells["B4"].Value = 7;
					worksheet.Cells["A1"].Formula = "=subtotal(111,B1:B5)";
					worksheet.Calculate();
					Assert.AreEqual(9, (double)worksheet.Cells["A1"].Value, .00001);
				}
			}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum1TestingNested()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["B4"].Value = 5;
				worksheet.Cells["B5"].Value = 10;
				worksheet.Cells["B6"].Formula = "=subtotal(1,B1:B5)";
				worksheet.Cells["B7"].Formula = "=subtotal(1,B1:B6)";
				worksheet.Calculate();
				Assert.AreEqual(4.2, (double)worksheet.Cells["B6"].Value, .00001);
				Assert.AreEqual(4.2, (double)worksheet.Cells["B7"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum6TestTwoTestingNested()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = -1;
				worksheet.Cells["B2"].Value = -2;
				worksheet.Cells["B3"].Value = -3;
				worksheet.Cells["B4"].Formula = "=subtotal(6,B1:B3)";
				worksheet.Cells["B5"].Formula = "=subtotal(6,B1:B4)";
				worksheet.Calculate();
				Assert.AreEqual(-6, (double)worksheet.Cells["B4"].Value, .00001);
				Assert.AreEqual(-6, (double)worksheet.Cells["B5"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum101TestingNested()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["B4"].Value = 5;
				worksheet.Cells["B5"].Value = 10;
				worksheet.Cells["B6"].Formula = "=subtotal(101,B1:B5)";
				worksheet.Cells["B7"].Formula = "=subtotal(101,B1:B6)";
				worksheet.Calculate();
				Assert.AreEqual(4.2, (double)worksheet.Cells["B6"].Value, .00001);
				Assert.AreEqual(4.2, (double)worksheet.Cells["B7"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum2TestingNested()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["B4"].Value = 5;
				worksheet.Cells["B5"].Value = 10;
				worksheet.Cells["B6"].Formula = "=subtotal(2,B1:B5)";
				worksheet.Cells["B7"].Formula = "=subtotal(2,B1:B6)";
				worksheet.Calculate();
				Assert.AreEqual(5, (double)worksheet.Cells["B6"].Value, .00001);
				Assert.AreEqual(5, (double)worksheet.Cells["B7"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum102TestingNested()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["B4"].Value = 5;
				worksheet.Cells["B5"].Value = 10;
				worksheet.Cells["B6"].Formula = "=subtotal(102,B1:B5)";
				worksheet.Cells["B7"].Formula = "=subtotal(102,B1:B6)";
				worksheet.Calculate();
				Assert.AreEqual(5, (double)worksheet.Cells["B6"].Value, .00001);
				Assert.AreEqual(5, (double)worksheet.Cells["B7"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum3TestingNested()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["B4"].Value = 5;
				worksheet.Cells["B5"].Value = 10;
				worksheet.Cells["B6"].Formula = "=subtotal(3,B1:B5)";
				worksheet.Cells["B7"].Formula = "=subtotal(3,B1:B6)";
				worksheet.Calculate();
				Assert.AreEqual(5, (double)worksheet.Cells["B6"].Value, .00001);
				Assert.AreEqual(5, (double)worksheet.Cells["B7"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum103TestingNested()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["B4"].Value = 5;
				worksheet.Cells["B5"].Value = 10;
				worksheet.Cells["B6"].Formula = "=subtotal(103,B1:B5)";
				worksheet.Cells["B7"].Formula = "=subtotal(103,B1:B6)";
				worksheet.Calculate();
				Assert.AreEqual(5, (double)worksheet.Cells["B6"].Value, .00001);
				Assert.AreEqual(5, (double)worksheet.Cells["B7"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum4TestingNested()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["B4"].Value = 5;
				worksheet.Cells["B5"].Value = 10;
				worksheet.Cells["B6"].Formula = "=subtotal(4,B1:B5)";
				worksheet.Cells["B7"].Formula = "=subtotal(4,B1:B6)";
				worksheet.Calculate();
				Assert.AreEqual(10, (double)worksheet.Cells["B6"].Value, .00001);
				Assert.AreEqual(10, (double)worksheet.Cells["B7"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum104TestingNested()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["B4"].Value = 5;
				worksheet.Cells["B5"].Value = 10;
				worksheet.Cells["B6"].Formula = "=subtotal(104,B1:B5)";
				worksheet.Cells["B7"].Formula = "=subtotal(104,B1:B6)";
				worksheet.Calculate();
				Assert.AreEqual(10, (double)worksheet.Cells["B6"].Value, .00001);
				Assert.AreEqual(10, (double)worksheet.Cells["B7"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum5TestingNested()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["B4"].Value = 5;
				worksheet.Cells["B5"].Value = 10;
				worksheet.Cells["B6"].Formula = "=subtotal(5,B1:B5)";
				worksheet.Cells["B7"].Formula = "=subtotal(5,B1:B6)";
				worksheet.Calculate();
				Assert.AreEqual(1, (double)worksheet.Cells["B6"].Value, .00001);
				Assert.AreEqual(1, (double)worksheet.Cells["B7"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum105TestingNested()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["B4"].Value = 5;
				worksheet.Cells["B5"].Value = 10;
				worksheet.Cells["B6"].Formula = "=subtotal(105,B1:B5)";
				worksheet.Cells["B7"].Formula = "=subtotal(105,B1:B6)";
				worksheet.Calculate();
				Assert.AreEqual(1, (double)worksheet.Cells["B6"].Value, .00001);
				Assert.AreEqual(1, (double)worksheet.Cells["B7"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum6TestingNested()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["B4"].Value = 5;
				worksheet.Cells["B5"].Value = 10;
				worksheet.Cells["B6"].Formula = "=subtotal(6,B1:B5)";
				worksheet.Cells["B7"].Formula = "=subtotal(6,B1:B6)";
				worksheet.Calculate();
				Assert.AreEqual(300, (double)worksheet.Cells["B6"].Value, .00001);
				Assert.AreEqual(300, (double)worksheet.Cells["B7"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum106TestingNested()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["B4"].Value = 5;
				worksheet.Cells["B5"].Value = 10;
				worksheet.Cells["B6"].Formula = "=subtotal(106,B1:B5)";
				worksheet.Cells["B7"].Formula = "=subtotal(106,B1:B6)";
				worksheet.Calculate();
				Assert.AreEqual(300, (double)worksheet.Cells["B6"].Value, .00001);
				Assert.AreEqual(300, (double)worksheet.Cells["B7"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum7TestingNested()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["B4"].Value = 5;
				worksheet.Cells["B5"].Value = 10;
				worksheet.Cells["B6"].Formula = "=subtotal(7,B1:B5)";
				worksheet.Cells["B7"].Formula = "=subtotal(7,B1:B6)";
				worksheet.Calculate();
				Assert.AreEqual(3.563705936, (double)worksheet.Cells["B6"].Value, .00001);
				Assert.AreEqual(3.563705936, (double)worksheet.Cells["B7"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum107TestingNested()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["B4"].Value = 5;
				worksheet.Cells["B5"].Value = 10;
				worksheet.Cells["B6"].Formula = "=subtotal(107,B1:B5)";
				worksheet.Cells["B7"].Formula = "=subtotal(107,B1:B6)";
				worksheet.Calculate();
				Assert.AreEqual(3.563705936, (double)worksheet.Cells["B6"].Value, .00001);
				Assert.AreEqual(3.563705936, (double)worksheet.Cells["B7"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum8TestingNested()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["B4"].Value = 5;
				worksheet.Cells["B5"].Value = 10;
				worksheet.Cells["B6"].Formula = "=subtotal(8,B1:B5)";
				worksheet.Cells["B7"].Formula = "=subtotal(8,B1:B6)";
				worksheet.Calculate();
				Assert.AreEqual(3.18747549, (double)worksheet.Cells["B6"].Value, .00001);
				Assert.AreEqual(3.18747549, (double)worksheet.Cells["B7"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum108TestingNested()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["B4"].Value = 5;
				worksheet.Cells["B5"].Value = 10;
				worksheet.Cells["B6"].Formula = "=subtotal(108,B1:B5)";
				worksheet.Cells["B7"].Formula = "=subtotal(108,B1:B6)";
				worksheet.Calculate();
				Assert.AreEqual(3.18747549, (double)worksheet.Cells["B6"].Value, .00001);
				Assert.AreEqual(3.18747549, (double)worksheet.Cells["B7"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum9TestingNested()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["B4"].Value = 5;
				worksheet.Cells["B5"].Value = 10;
				worksheet.Cells["B6"].Formula = "=subtotal(9,B1:B5)";
				worksheet.Cells["B7"].Formula = "=subtotal(9,B1:B6)";
				worksheet.Calculate();
				Assert.AreEqual(21, (double)worksheet.Cells["B6"].Value, .00001);
				Assert.AreEqual(21, (double)worksheet.Cells["B7"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum109TestingNested()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["B4"].Value = 5;
				worksheet.Cells["B5"].Value = 10;
				worksheet.Cells["B6"].Formula = "=subtotal(109,B1:B5)";
				worksheet.Cells["B7"].Formula = "=subtotal(109,B1:B6)";
				worksheet.Calculate();
				Assert.AreEqual(21, (double)worksheet.Cells["B6"].Value, .00001);
				Assert.AreEqual(21, (double)worksheet.Cells["B7"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum10TestingNested()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["B4"].Value = 5;
				worksheet.Cells["B5"].Value = 10;
				worksheet.Cells["B6"].Formula = "=subtotal(10,B1:B5)";
				worksheet.Cells["B7"].Formula = "=subtotal(10,B1:B6)";
				worksheet.Calculate();
				Assert.AreEqual(12.7, (double)worksheet.Cells["B6"].Value, .00001);
				Assert.AreEqual(12.7, (double)worksheet.Cells["B7"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum110TestingNested()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["B4"].Value = 5;
				worksheet.Cells["B5"].Value = 10;
				worksheet.Cells["B6"].Formula = "=subtotal(110,B1:B5)";
				worksheet.Cells["B7"].Formula = "=subtotal(110,B1:B6)";
				worksheet.Calculate();
				Assert.AreEqual(12.7, (double)worksheet.Cells["B6"].Value, .00001);
				Assert.AreEqual(12.7, (double)worksheet.Cells["B7"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum11TestingNested()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["B4"].Value = 5;
				worksheet.Cells["B5"].Value = 10;
				worksheet.Cells["B6"].Formula = "=subtotal(11,B1:B5)";
				worksheet.Cells["B7"].Formula = "=subtotal(11,B1:B6)";
				worksheet.Calculate();
				Assert.AreEqual(10.16, (double)worksheet.Cells["B6"].Value, .00001);
				Assert.AreEqual(10.16, (double)worksheet.Cells["B7"].Value, .00001);
			}
		}

		[TestMethod]
		public void SubtotalIsGivenAListOfInputsFuntionNum111TestingNested()
		{
			var function = new Subtotal();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Sheet1");

				worksheet.Cells["B1"].Value = 1;
				worksheet.Cells["B2"].Value = 2;
				worksheet.Cells["B3"].Value = 3;
				worksheet.Cells["B4"].Value = 5;
				worksheet.Cells["B5"].Value = 10;
				worksheet.Cells["B6"].Formula = "=subtotal(111,B1:B5)";
				worksheet.Cells["B7"].Formula = "=subtotal(111,B1:B6)";
				worksheet.Calculate();
				Assert.AreEqual(10.16, (double)worksheet.Cells["B6"].Value, .00001);
				Assert.AreEqual(10.16, (double)worksheet.Cells["B7"].Value, .00001);
			}
		}
		#endregion

		#region Subtotal Tests Found in EPPlusTest.Excel.Functions

		private ParsingContext _context;
		private ExcelWorksheet _worksheet;
		private ExcelPackage _package;

		[TestInitialize]
		public void Setup()
		{
			_context = ParsingContext.Create();
			_context.Scopes.NewScope(RangeAddress.Empty);
			_package = new ExcelPackage(new MemoryStream());
			_worksheet = _package.Workbook.Worksheets.Add("Test");
		}

		[TestMethod]
		public void ShouldPoundValueIfInvalidFuncNumber()
		{
			var func = new Subtotal();
			var args = FunctionsHelper.CreateArgs(139, 1);
			var result = func.Execute(args, _context);
			Assert.AreEqual(OfficeOpenXml.FormulaParsing.ExpressionGraph.DataType.ExcelError, result.DataType);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)(result.Result)).Type);
		}

		[TestMethod]
		public void ShouldCalculateAverageWhenCalcTypeIs1()
		{
			var func = new Subtotal();
			var args = FunctionsHelper.CreateArgs(1, 10, 20, 30, 40, 50);
			var result = func.Execute(args, _context);
			Assert.AreEqual(30d, result.Result);
		}

		[TestMethod]
		public void ShouldCalculateAverageWhenCalcTypeIs1WithZero()
		{
			var func = new Subtotal();
			var args = FunctionsHelper.CreateArgs(1, 10, 20, 30, 40, 0);
			var result = func.Execute(args, _context);
			Assert.AreEqual(20d, result.Result);
		}

		[TestMethod]
		public void ShouldCalculateCountWhenCalcTypeIs2()
		{
			var func = new Subtotal();
			var args = FunctionsHelper.CreateArgs(2, 10, 20, 30, 40, 50);
			var result = func.Execute(args, _context);
			Assert.AreEqual(5d, result.Result);
		}

		[TestMethod]
		public void ShouldCalculateCountAWhenCalcTypeIs3()
		{
			var func = new Subtotal();
			var args = FunctionsHelper.CreateArgs(3, 10, 20, 30, 40, 50);
			var result = func.Execute(args, _context);
			Assert.AreEqual(5d, result.Result);
		}

		[TestMethod]
		public void ShouldCalculateMaxWhenCalcTypeIs4()
		{
			var func = new Subtotal();
			var args = FunctionsHelper.CreateArgs(4, 10, 20, 30, 40, 50);
			var result = func.Execute(args, _context);
			Assert.AreEqual(50d, result.Result);
		}

		[TestMethod]
		public void ShouldCalculateMinWhenCalcTypeIs5()
		{
			var func = new Subtotal();
			var args = FunctionsHelper.CreateArgs(5, 10, 20, 30, 40, 50);
			var result = func.Execute(args, _context);
			Assert.AreEqual(10d, result.Result);
		}

		[TestMethod]
		public void ShouldCalculateProductWhenCalcTypeIs6()
		{
			var func = new Subtotal();
			var args = FunctionsHelper.CreateArgs(6, 10, 20, 30, 40, 50);
			var result = func.Execute(args, _context);
			Assert.AreEqual(12000000d, result.Result);
		}

		[TestMethod]
		public void ShouldCalculateStdevWhenCalcTypeIs7()
		{
			var func = new Subtotal();
			var args = FunctionsHelper.CreateArgs(7, 10, 20, 30, 40, 50);
			var result = func.Execute(args, _context);
			var resultRounded = System.Math.Round((double)result.Result, 5);
			Assert.AreEqual(15.81139d, resultRounded);
		}

		[TestMethod]
		public void ShouldCalculateStdevPWhenCalcTypeIs8()
		{
			var func = new Subtotal();
			var args = FunctionsHelper.CreateArgs(8, 10, 20, 30, 40, 50);
			var result = func.Execute(args, _context);
			var resultRounded = System.Math.Round((double)result.Result, 8);
			Assert.AreEqual(14.14213562, resultRounded);
		}

		[TestMethod]
		public void ShouldCalculateSumWhenCalcTypeIs9()
		{
			var func = new Subtotal();
			var args = FunctionsHelper.CreateArgs(9, 10, 20, 30, 40, 50);
			var result = func.Execute(args, _context);
			Assert.AreEqual(150d, result.Result);
		}

		[TestMethod]
		public void ShouldCalculateVarWhenCalcTypeIs10()
		{
			var func = new Subtotal();
			var args = FunctionsHelper.CreateArgs(10, 10, 20, 30, 40, 50);
			var result = func.Execute(args, _context);
			Assert.AreEqual(250d, result.Result);
		}

		[TestMethod]
		public void ShouldCalculateVarPWhenCalcTypeIs11()
		{
			var func = new Subtotal();
			var args = FunctionsHelper.CreateArgs(11, 10, 20, 30, 40, 50);
			var result = func.Execute(args, _context);
			Assert.AreEqual(200d, result.Result);
		}

		[TestMethod]
		public void ShouldCalculateAverageWhenCalcTypeIs101()
		{
			var func = new Subtotal();
			var args = FunctionsHelper.CreateArgs(101, 10, 20, 30, 40, 50);
			var result = func.Execute(args, _context);
			Assert.AreEqual(30d, result.Result);
		}

		[TestMethod]
		public void ShouldCalculateCountWhenCalcTypeIs102()
		{
			var func = new Subtotal();
			var args = FunctionsHelper.CreateArgs(102, 10, 20, 30, 40, 50);
			var result = func.Execute(args, _context);
			Assert.AreEqual(5d, result.Result);
		}

		[TestMethod]
		public void ShouldCalculateCountAWhenCalcTypeIs103()
		{
			var func = new Subtotal();
			var args = FunctionsHelper.CreateArgs(103, 10, 20, 30, 40, 50);
			var result = func.Execute(args, _context);
			Assert.AreEqual(5d, result.Result);
		}

		[TestMethod]
		public void ShouldCalculateMaxWhenCalcTypeIs104()
		{
			var func = new Subtotal();
			var args = FunctionsHelper.CreateArgs(104, 10, 20, 30, 40, 50);
			var result = func.Execute(args, _context);
			Assert.AreEqual(50d, result.Result);
		}

		[TestMethod]
		public void ShouldCalculateMinWhenCalcTypeIs105()
		{
			var func = new Subtotal();
			var args = FunctionsHelper.CreateArgs(105, 10, 20, 30, 40, 50);
			var result = func.Execute(args, _context);
			Assert.AreEqual(10d, result.Result);
		}

		[TestMethod]
		public void ShouldCalculateProductWhenCalcTypeIs106()
		{
			var func = new Subtotal();
			var args = FunctionsHelper.CreateArgs(106, 10, 20, 30, 40, 50);
			var result = func.Execute(args, _context);
			Assert.AreEqual(12000000d, result.Result);
		}

		[TestMethod]
		public void ShouldCalculateStdevWhenCalcTypeIs107()
		{
			var func = new Subtotal();
			var args = FunctionsHelper.CreateArgs(107, 10, 20, 30, 40, 50);
			var result = func.Execute(args, _context);
			var resultRounded = System.Math.Round((double)result.Result, 5);
			Assert.AreEqual(15.81139d, resultRounded);
		}

		[TestMethod]
		public void ShouldCalculateStdevPWhenCalcTypeIs108()
		{
			var func = new Subtotal();
			var args = FunctionsHelper.CreateArgs(108, 10, 20, 30, 40, 50);
			var result = func.Execute(args, _context);
			var resultRounded = System.Math.Round((double)result.Result, 8);
			Assert.AreEqual(14.14213562, resultRounded);
		}

		[TestMethod]
		public void ShouldCalculateSumWhenCalcTypeIs109()
		{
			var func = new Subtotal();
			var args = FunctionsHelper.CreateArgs(109, 10, 20, 30, 40, 50);
			var result = func.Execute(args, _context);
			Assert.AreEqual(150d, result.Result);
		}

		[TestMethod]
		public void ShouldCalculateVarWhenCalcTypeIs110()
		{
			var func = new Subtotal();
			var args = FunctionsHelper.CreateArgs(110, 10, 20, 30, 40, 50, 51);
			args.Last().SetExcelStateFlag(ExcelCellState.HiddenCell);
			var result = func.Execute(args, _context);
			Assert.AreEqual(250d, result.Result);
		}

		[TestMethod]
		public void ShouldCalculateVarPWhenCalcTypeIs111()
		{
			var func = new Subtotal();
			var args = FunctionsHelper.CreateArgs(111, 10, 20, 30, 40, 50);
			var result = func.Execute(args, _context);
			Assert.AreEqual(200d, result.Result);
		}
		#endregion

		#region Subtotal Tests Found in EPPlusTest.FormulaParsing.IntegrationTests.BuiltInFunctions

		[TestCleanup]
		public void Cleanup()
		{
			_package.Dispose();
		}

		[TestMethod]
		public void SubtotalShouldNotIncludeSubtotalChildren_Avg()
		{
			_worksheet.Cells["A1"].Formula = "SUBTOTAL(1, A2:A3)";
			_worksheet.Cells["A2"].Formula = "SUBTOTAL(9, A5:A6)";
			_worksheet.Cells["A3"].Value = 2d;
			_worksheet.Cells["A5"].Value = 2d;
			_worksheet.Calculate();
			var result = _worksheet.Cells["A1"].Value;
			Assert.AreEqual(2d, result);
		}

		[TestMethod]
		public void SubtotalShouldNotIncludeSubtotalChildren_Count()
		{
			_worksheet.Cells["A1"].Formula = "SUBTOTAL(2, A2:A3)";
			_worksheet.Cells["A2"].Formula = "SUBTOTAL(9, A5:A6)";
			_worksheet.Cells["A3"].Value = 2d;
			_worksheet.Cells["A5"].Value = 2d;
			_worksheet.Calculate();
			var result = _worksheet.Cells["A1"].Value;
			Assert.AreEqual(1d, result);
		}

		[TestMethod]
		public void SubtotalShouldNotIncludeSubtotalChildren_CountA()
		{
			_worksheet.Cells["A1"].Formula = "SUBTOTAL(3, A2:A3)";
			_worksheet.Cells["A2"].Formula = "SUBTOTAL(9, A5:A6)";
			_worksheet.Cells["A3"].Value = 2d;
			_worksheet.Cells["A5"].Value = 2d;
			_worksheet.Calculate();
			var result = _worksheet.Cells["A1"].Value;
			Assert.AreEqual(1d, result);
		}

		[TestMethod]
		public void SubtotalShouldNotIncludeSubtotalChildren_Max()
		{
			_worksheet.Cells["A1"].Formula = "SUBTOTAL(4, A2:A3)";
			_worksheet.Cells["A2"].Formula = "SUBTOTAL(9, A5:A6)";
			_worksheet.Cells["A3"].Value = 2d;
			_worksheet.Cells["A5"].Value = 2d;
			_worksheet.Calculate();
			var result = _worksheet.Cells["A1"].Value;
			Assert.AreEqual(2d, result);
		}

		[TestMethod]
		public void SubtotalShouldNotIncludeSubtotalChildren_Min()
		{
			_worksheet.Cells["A1"].Formula = "SUBTOTAL(5, A2:A3)";
			_worksheet.Cells["A2"].Formula = "SUBTOTAL(9, A5:A6)";
			_worksheet.Cells["A3"].Value = 2d;
			_worksheet.Cells["A5"].Value = 2d;
			_worksheet.Calculate();
			var result = _worksheet.Cells["A1"].Value;
			Assert.AreEqual(2d, result);
		}

		[TestMethod]
		public void SubtotalShouldNotIncludeSubtotalChildren_Product()
		{
			_worksheet.Cells["A1"].Formula = "SUBTOTAL(6, A2:A3)";
			_worksheet.Cells["A2"].Formula = "SUBTOTAL(9, A5:A6)";
			_worksheet.Cells["A3"].Value = 2d;
			_worksheet.Cells["A5"].Value = 2d;
			_worksheet.Calculate();
			var result = _worksheet.Cells["A1"].Value;
			Assert.AreEqual(2d, result);
		}

		[TestMethod]
		public void SubtotalShouldNotIncludeSubtotalChildren_Stdev()
		{
			_worksheet.Cells["A1"].Formula = "SUBTOTAL(7, A2:A4)";
			_worksheet.Cells["A2"].Formula = "SUBTOTAL(7, A5:A6)";
			_worksheet.Cells["A3"].Value = 5d;
			_worksheet.Cells["A4"].Value = 4d;
			_worksheet.Cells["A5"].Value = 2d;
			_worksheet.Cells["A6"].Value = 4d;
			_worksheet.Calculate();
			var result = _worksheet.Cells["A1"].Value;
			result = System.Math.Round((double)result, 9);
			Assert.AreEqual(0.707106781d, result);
		}

		[TestMethod]
		public void SubtotalShouldNotIncludeSubtotalChildren_StdevP()
		{
			_worksheet.Cells["A1"].Formula = "SUBTOTAL(8, A2:A4)";
			_worksheet.Cells["A2"].Formula = "SUBTOTAL(8, A5:A6)";
			_worksheet.Cells["A3"].Value = 5d;
			_worksheet.Cells["A4"].Value = 4d;
			_worksheet.Cells["A5"].Value = 2d;
			_worksheet.Cells["A6"].Value = 4d;
			_worksheet.Calculate();
			var result = _worksheet.Cells["A1"].Value;
			Assert.AreEqual(0.5d, result);
		}

		[TestMethod]
		public void SubtotalShouldNotIncludeSubtotalChildren_Sum()
		{
			_worksheet.Cells["A1"].Formula = "SUBTOTAL(9, A2:A3)";
			_worksheet.Cells["A2"].Formula = "SUBTOTAL(9, A5:A6)";
			_worksheet.Cells["A3"].Value = 2d;
			_worksheet.Cells["A5"].Value = 2d;
			_worksheet.Calculate();
			var result = _worksheet.Cells["A1"].Value;
			Assert.AreEqual(2d, result);
		}

		[TestMethod]
		public void SubtotalShouldNotIncludeSubtotalChildren_Var()
		{
			_worksheet.Cells["A1"].Formula = "SUBTOTAL(9, A2:A4)";
			_worksheet.Cells["A2"].Formula = "SUBTOTAL(8, A5:A6)";
			_worksheet.Cells["A3"].Value = 5d;
			_worksheet.Cells["A4"].Value = 4d;
			_worksheet.Cells["A5"].Value = 2d;
			_worksheet.Cells["A6"].Value = 4d;
			_worksheet.Calculate();
			var result = _worksheet.Cells["A1"].Value;
			Assert.AreEqual(9d, result);
		}

		[TestMethod]
		public void SubtotalShouldNotIncludeSubtotalChildren_VarP()
		{
			_worksheet.Cells["A1"].Formula = "SUBTOTAL(10, A2:A4)";
			_worksheet.Cells["A2"].Formula = "SUBTOTAL(8, A5:A6)";
			_worksheet.Cells["A3"].Value = 5d;
			_worksheet.Cells["A4"].Value = 4d;
			_worksheet.Cells["A5"].Value = 2d;
			_worksheet.Cells["A6"].Value = 4d;
			_worksheet.Calculate();
			var result = _worksheet.Cells["A1"].Value;
			Assert.AreEqual(0.5d, result);
		}
		#endregion

	#endregion
	}
}
