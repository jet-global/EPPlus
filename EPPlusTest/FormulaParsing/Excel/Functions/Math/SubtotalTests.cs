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
using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math
{

	[TestClass]
	public class SubtotalTests : MathFunctionsTestBase
	{
		#region Subtotal Function(Execute) Tests

		[TestMethod]
		public void SubtotalIsGeivenAListOfInputsFuntionNum1()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum6TestTwo()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum101()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum2()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum102()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum3()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum103()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum4()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum104()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum5()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum105()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum6()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum106()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum7()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum107()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum8()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum108()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum9()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum109()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum10()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum110()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum11()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum111()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum1TestingBools()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum101TestingBools()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum2TestingBools()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum102TestingBools()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum3TestingBools()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum103TestingBools()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum4TestingBools()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum104TestingBools()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum5TestingBools()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum105TestingBools()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum6TestingBools()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum106TestingBools()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum7TestingBools()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum107TestingBools()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum8TestingBools()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum108TestingBools()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum9TestingBools()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum109TestingBools()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum10TestingBools()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum110TestingBools()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum11TestingBools()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum111TestingBools()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum1TimeInputs()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum101TimeInputs()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum2TimeInputs()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum102TimeInputs()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum3TimeInputs()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum103TimeInputs()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum4TimeInputs()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum104TimeInputs()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum5TimeInputs()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum105TimeInputs()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum6TimeInputs()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum106TimeInputs()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum7TimeInputs()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum107TimeInputs()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum8TimeInputs()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum108TimeInputs()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum9TimeInputs()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum109TimeInputs()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum10TimeInputs()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum110TimeInputs()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum11TimeInputs()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum111TimeInputs()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum1DateTime()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum101DateTime()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum2DateTime()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum102DateTime()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum3DateTime()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum103DateTime()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum4DateTime()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum104DateTime()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum5DateTime()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum105DateTime()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum6DateTime()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum106DateTime()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum7DateTime()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum107DateTime()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum8DateTime()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum108DateTime()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum9DateTime()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum109DateTime()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum10DateTime()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum110DateTime()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum11DateTime()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum111DateTime()
		{
			var function = new VarP();
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
			public void SubtotalIsGeivenAListOfInputsFuntionNum1EmptyInputInFrontMiddleAndBack()
			{
				var function = new VarP();
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
			public void SubtotalIsGeivenAListOfInputsFuntionNum101EmptyInputInFrontMiddleAndBack()
			{
				var function = new VarP();
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
			public void SubtotalIsGeivenAListOfInputsFuntionNum2EmptyInputInFrontMiddleAndBack()
			{
				var function = new VarP();
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
			public void SubtotalIsGeivenAListOfInputsFuntionNum102EmptyInputInFrontMiddleAndBack()
			{
				var function = new VarP();
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
			public void SubtotalIsGeivenAListOfInputsFuntionNum3EmptyInputInFrontMiddleAndBack()
			{
				var function = new VarP();
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
			public void SubtotalIsGeivenAListOfInputsFuntionNum103EmptyInputInFrontMiddleAndBack()
			{
				var function = new VarP();
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
			public void SubtotalIsGeivenAListOfInputsFuntionNum4EmptyInputInFrontMiddleAndBack()
			{
				var function = new VarP();
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
			public void SubtotalIsGeivenAListOfInputsFuntionNum104EmptyInputInFrontMiddleAndBack()
			{
				var function = new VarP();
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
			public void SubtotalIsGeivenAListOfInputsFuntionNum5EmptyInputInFrontMiddleAndBack()
			{
				var function = new VarP();
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
			public void SubtotalIsGeivenAListOfInputsFuntionNum105EmptyInputInFrontMiddleAndBack()
			{
				var function = new VarP();
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
			public void SubtotalIsGeivenAListOfInputsFuntionNum6EmptyInputInFrontMiddleAndBack()
			{
				var function = new VarP();
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
			public void SubtotalIsGeivenAListOfInputsFuntionNum106EmptyInputInFrontMiddleAndBack()
			{
				var function = new VarP();
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
			public void SubtotalIsGeivenAListOfInputsFuntionNum7EmptyInputInFrontMiddleAndBack()
			{
				var function = new VarP();
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
			public void SubtotalIsGeivenAListOfInputsFuntionNum107EmptyInputInFrontMiddleAndBack()
			{
				var function = new VarP();
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
			public void SubtotalIsGeivenAListOfInputsFuntionNum8EmptyInputInFrontMiddleAndBack()
			{
				var function = new VarP();
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
			public void SubtotalIsGeivenAListOfInputsFuntionNum108EmptyInputInFrontMiddleAndBack()
			{
				var function = new VarP();
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
			public void SubtotalIsGeivenAListOfInputsFuntionNum9EmptyInputInFrontMiddleAndBack()
			{
				var function = new VarP();
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
			public void SubtotalIsGeivenAListOfInputsFuntionNum109EmptyInputInFrontMiddleAndBack()
			{
				var function = new VarP();
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
			public void SubtotalIsGeivenAListOfInputsFuntionNum10EmptyInputInFrontMiddleAndBack()
			{
				var function = new VarP();
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
			public void SubtotalIsGeivenAListOfInputsFuntionNum110EmptyInputInFrontMiddleAndBack()
			{
				var function = new VarP();
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
			public void SubtotalIsGeivenAListOfInputsFuntionNum11EmptyInputInFrontMiddleAndBack()
			{
				var function = new VarP();
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
			public void SubtotalIsGeivenAListOfInputsFuntionNum111EmptyInputInFrontMiddleAndBack()
			{
				var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum1TestingNested()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum6TestTwoTestingNested()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum101TestingNested()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum2TestingNested()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum102TestingNested()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum3TestingNested()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum103TestingNested()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum4TestingNested()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum104TestingNested()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum5TestingNested()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum105TestingNested()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum6TestingNested()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum106TestingNested()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum7TestingNested()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum107TestingNested()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum8TestingNested()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum108TestingNested()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum9TestingNested()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum109TestingNested()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum10TestingNested()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum110TestingNested()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum11TestingNested()
		{
			var function = new VarP();
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
		public void SubtotalIsGeivenAListOfInputsFuntionNum111TestingNested()
		{
			var function = new VarP();
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
	}
}
