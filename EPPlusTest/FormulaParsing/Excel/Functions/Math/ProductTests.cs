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
using System.Linq;
using OfficeOpenXml.FormulaParsing.Excel;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math
{
	[TestClass]
	public class ProductTests : MathFunctionsTestBase
	{
		#region Product Function (Execute) Tests
		[TestMethod]
		public void ProductWithIntegerInputReturnsCorrectValue()
		{
			var function = new Product();
			var result = function.Execute(FunctionsHelper.CreateArgs(5, 8), this.ParsingContext);
			Assert.AreEqual(40d, result.Result);
		}

		[TestMethod]
		public void ProductIsGiven5MilitaryTimes()
		{
			var function = new Product();
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
		public void ProductIsGiven5DateTimes()
		{
			var function = new Product();
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
		public void ProductWithZeroReturnsCorrectValue()
		{
			var function = new Product();
			var result = function.Execute(FunctionsHelper.CreateArgs(88, 0), this.ParsingContext);
			Assert.AreEqual(0d, result.Result);
		}

		[TestMethod]
		public void ProductWithDoublesReturnsCorrectValue()
		{
			var function = new Product();
			var result = function.Execute(FunctionsHelper.CreateArgs(5.5, 3.8), this.ParsingContext);
			Assert.AreEqual(20.9, result.Result);
		}

		[TestMethod]
		public void ProductWithTwoNegativeIntegersReturnsCorrectValue()
		{
			var function = new Product();
			var result = function.Execute(FunctionsHelper.CreateArgs(-9, -52), this.ParsingContext);
			Assert.AreEqual(468d, result.Result);
		}

		[TestMethod]
		public void ProductWithOneNegativeIntegerOnePositiveIntegerReturnsCorrectValue()
		{
			var function = new Product();
			var result = function.Execute(FunctionsHelper.CreateArgs(-15, 2), this.ParsingContext);
			Assert.AreEqual(-30d, result.Result);
		}

		[TestMethod]
		public void ProductWithFractionInputsReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["B1"].Formula = "PRODUCT((2/3),(1/5))";
				ws.Calculate();
				Assert.AreEqual(0.133333333, (double)ws.Cells["B1"].Value, 0.000001);
			}
		}

		[TestMethod]
		public void ProductWithDatesAsResultOfDateFunctionReturnsCorrectValue()
		{
			var function = new Product();
			var dateInput = new DateTime(2017, 5, 1);
			var result = function.Execute(FunctionsHelper.CreateArgs(dateInput, 1), this.ParsingContext);
			Assert.AreEqual(42856d, result.Result);
		}

		[TestMethod]
		public void ProductWithDatesAsStringsReturnsCorrectValue()
		{
			var function = new Product();
			var result = function.Execute(FunctionsHelper.CreateArgs("5/5/2017", 2), this.ParsingContext);
			Assert.AreEqual(85720d, result.Result);
		}

		[TestMethod]
		public void ProductWithDateAsStringSecondArgReturnsCorrectValue()
		{
			var function = new Product();
			var result = function.Execute(FunctionsHelper.CreateArgs(2, "5/5/2017"), this.ParsingContext);
			Assert.AreEqual(85720d, result.Result);
		}

		[TestMethod]
		public void ProductWithGeneralAndEmptyStringReturnsPoundValue()
		{
			var function = new Product();
			var result = function.Execute(FunctionsHelper.CreateArgs("string", ""), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void ProductWithNumericFisrtArgAndStringSecondArgReturnsPoundValue()
		{
			var function = new Product();
			var result = function.Execute(FunctionsHelper.CreateArgs(5, "string"), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void ProductWithStringFirstArgAndNumericSecondArgReturnsPoundValue()
		{
			var function = new Product();
			var result = function.Execute(FunctionsHelper.CreateArgs("string", 5), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void ProductWithOneArgumentReturnsCorrectValue()
		{
			var function = new Product();
			var result = function.Execute(FunctionsHelper.CreateArgs(5), this.ParsingContext);
			Assert.AreEqual(5d, result.Result);
		}

		[TestMethod]
		public void ProductWithNullSecondArgumentReturnsCorrectValue()
		{
			var function = new Product();
			var result = function.Execute(FunctionsHelper.CreateArgs(5, null), this.ParsingContext);
			Assert.AreEqual(5d, result.Result);
		}

		[TestMethod]
		public void ProductWithNullFirstArgumentReturnsCorrectValue()
		{
			var function = new Product();
			var result = function.Execute(FunctionsHelper.CreateArgs(null, 5), this.ParsingContext);
			Assert.AreEqual(5d, result.Result);
		}

		[TestMethod]
		public void ProductWithNoArgumentsReturnsPoundValue()
		{
			var function = new Product();
			var result = function.Execute(FunctionsHelper.CreateArgs(), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void ProductWithNumbersAsStringsReturnsCorrectValue()
		{
			var function = new Product();
			var result = function.Execute(FunctionsHelper.CreateArgs("5", "8"), this.ParsingContext);
			Assert.AreEqual(40d, result.Result);
		}

		[TestMethod]
		public void ProductWithOneInputAsExcelRangeReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["B1"].Value = 5;
				ws.Cells["B2"].Value = 8;
				ws.Cells["B3"].Formula = "PRODUCT(B1:B2)";
				ws.Calculate();
				Assert.AreEqual(40d, ws.Cells["B3"].Value);
			}
		}

		[TestMethod]
		public void ProductWithTwoExcelRangesAsInputsReturnsCorrectValue()
		{
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				ws.Cells["B1"].Value = 5;
				ws.Cells["B2"].Value = 8;
				ws.Cells["C1"].Value = 10;
				ws.Cells["C2"].Value = 5;
				ws.Cells["B3"].Formula = "PRODUCT(B1:B2, C1:C2)";
				ws.Calculate();
				Assert.AreEqual(2000d, ws.Cells["B3"].Value);
			}
		}

		[TestMethod]
		public void ProductWithMaxInputsReturnsCorrectValue()
		{
			// The maximum number of inputs the function takes is 264.
			using (var package = new ExcelPackage())
			{
				var ws = package.Workbook.Worksheets.Add("Sheet1");
				for(int i = 1; i < 270; i++)
				{
					for (int j = 2; j < 3; j++)
					{
						ws.Cells[i, j].Value = 2;
					}
				}
				ws.Cells["C1"].Formula = "PRODUCT(B1, B2, B3, B4, B5, B6, B7, B8, B9, B10, B11, B12, B13, B14, B15, B16, B17, B18, B19, B20, B21, " +
					"B22, B23, B24, B25, B26, B27, B28, B29, B30, B31, B32, B33, B34, B35, B36, B37, B38, B39, B40, B41, B42, B43, B44, B45, B46, " +
					"B47, B48, B49, B50, B51, B52, B53, B54, B55, B56, B57, B58, B59, B60, B61, B62, B63, B64, B65, B66, B67, B68, B69, B70, B71, " +
					"B72, B73, B74, B75, B76, B77, B78, B79, B80, B81, B82, B83, B84, B85, B86, B87, B88, B89, B90, B91, B92, B93, B94, B95, B96," +
					"B97, B98, B99, B100, B101, B102, B103, B104, B105, B106, B107, B108, B109, B110, B111, B112, B113, B114, B115, B116, B117," +
					"B118, B119, B120, B121, B122, B123, B124, B125, B126, B127, B128, B129, B130, B131, B132, B133, B134, B135, B136, B137, B138," +
					"B139, B140, B141, B142, B143, B144, B145, B146, B147, B148, B149, B150, B151, B152, B153, B154, B155, B156, B157, B158, B159," +
					"B160, B170, B171, B172, B173, B174, B175, B176, B187, B179, B180, B181, B182, B183, B184, B185, B186, B187, B188, B189, B190," +
					"B191, B192, B193, B194, B194, B195, B196, B197, B198, B199, B200, B201, B202, B203, B204, B205, B206, B207, B208, B209, B210, " +
					"B211, B212, B213, B214, B215, B216, B217, B218, B219, B220, B221, B222, B223, B224, B225, B226, B227, B228, B229, B230, B231," +
					"B232, B233, B234, B235, B236, B237, B238, B239, B240, B241, B242, B243, B244, B245, B246, B247, B248, B249, B250, B251, B252," +
					"B253, B254, B255, B256, B257, B258, B259, B260, B261, B262, B263, B264)";
				ws.Calculate();
				Assert.AreEqual(System.Math.Pow(2, 255), ws.Cells["C1"].Value);
			}
		}

		[TestMethod]
		public void ProductShouldPoundValueWhenThereAreTooFewArguments()
		{
			var func = new Product();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void ProductShouldMultiplyArguments()
		{
			var func = new Product();
			var args = FunctionsHelper.CreateArgs(2d, 2d, 4d);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(16d, result.Result);
		}

		[TestMethod]
		public void ProductShouldHandleEnumerable()
		{
			var func = new Product();
			var args = FunctionsHelper.CreateArgs(2d, 2d, FunctionsHelper.CreateArgs(4d, 2d));
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(32d, result.Result);
		}

		[TestMethod]
		public void ProductShouldIgnoreHiddenValuesIfIgnoreHiddenIsTrue()
		{
			var func = new Product();
			func.IgnoreHiddenValues = true;
			var args = FunctionsHelper.CreateArgs(2d, 2d, FunctionsHelper.CreateArgs(4d, 2d));
			args.ElementAt(1).SetExcelStateFlag(ExcelCellState.HiddenCell);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(16d, result.Result);
		}

		[TestMethod]
		public void ProductShouldHandleFirstItemIsEnumerable()
		{
			var func = new Product();
			var args = FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(4d, 2d), 2d, 2d);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(32d, result.Result);
		}

		[TestMethod]
		public void ProductFunctionWithErrorValuesAsInputReturnsTheInputErrorValue()
		{
			var func = new Product();
			var argNA = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.NA));
			var argNAME = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Name));
			var argVALUE = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Value));
			var argNUM = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Num));
			var argDIV0 = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Div0));
			var argREF = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Ref));
			var resultNA = func.Execute(argNA, this.ParsingContext);
			var resultNAME = func.Execute(argNAME, this.ParsingContext);
			var resultVALUE = func.Execute(argVALUE, this.ParsingContext);
			var resultNUM = func.Execute(argNUM, this.ParsingContext);
			var resultDIV0 = func.Execute(argDIV0, this.ParsingContext);
			var resultREF = func.Execute(argREF, this.ParsingContext);
			Assert.AreEqual(eErrorType.NA, ((ExcelErrorValue)resultNA.Result).Type);
			Assert.AreEqual(eErrorType.Name, ((ExcelErrorValue)resultNAME.Result).Type);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)resultVALUE.Result).Type);
			Assert.AreEqual(eErrorType.Num, ((ExcelErrorValue)resultNUM.Result).Type);
			Assert.AreEqual(eErrorType.Div0, ((ExcelErrorValue)resultDIV0.Result).Type);
			Assert.AreEqual(eErrorType.Ref, ((ExcelErrorValue)resultREF.Result).Type);
		}
		#endregion
	}
}
