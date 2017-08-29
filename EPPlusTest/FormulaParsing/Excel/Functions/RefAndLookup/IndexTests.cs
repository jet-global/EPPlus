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
using System.IO;
using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;

namespace EPPlusTest.FormulaParsing.Excel.Functions.RefAndLookup
{
	[TestClass]
	public class IndexTests
	{
		#region Properties
		private ParsingContext ParsingContext { get; set; }
		private ExcelPackage Package { get; set; }
		private ExcelWorksheet Worksheet { get; set; }
		#endregion

		#region TestInitialize/TestCleanup
		[TestInitialize]
		public void Initialize()
		{
			this.ParsingContext = ParsingContext.Create();
			this.Package = new ExcelPackage(new MemoryStream());
			this.Worksheet = this.Package.Workbook.Worksheets.Add("test");
		}

		[TestCleanup]
		public void Cleanup()
		{
			this.Package.Dispose();
		}
		#endregion

		#region Index Tests
		[TestMethod]
		public void IndexReturnsPoundValueWhenTooFewArgumentsAreSupplied()
		{
			this.Worksheet.Cells["A1"].Value = 1d;
			this.Worksheet.Cells["A2"].Value = 3d;
			this.Worksheet.Cells["A3"].Value = 5d;
			this.Worksheet.Cells["A4"].Formula = "INDEX(A1:A3,213)";
			this.Worksheet.Calculate();
			Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), this.Worksheet.Cells["A4"].Value);
		}

		[TestMethod]
		public void IndexReturnsValueByIndex()
		{
			var func = new Index();
			var result = func.Execute(FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(1, 2, 5), 3), this.ParsingContext);
			Assert.AreEqual(5, result.Result);
		}

		[TestMethod]
		public void IndexHandlesSingleRange()
		{
			this.Worksheet.Cells["A1"].Value = 1d;
			this.Worksheet.Cells["A2"].Value = 3d;
			this.Worksheet.Cells["A3"].Value = 5d;
			this.Worksheet.Cells["A4"].Formula = "INDEX(A1:A3,3)";
			this.Worksheet.Calculate();
			Assert.AreEqual(5d, this.Worksheet.Cells["A4"].Value);
		}

		[TestMethod]
		public void IndexHandlesNAError()
		{
			this.Worksheet.Cells["A1"].Value = 1d;
			this.Worksheet.Cells["A2"].Value = 3d;
			this.Worksheet.Cells["A3"].Value = 5d;
			this.Worksheet.Cells["A4"].Value = ExcelErrorValue.Create(eErrorType.NA);
			this.Worksheet.Cells["A5"].Formula = "INDEX(A1:A3,A4)";
			this.Worksheet.Calculate();
			Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), this.Worksheet.Cells["A5"].Value);
		}

		[TestMethod]
		public void IndexWithMatchInputParameters()
		{
			using (var package = new ExcelPackage())
			{
				var workSheet = package.Workbook.Worksheets.Add("Sheet1");
				workSheet.Cells["C18"].Formula = "DATE(2017, 8, 22)";
				workSheet.Cells["C19"].Formula = "DATE(2017, 8, 21)";
				workSheet.Cells["C20"].Formula = "DATE(2017, 8, 20)";
				workSheet.Cells["C21"].Formula = "DATE(2017, 8, 19)";
				workSheet.Cells["C22"].Formula = "DATE(2017, 8, 18)";
				workSheet.Cells["C23"].Formula = "DATE(2017, 8, 17)";
				workSheet.Cells["C24"].Formula = "DATE(2017, 8, 16)";
				workSheet.Cells["C25"].Formula = "DATE(2017, 8, 15)";
				workSheet.Cells["C26"].Formula = "DATE(2017, 8, 14)";
				workSheet.Cells["C27"].Formula = "DATE(2017, 8, 13)";
				workSheet.Cells["C28"].Formula = "DATE(2017, 8, 12)";
				workSheet.Cells["C29"].Formula = "DATE(2017, 8, 11)";
				workSheet.Cells["C30"].Formula = "DATE(2017, 8, 10)";
				workSheet.Cells["C31"].Formula = "DATE(2017, 8, 9)";
				workSheet.Cells["C32"].Formula = "DATE(2017, 8, 8)";
				workSheet.Cells["C33"].Formula = "DATE(2017, 8, 7)";
				workSheet.Cells["D18"].Value = 2;
				workSheet.Cells["D19"].Value = 1;
				workSheet.Cells["D20"].Value = 7;
				workSheet.Cells["D21"].Value = 6;
				workSheet.Cells["D22"].Value = 5;
				workSheet.Cells["D23"].Value = 4;
				workSheet.Cells["D24"].Value = 3;
				workSheet.Cells["D25"].Value = 2;
				workSheet.Cells["D26"].Value = 1;
				workSheet.Cells["D27"].Value = 7;
				workSheet.Cells["D28"].Value = 6;
				workSheet.Cells["D29"].Value = 5;
				workSheet.Cells["D30"].Value = 4;
				workSheet.Cells["D31"].Value = 3;
				workSheet.Cells["D32"].Value = 2;
				workSheet.Cells["D33"].Value = 1;
				workSheet.Cells["F21"].Formula = "INDEX($C$18:$C33, MATCH(7, $D$18:$D$33, 0), 0) - 6";
				workSheet.Cells["F22"].Formula = "INDEX($C$18:$C33, 3, 0) - 6";
				workSheet.Cells["E3"].Formula = "DATE(2017, 8, 14)";
				workSheet.Calculate();
				Assert.AreEqual(workSheet.Cells["E3"].Value, workSheet.Cells["F22"].Value);
				Assert.AreEqual(workSheet.Cells["E3"].Value, workSheet.Cells["F21"].Value);
			}
		}

		[TestMethod]
		public void IndexWithNoParametersReutrnsPoundValue()
		{
			var function = new Index();
			var result = function.Execute(FunctionsHelper.CreateArgs(), this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void IndexWithRegularInputReturnsCorrectValue()
		{
			this.Worksheet.Cells["B1"].Value = 5;
			this.Worksheet.Cells["B2"].Value = 6;
			this.Worksheet.Cells["B3"].Value = 4;
			this.Worksheet.Cells["B4"].Value = 9;
			this.Worksheet.Cells["C2"].Formula = "INDEX(B1:B4, 3)";
			this.Worksheet.Calculate();
			Assert.AreEqual(4, this.Worksheet.Cells["C2"].Value);
		}

		[TestMethod]
		public void IndexWithRowAsNumericStringReturnsCorrectValue()
		{
			this.Worksheet.Cells["B1"].Value = 5;
			this.Worksheet.Cells["B2"].Value = 6;
			this.Worksheet.Cells["B3"].Value = 4;
			this.Worksheet.Cells["B4"].Value = 9;
			this.Worksheet.Cells["C2"].Formula = "INDEX(B1:B4, \"2\")";
			this.Worksheet.Calculate();
			Assert.AreEqual(6, this.Worksheet.Cells["C2"].Value);
		}

		[TestMethod]
		public void IndexWithRowAsDateReturnsPoundRef()
		{
			this.Worksheet.Cells["B1"].Value = 5;
			this.Worksheet.Cells["B2"].Value = 6;
			this.Worksheet.Cells["B3"].Value = 4;
			this.Worksheet.Cells["B4"].Value = 9;
			this.Worksheet.Cells["C2"].Formula = "INDEX(B1:B4, DATE(1990, 1, 1))";
			this.Worksheet.Calculate();
			Assert.AreEqual(eErrorType.Ref, ((ExcelErrorValue)this.Worksheet.Cells["C2"].Value).Type);
		}

		[TestMethod]
		public void IndexWithRowAsGenericStringReturnsPoundValue()
		{
			this.Worksheet.Cells["B1"].Value = 5;
			this.Worksheet.Cells["B2"].Value = 6;
			this.Worksheet.Cells["B3"].Value = 4;
			this.Worksheet.Cells["B4"].Value = 9;
			this.Worksheet.Cells["C2"].Formula = "INDEX(B1:B4, \"string\")";
			this.Worksheet.Calculate();
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)this.Worksheet.Cells["C2"].Value).Type);
		}

		[TestMethod]
		public void IndexWithRowAsZeroReturnsPoundValue()
		{
			this.Worksheet.Cells["B1"].Value = 5;
			this.Worksheet.Cells["B2"].Value = 6;
			this.Worksheet.Cells["B3"].Value = 4;
			this.Worksheet.Cells["B4"].Value = 9;
			this.Worksheet.Cells["C2"].Formula = "INDEX(B1:B4, 0)";
			this.Worksheet.Calculate();
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)this.Worksheet.Cells["C2"].Value).Type);
		}

		[TestMethod]
		public void IndexWithNegativeRowReturnsPoundValue()
		{
			this.Worksheet.Cells["B1"].Value = 5;
			this.Worksheet.Cells["B2"].Value = 6;
			this.Worksheet.Cells["B3"].Value = 4;
			this.Worksheet.Cells["B4"].Value = 9;
			this.Worksheet.Cells["C2"].Formula = "INDEX(B1:B4, -2)";
			this.Worksheet.Calculate();
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)this.Worksheet.Cells["C2"].Value).Type);
		}

		[TestMethod]
		public void IndexWithRowLargeThanArraySizeReturnsPoundRef()
		{
			this.Worksheet.Cells["B1"].Value = 5;
			this.Worksheet.Cells["B2"].Value = 6;
			this.Worksheet.Cells["B3"].Value = 4;
			this.Worksheet.Cells["B4"].Value = 9;
			this.Worksheet.Cells["C2"].Formula = "INDEX(B1:B4, 34)";
			this.Worksheet.Calculate();
			Assert.AreEqual(eErrorType.Ref, ((ExcelErrorValue)this.Worksheet.Cells["C2"].Value).Type);
		}

		[TestMethod]
		public void IndexWithNullSecondAndThirdParametersReturnsPoundValue()
		{
			this.Worksheet.Cells["B1"].Value = 5;
			this.Worksheet.Cells["B2"].Value = 6;
			this.Worksheet.Cells["B3"].Value = 4;
			this.Worksheet.Cells["B4"].Value = 9;
			this.Worksheet.Cells["C2"].Formula = "INDEX(B1:B4, , )";
			this.Worksheet.Calculate();
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)this.Worksheet.Cells["C2"].Value).Type);
		}

		[TestMethod]
		public void IndexWithRowAsRefereneToCellReturnsCorrectValue()
		{
			this.Worksheet.Cells["B1"].Value = 5;
			this.Worksheet.Cells["B2"].Value = 6;
			this.Worksheet.Cells["B3"].Value = 4;
			this.Worksheet.Cells["B4"].Value = 9;
			this.Worksheet.Cells["C2"].Formula = "INDEX(B1:B4, B4)";
			this.Worksheet.Calculate();
			Assert.AreEqual(9, this.Worksheet.Cells["C2"].Value);
		}

		[TestMethod]
		public void IndexWithOneDimensionRowNumberArrayReturnsCorrectValue()
		{
			this.Worksheet.Cells["B1"].Value = 3;
			this.Worksheet.Cells["C1"].Value = 345;
			this.Worksheet.Cells["D3"].Formula = "INDEX(B1:C1, 2)";
			this.Worksheet.Calculate();
			Assert.AreEqual(345, this.Worksheet.Cells["D3"].Value);
		}

		[TestMethod]
		public void IndexWithOneDimensionColumnNumberArrayReturnsCorrectValue()
		{
			this.Worksheet.Cells["B1"].Value = 5;
			this.Worksheet.Cells["B2"].Value = 6;
			this.Worksheet.Cells["B3"].Value = 4;
			this.Worksheet.Cells["B4"].Value = 9;
			this.Worksheet.Cells["C2"].Formula = "INDEX(B1:B4, 2)";
			this.Worksheet.Calculate();
			Assert.AreEqual(6, this.Worksheet.Cells["C2"].Value);
		}

		[TestMethod]
		public void IndexWithLargeNumberArrayWithNullThirdParameterReturnsPoundValue()
		{
			this.Worksheet.Cells["B1"].Value = 5;
			this.Worksheet.Cells["B2"].Value = 6;
			this.Worksheet.Cells["C1"].Value = 3;
			this.Worksheet.Cells["C2"].Value = 8;
			this.Worksheet.Cells["D5"].Formula = "INDEX(B1:C2, 2, )";
			this.Worksheet.Calculate();
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)this.Worksheet.Cells["D5"].Value).Type);
		}

		[TestMethod]
		public void IndexWithLargeNumberArrayWithNoThirdParameterReturnsPoundRef()
		{
			this.Worksheet.Cells["B1"].Value = 5;
			this.Worksheet.Cells["B2"].Value = 6;
			this.Worksheet.Cells["C1"].Value = 3;
			this.Worksheet.Cells["C2"].Value = 8;
			this.Worksheet.Cells["D5"].Formula = "INDEX(B1:C2, 2)";
			this.Worksheet.Calculate();
			Assert.AreEqual(eErrorType.Ref, ((ExcelErrorValue)this.Worksheet.Cells["D5"].Value).Type);
		}

		[TestMethod]
		public void IndexWithOneDimensionRowStringArrayReturnsCorrectValue()
		{
			this.Worksheet.Cells["B1"].Value = "string";
			this.Worksheet.Cells["C1"].Value = "string";
			this.Worksheet.Cells["D1"].Value = "string";
			this.Worksheet.Cells["C2"].Formula = "INDEX(B1:B3, 2)";
			this.Worksheet.Calculate();
			Assert.AreEqual("string", this.Worksheet.Cells["C2"].Value);
		}

		[TestMethod]
		public void IndexWithOneDimensionColumnStringArrayReturnsCorrectValue()
		{
			this.Worksheet.Cells["B1"].Value = "string";
			this.Worksheet.Cells["B2"].Value = "string";
			this.Worksheet.Cells["B3"].Value = "string";
			this.Worksheet.Cells["C2"].Formula = "INDEX(B1:B3, 2)";
			this.Worksheet.Calculate();
			Assert.AreEqual("string", this.Worksheet.Cells["C2"].Value);
		}

		[TestMethod]
		public void IndexWithLargeStringArrayWithNullThirdParameterReturnsPoundValue()
		{
			this.Worksheet.Cells["B1"].Value = "string";
			this.Worksheet.Cells["B2"].Value = "string";
			this.Worksheet.Cells["C1"].Value = "string";
			this.Worksheet.Cells["C2"].Value = "string";
			this.Worksheet.Cells["D5"].Formula = "INDEX(B1:C2, 2, )";
			this.Worksheet.Calculate();
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)this.Worksheet.Cells["D5"].Value).Type);
		}

		[TestMethod]
		public void IndexWithLargeStringArrayWithNoThirdParameterReturnsPoundRef()
		{
			this.Worksheet.Cells["B1"].Value = "string";
			this.Worksheet.Cells["B2"].Value = "string";
			this.Worksheet.Cells["C1"].Value = "string";
			this.Worksheet.Cells["C2"].Value = "string";
			this.Worksheet.Cells["D5"].Formula = "INDEX(B1:C2, 2)";
			this.Worksheet.Calculate();
			Assert.AreEqual(eErrorType.Ref, ((ExcelErrorValue)this.Worksheet.Cells["D5"].Value).Type);
		}
		#endregion
	}
}
