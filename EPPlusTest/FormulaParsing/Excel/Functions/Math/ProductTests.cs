using System;
using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml;

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
			var function = new Product();
			var result = function.Execute(FunctionsHelper.CreateArgs((2 / 3), (1 / 5)), this.ParsingContext);
			Assert.AreEqual(0.13333333, result.Result);
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
		public void ProductWithDateNotAsStringReturnsCorrectValue()
		{
			var function = new Product();
			var result = function.Execute(FunctionsHelper.CreateArgs(5 / 5 / 2017, 2), this.ParsingContext);
			Assert.AreEqual(0.000991572, result.Result);
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
				ws.Cells["B1"].Value = 2;
				ws.Cells["B2"].Value = 2;
				ws.Cells["B3"].Value = 2;
				ws.Cells["B4"].Value = 2;
				ws.Cells["B5"].Value = 2;
				ws.Cells["B6"].Value = 2;
				ws.Cells["B7"].Value = 2;
				ws.Cells["B8"].Value = 2;
				ws.Cells["B9"].Value = 2;
				ws.Cells["B10"].Value = 2;
				ws.Cells["B11"].Value = 2;
				ws.Cells["B12"].Value = 2;
				ws.Cells["B13"].Value = 2;
				ws.Cells["B14"].Value = 2;
				ws.Cells["B15"].Value = 2;
				ws.Cells["B16"].Value = 2;
				ws.Cells["B17"].Value = 2;
				ws.Cells["B18"].Value = 2;
				ws.Cells["B19"].Value = 2;
				ws.Cells["B20"].Value = 2;
				ws.Cells["B21"].Value = 2;
				ws.Cells["B22"].Value = 2;
				ws.Cells["B23"].Value = 2;
				ws.Cells["B24"].Value = 2;
				ws.Cells["B25"].Value = 2;
				ws.Cells["B26"].Value = 2;
				ws.Cells["B27"].Value = 2;
				ws.Cells["B28"].Value = 2;
				ws.Cells["B29"].Value = 2;
				ws.Cells["B30"].Value = 2;
				ws.Cells["B31"].Value = 2;
				ws.Cells["B32"].Value = 2;
				ws.Cells["B33"].Value = 2;
				ws.Cells["B34"].Value = 2;
				ws.Cells["B35"].Value = 2;
				ws.Cells["B36"].Value = 2;
				ws.Cells["B37"].Value = 2;
				ws.Cells["B38"].Value = 2;
				ws.Cells["B39"].Value = 2;
				ws.Cells["B40"].Value = 2;
				ws.Cells["B41"].Value = 2;
				ws.Cells["B42"].Value = 2;
				ws.Cells["B43"].Value = 2;
				ws.Cells["B44"].Value = 2;
				ws.Cells["B45"].Value = 2;
				ws.Cells["B46"].Value = 2;
				ws.Cells["B47"].Value = 2;
				ws.Cells["B48"].Value = 2;
				ws.Cells["B49"].Value = 2;
				ws.Cells["B50"].Value = 2;
				ws.Cells["B51"].Value = 2;
				ws.Cells["B52"].Value = 2;
				ws.Cells["B53"].Value = 2;
				ws.Cells["B54"].Value = 2;
				ws.Cells["B55"].Value = 2;
				ws.Cells["B56"].Value = 2;
				ws.Cells["B57"].Value = 2;
				ws.Cells["B58"].Value = 2;
				ws.Cells["B59"].Value = 2;
				ws.Cells["B60"].Value = 2;
				ws.Cells["B61"].Value = 2;
				ws.Cells["B62"].Value = 2;
				ws.Cells["B63"].Value = 2;
				ws.Cells["B64"].Value = 2;
				ws.Cells["B65"].Value = 2;
				ws.Cells["B66"].Value = 2;
				ws.Cells["B67"].Value = 2;
				ws.Cells["B68"].Value = 2;
				ws.Cells["B69"].Value = 2;
				ws.Cells["B70"].Value = 2;
				ws.Cells["B71"].Value = 2;
				ws.Cells["B72"].Value = 2;
				ws.Cells["B73"].Value = 2;
				ws.Cells["B74"].Value = 2;
				ws.Cells["B75"].Value = 2;
				ws.Cells["B76"].Value = 2;
				ws.Cells["B77"].Value = 2;
				ws.Cells["B78"].Value = 2;
				ws.Cells["B79"].Value = 2;
				ws.Cells["B80"].Value = 2;
				ws.Cells["B81"].Value = 2;
				ws.Cells["B82"].Value = 2;
				ws.Cells["B83"].Value = 2;
				ws.Cells["B84"].Value = 2;
				ws.Cells["B85"].Value = 2;
				ws.Cells["B86"].Value = 2;
				ws.Cells["B87"].Value = 2;
				ws.Cells["B88"].Value = 2;
				ws.Cells["B89"].Value = 2;
				ws.Cells["B90"].Value = 2;
				ws.Cells["B91"].Value = 2;
				ws.Cells["B92"].Value = 2;
				ws.Cells["B93"].Value = 2;
				ws.Cells["B94"].Value = 2;
				ws.Cells["B95"].Value = 2;
				ws.Cells["B96"].Value = 2;
				ws.Cells["B97"].Value = 2;
				ws.Cells["B98"].Value = 2;
				ws.Cells["B99"].Value = 2;
				ws.Cells["B100"].Value = 2;
				ws.Cells["B101"].Value = 2;
				ws.Cells["B102"].Value = 2;
				ws.Cells["B103"].Value = 2;
				ws.Cells["B104"].Value = 2;
				ws.Cells["B105"].Value = 2;
				ws.Cells["B106"].Value = 2;
				ws.Cells["B107"].Value = 2;
				ws.Cells["B108"].Value = 2;
				ws.Cells["B109"].Value = 2;
				ws.Cells["B110"].Value = 2;
				ws.Cells["B111"].Value = 2;
				ws.Cells["B112"].Value = 2;
				ws.Cells["B113"].Value = 2;
				ws.Cells["B114"].Value = 2;
				ws.Cells["B115"].Value = 2;
				ws.Cells["B116"].Value = 2;
				ws.Cells["B117"].Value = 2;
				ws.Cells["B118"].Value = 2;
				ws.Cells["B119"].Value = 2;
				ws.Cells["B120"].Value = 2;
				ws.Cells["B121"].Value = 2;
				ws.Cells["B122"].Value = 2;
				ws.Cells["B123"].Value = 2;
				ws.Cells["B124"].Value = 2;
				ws.Cells["B125"].Value = 2;
				ws.Cells["B126"].Value = 2;
				ws.Cells["B127"].Value = 2;
				ws.Cells["B128"].Value = 2;
				ws.Cells["B129"].Value = 2;
				ws.Cells["B130"].Value = 2;
				ws.Cells["B131"].Value = 2;
				ws.Cells["B132"].Value = 2;
				ws.Cells["B133"].Value = 2;
				ws.Cells["B134"].Value = 2;
				ws.Cells["B135"].Value = 2;
				ws.Cells["B136"].Value = 2;
				ws.Cells["B137"].Value = 2;
				ws.Cells["B138"].Value = 2;
				ws.Cells["B139"].Value = 2;
				ws.Cells["B140"].Value = 2;
				ws.Cells["B141"].Value = 2;
				ws.Cells["B142"].Value = 2;
				ws.Cells["B143"].Value = 2;
				ws.Cells["B144"].Value = 2;
				ws.Cells["B145"].Value = 2;
				ws.Cells["B146"].Value = 2;
				ws.Cells["B147"].Value = 2;
				ws.Cells["B148"].Value = 2;
				ws.Cells["B149"].Value = 2;
				ws.Cells["B150"].Value = 2;
				ws.Cells["B151"].Value = 2;
				ws.Cells["B152"].Value = 2;
				ws.Cells["B153"].Value = 2;
				ws.Cells["B154"].Value = 2;
				ws.Cells["B155"].Value = 2;
				ws.Cells["B156"].Value = 2;
				ws.Cells["B157"].Value = 2;
				ws.Cells["B158"].Value = 2;
				ws.Cells["B159"].Value = 2;
				ws.Cells["B160"].Value = 2;
				ws.Cells["B161"].Value = 2;
				ws.Cells["B162"].Value = 2;
				ws.Cells["B163"].Value = 2;
				ws.Cells["B164"].Value = 2;
				ws.Cells["B165"].Value = 2;
				ws.Cells["B166"].Value = 2;
				ws.Cells["B167"].Value = 2;
				ws.Cells["B168"].Value = 2;
				ws.Cells["B169"].Value = 2;
				ws.Cells["B170"].Value = 2;
				ws.Cells["B171"].Value = 2;
				ws.Cells["B172"].Value = 2;
				ws.Cells["B173"].Value = 2;
				ws.Cells["B174"].Value = 2;
				ws.Cells["B175"].Value = 2;
				ws.Cells["B176"].Value = 2;
				ws.Cells["B177"].Value = 2;
				ws.Cells["B178"].Value = 2;
				ws.Cells["B179"].Value = 2;
				ws.Cells["B180"].Value = 2;
				ws.Cells["B181"].Value = 2;
				ws.Cells["B182"].Value = 2;
				ws.Cells["B183"].Value = 2;
				ws.Cells["B184"].Value = 2;
				ws.Cells["B185"].Value = 2;
				ws.Cells["B186"].Value = 2;
				ws.Cells["B187"].Value = 2;
				ws.Cells["B188"].Value = 2;
				ws.Cells["B189"].Value = 2;
				ws.Cells["B190"].Value = 2;
				ws.Cells["B191"].Value = 2;
				ws.Cells["B192"].Value = 2;
				ws.Cells["B193"].Value = 2;
				ws.Cells["B194"].Value = 2;
				ws.Cells["B195"].Value = 2;
				ws.Cells["B196"].Value = 2;
				ws.Cells["B197"].Value = 2;
				ws.Cells["B198"].Value = 2;
				ws.Cells["B199"].Value = 2;
				ws.Cells["B200"].Value = 2;
				ws.Cells["B201"].Value = 2;
				ws.Cells["B202"].Value = 2;
				ws.Cells["B203"].Value = 2;
				ws.Cells["B204"].Value = 2;
				ws.Cells["B205"].Value = 2;
				ws.Cells["B206"].Value = 2;
				ws.Cells["B207"].Value = 2;
				ws.Cells["B208"].Value = 2;
				ws.Cells["B209"].Value = 2;
				ws.Cells["B210"].Value = 2;
				ws.Cells["B211"].Value = 2;
				ws.Cells["B212"].Value = 2;
				ws.Cells["B213"].Value = 2;
				ws.Cells["B214"].Value = 2;
				ws.Cells["B215"].Value = 2;
				ws.Cells["B216"].Value = 2;
				ws.Cells["B217"].Value = 2;
				ws.Cells["B218"].Value = 2;
				ws.Cells["B219"].Value = 2;
				ws.Cells["B220"].Value = 2;
				ws.Cells["B221"].Value = 2;
				ws.Cells["B222"].Value = 2;
				ws.Cells["B223"].Value = 2;
				ws.Cells["B224"].Value = 2;
				ws.Cells["B225"].Value = 2;
				ws.Cells["B226"].Value = 2;
				ws.Cells["B227"].Value = 2;
				ws.Cells["B228"].Value = 2;
				ws.Cells["B229"].Value = 2;
				ws.Cells["B230"].Value = 2;
				ws.Cells["B231"].Value = 2;
				ws.Cells["B232"].Value = 2;
				ws.Cells["B233"].Value = 2;
				ws.Cells["B234"].Value = 2;
				ws.Cells["B235"].Value = 2;
				ws.Cells["B236"].Value = 2;
				ws.Cells["B237"].Value = 2;
				ws.Cells["B238"].Value = 2;
				ws.Cells["B239"].Value = 2;
				ws.Cells["B240"].Value = 2;
				ws.Cells["B241"].Value = 2;
				ws.Cells["B242"].Value = 2;
				ws.Cells["B243"].Value = 2;
				ws.Cells["B244"].Value = 2;
				ws.Cells["B245"].Value = 2;
				ws.Cells["B246"].Value = 2;
				ws.Cells["B247"].Value = 2;
				ws.Cells["B248"].Value = 2;
				ws.Cells["B249"].Value = 2;
				ws.Cells["B250"].Value = 2;
				ws.Cells["B251"].Value = 2;
				ws.Cells["B252"].Value = 2;
				ws.Cells["B253"].Value = 2;
				ws.Cells["B254"].Value = 2;
				ws.Cells["B255"].Value = 2;
				ws.Cells["B256"].Value = 2;
				ws.Cells["B257"].Value = 2;
				ws.Cells["B258"].Value = 2;
				ws.Cells["B259"].Value = 2;
				ws.Cells["B260"].Value = 2;
				ws.Cells["B261"].Value = 2;
				ws.Cells["B262"].Value = 2;
				ws.Cells["B263"].Value = 2;
				ws.Cells["B264"].Value = 2;
				ws.Cells["C1"].Formula = "PRODUCT(B1, B2, B3, B4, B5, B6, B7, B8, B9, B10, B11, B12, B13, B14, B15, B16, B17, B18, B19, B20, B21, B22, B23, B24, B25, B26, B27, B28, B29, B30, B31, B32, B33, B34, B35, B36, B37, B38, B39, B40, B41, B42, B43, B44, B45, B46, B47, B48, B49, B50, B51, B52, B53, B54, B55, B56, B57, B58, B59, B60, B61, B62, B63, B64, B65, B66, B67, B68, B69, B70, B71, B72, B73, B74, B75, B76, B77, B78, B79, B80, B81, B82, B83, B84, B85, B86, B87, B88, B89, B90, B91, B92, B93, B94, B95, B96, B97, B98, B99, B100, B101, B102, B103, B104, B105, B106, B107, B108, B109, B110, B111, B112, B113, B114, B115, B116, B117, B118, B119, B120, B121, B122, B123, B124, B125, B126, B127, B128, B129, B130, B131, B132, B133, B134, B135, B136, B137, B138, B139, B140, B141, B142, B143, B144, B145, B146, B147, B148, B149, B150, B151, B152, B153, B154, B155, B156, B157, B158, B159, B160, B170, B171, B172, B173, B174, B175, B176, B187, B179, B180, B181, B182, B183, B184, B185, B186, B187, B188, B189, B190, B191, B192, B193, B194, B194, B195, B196, B197, B198, B199, B200, B201, B202, B203, B204, B205, B206, B207, B208, B209, B210, B211, B212, B213, B214, B215, B216, B217, B218, B219, B220, B221, B222, B223, B224, B225, B226, B227, B228, B229, B230, B231, B232, B233, B234, B235, B236, B237, B238, B239, B240, B241, B242, B243, B244, B245, B246, B247, B248, B249, B250, B251, B252, B253, B254, B255, B256, B257, B258, B259, B260, B261, B262, B263, B264)";
				ws.Calculate();
				Assert.AreEqual(System.Math.Pow(2, 255), ws.Cells["C1"].Value);
			}
		}
		#endregion
	}
}
