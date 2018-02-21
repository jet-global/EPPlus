using System.Collections.Generic;
using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;

namespace EPPlusTest.FormulaParsing.Excel.Functions.RefAndLookup
{
	[TestClass]
	public class OffsetTests
	{
		#region Offset/OffsetAddress Tests
		[TestMethod]
		public void OffsetReturnsPoundValueIfTooFewArgumentsAreSupplied()
		{
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs("B2", 2);
			this.ValidateOffsetAndOffsetAddress(args, parsingContext, eErrorType.Value, eErrorType.Value, true);
		}

		[TestMethod]
		public void OffsetReturnsPoundRefIfInvalidArgumentsAreSupplied()
		{
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs("B2", 0, 0, 0, 0);
			this.ValidateOffsetAndOffsetAddress(args, parsingContext, eErrorType.Ref, eErrorType.Ref, true);
		}

		[TestMethod]
		public void OffsetWithInvalidArgumentReturnsPoundValue()
		{
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs();
			this.ValidateOffsetAndOffsetAddress(args, parsingContext, eErrorType.Value, eErrorType.Value, true);
		}
		#endregion

		#region Offset Error Handling Tests
		[TestMethod]
		public void OffsetWithInvalidAddressDependencyChainCompletesCalculation()
		{
			const string formula = "OFFSET('not a sheet'!G6, 5, 4)";
			using (var excelPackage = new ExcelPackage())
			{
				var worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells[2, 2].Formula = "C3";
				worksheet.Cells[3, 3].Formula = formula;
				worksheet.Cells[2, 2].Calculate();
				var result = worksheet.Cells[2, 2].Value;
				Assert.IsInstanceOfType(result, typeof(ExcelErrorValue));
				Assert.AreEqual(ExcelErrorValue.Values.Value, result.ToString());
				Assert.AreEqual(formula, worksheet.Cells[3, 3].Formula);
			}
		}

		[TestMethod]
		public void OffsetWithArgumentErrorDependencyChainCompletesCalculation()
		{
			const string formula = "OFFSET('Sheet1'!G6, 5, C2)";
			using (var excelPackage = new ExcelPackage())
			{
				var worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells[2, 2].Formula = "C3";
				worksheet.Cells[2, 3].Formula = "NOTAFUNCTION(1,2,3)";
				worksheet.Cells[3, 3].Formula = formula;
				worksheet.Cells[2, 2].Calculate();
				var result = worksheet.Cells[2, 2].Value;
				Assert.IsInstanceOfType(result, typeof(ExcelErrorValue));
				Assert.AreEqual(ExcelErrorValue.Values.Name, result.ToString());
				Assert.AreEqual(formula, worksheet.Cells[3, 3].Formula);
			}
		}

		[TestMethod]
		public void OffsetWithNAErrorDependencyChainCompletesCalculation()
		{
			const string formula = "OFFSET('Sheet1'!G6, 5, C2)";
			using (var excelPackage = new ExcelPackage())
			{
				var worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");
				worksheet.Cells[2, 2].Formula = "C3";
				worksheet.Cells[2, 3].Formula = "MATCH(2,F4:F9,0)"; // Should return #NA.
				worksheet.Cells[3, 3].Formula = formula;
				worksheet.Calculate();
				var result = worksheet.Cells[2, 2].Value;
				Assert.IsInstanceOfType(result, typeof(ExcelErrorValue));
				Assert.AreEqual(ExcelErrorValue.Values.NA, result.ToString());
				Assert.AreEqual(formula, worksheet.Cells[3, 3].Formula);
				result = worksheet.Cells[2, 3].Value;
				Assert.IsInstanceOfType(result, typeof(ExcelErrorValue));
				Assert.AreEqual(ExcelErrorValue.Values.NA, result.ToString());
			}
		}
		#endregion

		#region Offset Integration Tests
		[TestMethod]
		public void OffsetWithHeight()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet1");
				sheet.Cells[2, 2].Formula = "SUM(OFFSET(C3, 0, 0, 5))";
				sheet.Cells[3, 3].Value = 1;
				sheet.Cells[4, 3].Value = 2;
				sheet.Cells[5, 3].Value = 3;
				sheet.Cells[6, 3].Value = 4;
				sheet.Cells[7, 3].Value = 5;
				sheet.Calculate();
				Assert.AreEqual(15d, sheet.Cells[2, 2].Value);
			}
		}

		[TestMethod]
		public void OffsetWithWidth()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet1");
				sheet.Cells[2, 2].Formula = "SUM(OFFSET(C3, 0, 0, 1, 5))";
				sheet.Cells[3, 3].Value = 1;
				sheet.Cells[3, 4].Value = 2;
				sheet.Cells[3, 5].Value = 3;
				sheet.Cells[3, 6].Value = 4;
				sheet.Cells[3, 7].Value = 5;
				sheet.Calculate();
				Assert.AreEqual(15d, sheet.Cells[2, 2].Value);
			}
		}
		#endregion

		#region Helper Methods
		private void ValidateOffsetAndOffsetAddress(IEnumerable<FunctionArgument> arguments, ParsingContext context, object expectedOffsetResult, object expectedOffsetAddressResult, bool errorExpected = false)
		{
			Offset offsetFunction = new Offset();
			OffsetAddress offsetAddressFunction = new OffsetAddress();
			var offsetResult = offsetFunction.Execute(arguments, context);
			var offsetAddressResult = offsetAddressFunction.Execute(arguments, context);
			if (errorExpected)
			{
				Assert.AreEqual(expectedOffsetResult, ((ExcelErrorValue)offsetResult.Result).Type);
				Assert.AreEqual(expectedOffsetAddressResult, ((ExcelErrorValue)offsetAddressResult.Result).Type);
			}
			else
			{
				Assert.AreEqual(expectedOffsetResult, offsetResult);
				Assert.AreEqual(expectedOffsetAddressResult, offsetAddressResult);
			}
		}
		#endregion
	}
}
