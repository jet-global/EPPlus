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

		#region Integration Offset Tests
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
				worksheet.Cells[2, 3].Value = "NOTAFUNCTION(1,2,3)";
				worksheet.Cells[3, 3].Formula = formula;
				worksheet.Cells[2, 2].Calculate();
				var result = worksheet.Cells[2, 2].Value;
				Assert.IsInstanceOfType(result, typeof(ExcelErrorValue));
				Assert.AreEqual(ExcelErrorValue.Values.Value, result.ToString());
				Assert.AreEqual(formula, worksheet.Cells[3, 3].Formula);
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
