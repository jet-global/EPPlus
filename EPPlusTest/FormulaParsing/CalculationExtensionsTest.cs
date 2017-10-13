using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace EPPlusTest.FormulaParsing
{
	[TestClass]
	public class CalculationExtensionsTest
	{
		#region Calculate ExcelWorkbook Tests
		[TestMethod]
		[ExpectedException(typeof(OperationCanceledException))]
		public void CalculateExcelWorkbookAllowsOperationCanceledExceptionsThrough()
		{
			using (var package = new ExcelPackage())
			using (var worksheet = package.Workbook.Worksheets.Add("Sheet"))
			{
				package.Workbook.FormulaParserManager.AddOrReplaceFunction("Func", new CanceledFunction());
				worksheet.Cells[3, 3].Formula = "Func()";
				package.Workbook.Calculate();
			}
		}
		#endregion

		#region Calculate ExcelWorksheet Tests
		[TestMethod]
		[ExpectedException(typeof(OperationCanceledException))]
		public void CalculateExcelWorksheetAllowsOperationCanceledExceptionsThrough()
		{
			using (var package = new ExcelPackage())
			using (var worksheet = package.Workbook.Worksheets.Add("Sheet"))
			{
				package.Workbook.FormulaParserManager.AddOrReplaceFunction("Func", new CanceledFunction());
				worksheet.Cells[3, 3].Formula = "Func()";
				worksheet.Calculate();
			}
		}
		#endregion

		#region Calculate ExcelRange Tests
		[TestMethod]
		[ExpectedException(typeof(OperationCanceledException))]
		public void CalculateExcelRangeAllowsOperationCanceledExceptionsThrough()
		{
			using (var package = new ExcelPackage())
			using (var worksheet = package.Workbook.Worksheets.Add("Sheet"))
			{
				package.Workbook.FormulaParserManager.AddOrReplaceFunction("Func", new CanceledFunction());
				worksheet.Cells[3, 3].Formula = "Func()";
				worksheet.Cells[3, 3].Calculate();
			}
		}
		#endregion

		#region Calculate Function String Tests
		[TestMethod]
		public void CalculateFunctionString()
		{
			using (var package = new ExcelPackage())
			using (var worksheet = package.Workbook.Worksheets.Add("Sheet"))
			{
				worksheet.Cells["E5"].Formula = "=ROW()";
				Assert.AreEqual(null, worksheet.Cells["E5"].Value);
				var result = worksheet.Calculate("E5 * 5");
				Assert.AreEqual(5, worksheet.Cells["E5"].Value);
				Assert.AreEqual(25d, result);
			}
		}

		[TestMethod]
		public void CalculateFunctionStringDefaultAddress()
		{
			using (var package = new ExcelPackage())
			using (var worksheet = package.Workbook.Worksheets.Add("Sheet"))
			{
				var result = worksheet.Calculate("ROW()");
				Assert.AreEqual(-1, result);
				result = worksheet.Calculate("COLUMN()");
				Assert.AreEqual(-1, result);
			}
		}

		[TestMethod]
		public void CalculateFunctionStringSpecifiedAddress()
		{
			using (var package = new ExcelPackage())
			using (var worksheet = package.Workbook.Worksheets.Add("Sheet"))
			{
				var result = worksheet.Calculate("ROW()", 5, 6);
				Assert.AreEqual(5, result);
				result = worksheet.Calculate("COLUMN()", 5, 6);
				Assert.AreEqual(6, result);
			}
		}
		#endregion

		#region Nested Classes
		public class CanceledFunction : ExcelFunction
		{
			public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
			{
				throw new OperationCanceledException();
			}
		}
		#endregion
	}
}
