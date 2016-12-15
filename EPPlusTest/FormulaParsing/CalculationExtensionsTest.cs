using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
		#region Calculate ExcelWorkbook
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

		#region Calculate ExcelWorksheet
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

		#region Calculate ExcelRange
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
