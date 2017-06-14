using System;
using System.Linq;
using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math
{
	[TestClass]
	public class AverageATests : MathFunctionsTestBase
	{
		[TestMethod]
		public void AverageALiterals()
		{
			// For literals, AverageA always parses and include numeric strings, date strings, bools, etc.
			// The only exception is unparsable string literals, which cause a #VALUE.
			AverageA average = new AverageA();
			var date1 = new DateTime(2013, 1, 5);
			var date2 = new DateTime(2013, 1, 15);
			double value1 = 1000;
			double value2 = 2000;
			double value3 = 6000;
			double value4 = 1;
			double value5 = date1.ToOADate();
			double value6 = date2.ToOADate();
			var result = average.Execute(new FunctionArgument[]
			{
				new FunctionArgument(value1.ToString("n")),
				new FunctionArgument(value2),
				new FunctionArgument(value3.ToString("n")),
				new FunctionArgument(true),
				new FunctionArgument(date1),
				new FunctionArgument(date2.ToString("d"))
			}, ParsingContext.Create());
			Assert.AreEqual((value1 + value2 + value3 + value4 + value5 + value6) / 6, result.Result);
		}

		[TestMethod]
		public void AverageACellReferences()
		{
			// For cell references, AverageA divides by all cells, but only adds actual numbers, dates, and booleans.
			ExcelPackage package = new ExcelPackage();
			var worksheet = package.Workbook.Worksheets.Add("Test");
			double[] values =
			{
				0,
				2000,
				0,
				1,
				new DateTime(2013, 1, 5).ToOADate(),
				0
			};
			ExcelRange range1 = worksheet.Cells[1, 1];
			range1.Formula = "\"1000\"";
			range1.Calculate();
			var range2 = worksheet.Cells[1, 2];
			range2.Value = 2000;
			var range3 = worksheet.Cells[1, 3];
			range3.Formula = $"\"{new DateTime(2013, 1, 5).ToString("d")}\"";
			range3.Calculate();
			var range4 = worksheet.Cells[1, 4];
			range4.Value = true;
			var range5 = worksheet.Cells[1, 5];
			range5.Value = new DateTime(2013, 1, 5);
			var range6 = worksheet.Cells[1, 6];
			range6.Value = "Test";
			AverageA average = new AverageA();
			var rangeInfo1 = new EpplusExcelDataProvider.RangeInfo(worksheet, 1, 1, 1, 3);
			var rangeInfo2 = new EpplusExcelDataProvider.RangeInfo(worksheet, 1, 4, 1, 4);
			var rangeInfo3 = new EpplusExcelDataProvider.RangeInfo(worksheet, 1, 5, 1, 6);
			var context = ParsingContext.Create();
			var address = new OfficeOpenXml.FormulaParsing.ExcelUtilities.RangeAddress();
			address.FromRow = address.ToRow = address.FromCol = address.ToCol = 2;
			context.Scopes.NewScope(address);
			var result = average.Execute(new FunctionArgument[]
			{
				new FunctionArgument(rangeInfo1),
				new FunctionArgument(rangeInfo2),
				new FunctionArgument(rangeInfo3)
			}, context);
			Assert.AreEqual(values.Average(), result.Result);
		}

		[TestMethod]
		public void AverageAArray()
		{
			// For arrays, AverageA completely ignores booleans.  It divides by strings and numbers, but only
			// numbers are added to the total.  Real dates cannot be specified and string dates are not parsed.
			AverageA average = new AverageA();
			var date = new DateTime(2013, 1, 15);
			double[] values =
			{
				0,
				2000,
				0,
				0,
				0
			};
			var result = average.Execute(new FunctionArgument[]
			{
				new FunctionArgument(new FunctionArgument[]
				{
					new FunctionArgument(1000.ToString("n")),
					new FunctionArgument(2000),
					new FunctionArgument(6000.ToString("n")),
					new FunctionArgument(true),
					new FunctionArgument(date.ToString("d")),
					new FunctionArgument("test")
				})
			}, ParsingContext.Create());
			Assert.AreEqual(values.Average(), result.Result);
		}

		[TestMethod]
		public void AverageAUnparsableLiteral()
		{
			// In the case of literals, any unparsable string literal results in a #VALUE.
			AverageA average = new AverageA();
			var result = average.Execute(new FunctionArgument[]
			{
				new FunctionArgument(1000),
				new FunctionArgument("Test")
			}, ParsingContext.Create());
			Assert.AreEqual(OfficeOpenXml.FormulaParsing.ExpressionGraph.DataType.ExcelError, result.DataType);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)(result.Result)).Type);
		}

		[TestMethod]
		public void AverageAFunctionWithErrorValuesAsInputReturnsTheInputErrorValue()
		{
			var func = new AverageA();
			var argNA = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.NA),1,1,1,1);
			var argNAME = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Name),1,1,1,1);
			var argVALUE = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Value),1,1,1,1);
			var argNUM = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Num),1,1,1,1);
			var argDIV0 = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Div0),1,1,1,1);
			var argREF = FunctionsHelper.CreateArgs(ExcelErrorValue.Create(eErrorType.Ref),1,1,1,1);
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
	}
}
