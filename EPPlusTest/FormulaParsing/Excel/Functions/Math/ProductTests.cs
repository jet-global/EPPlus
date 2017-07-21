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
	public class ProductTests
	{
		#region Properties
		private ParsingContext ParsingContext { get; set; }
		#endregion

		#region Setup / Teardown
		[TestInitialize]
		public void Initialize()
		{
			this.ParsingContext = ParsingContext.Create();
			this.ParsingContext.Scopes.NewScope(RangeAddress.Empty);
		}
		#endregion

		#region PRODUCT Function (Execute) Tests
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
			var func = new Product()
			{
				IgnoreHiddenValues = true
			};
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
		#endregion
	}
}
