using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math
{
	[TestClass]
	public class DegreesTests : MathFunctionsTestBase
	{
		[TestMethod]
		public void DegreesFunctionWithTooFewArgumentsReturnsPoundValue()
		{
			var func = new Degrees();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(eErrorType.Value, ((ExcelErrorValue)result.Result).Type);
		}

		[TestMethod]
		public void DegreesFunctionWithInputZeroRadiansReturnsCorrectResult()
		{
			var func = new Degrees();
			var args = FunctionsHelper.CreateArgs(0);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(0.0, result.Result);
		}

		[TestMethod]
		public void DegreesFunctionWithInputPiOver2RadiansReturnsCorrectResult()
		{
			var func = new Degrees();
			var pi = System.Math.PI;
			var args = FunctionsHelper.CreateArgs(pi/2);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(90.0, result.Result);
		}

		[TestMethod]
		public void DegreesFunctionWithInput2TimesPiRadiansReturnsCorrectResult()
		{
			var func = new Degrees();
			var pi = System.Math.PI;
			var args = FunctionsHelper.CreateArgs(2*pi);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(360, result.Result);
		}

		[TestMethod]
		public void DegreesFunctionWithNegativeIntegerReturnsCorrectResult()
		{
			var func = new Degrees();
			var args = FunctionsHelper.CreateArgs(-1);
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(, result.Result);
		}

		[TestMethod]
		public void DegreesFunction()
		{
			var func = new Degrees();
			var args = FunctionsHelper.CreateArgs();
			var result = func.Execute(args, this.ParsingContext);
			Assert.AreEqual(, result.Result);
		}
	}
}
