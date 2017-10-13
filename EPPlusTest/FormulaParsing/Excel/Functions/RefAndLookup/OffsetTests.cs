using System.Collections.Generic;
using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;

namespace EPPlusTest.FormulaParsing.Excel.Functions.RefAndLookup
{
	public class OffsetTests
	{
		#region (WT)Offset Tests
		[TestMethod]
		public void OffsetReturnsPoundValueIfTooFewArgumentsAreSupplied()
		{
			var func = new Offset();
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs("B2", 2);
			this.ValidateOffsetAndWTOffset(args, parsingContext, eErrorType.Value, eErrorType.Value, true);
		}

		[TestMethod]
		public void OffsetReturnsPoundRefIfInvalidArgumentsAreSupplied()
		{
			var func = new Offset();
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs("B2", 0, 0, 0, 0);
			this.ValidateOffsetAndWTOffset(args, parsingContext, eErrorType.Ref, eErrorType.Ref, true);
		}

		[TestMethod]
		public void OffsetWithInvalidArgumentReturnsPoundValue()
		{
			var func = new Offset();
			var parsingContext = ParsingContext.Create();
			var args = FunctionsHelper.CreateArgs();
			this.ValidateOffsetAndWTOffset(args, parsingContext, eErrorType.Value, eErrorType.Value, true);
		}
		#endregion

		#region Helper Methods
		private void ValidateOffsetAndWTOffset(IEnumerable<FunctionArgument> arguments, ParsingContext context, object expectedOffsetResult, object expectedWTOffsetResult, bool errorExpected = false)
		{
			Offset offsetFunction = new Offset();
			WTOffset wTOffsetFunction = new WTOffset();
			var offsetResult = offsetFunction.Execute(arguments, context);
			var wTOffsetResult = wTOffsetFunction.Execute(arguments, context);
			if (errorExpected)
			{
				Assert.AreEqual(expectedOffsetResult, ((ExcelErrorValue)offsetResult.Result).Type);
				Assert.AreEqual(expectedWTOffsetResult, ((ExcelErrorValue)wTOffsetResult.Result).Type);
			}
			else
			{
				Assert.AreEqual(expectedOffsetResult, offsetResult);
				Assert.AreEqual(expectedWTOffsetResult, wTOffsetResult);
			}
		}
		#endregion
	}
}
