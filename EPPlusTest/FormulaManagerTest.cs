using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest
{
	[TestClass]
	public class FormulaManagerTest
	{
		#region UpdateFormulaReferences Tests
		[TestMethod]
		public void UpdateFormulaReferencesOnTheSameSheet()
		{
			var formulaManager = new FormulaManager();
			var result = formulaManager.UpdateFormulaReferences("C3", 3, 3, 2, 2, "sheet", "sheet");
			Assert.AreEqual("F6", result);
		}

		[TestMethod]
		public void UpdateFormulaReferencesIgnoresIncorrectSheet()
		{
			var formulaManager = new FormulaManager();
			var result = formulaManager.UpdateFormulaReferences("C3", 3, 3, 2, 2, "sheet", "other sheet");
			Assert.AreEqual("C3", result);
		}

		[TestMethod]
		public void UpdateFormulaReferencesFullyQualifiedReferenceOnTheSameSheet()
		{
			var formulaManager = new FormulaManager();
			var result = formulaManager.UpdateFormulaReferences("'sheet name here'!C3", 3, 3, 2, 2, "sheet name here", "sheet name here");
			Assert.AreEqual("'sheet name here'!F6", result);
		}

		[TestMethod]
		public void UpdateFormulaReferencesFullyQualifiedCrossSheetReferenceArray()
		{
			var formulaManager = new FormulaManager();
			var result = formulaManager.UpdateFormulaReferences("SUM('sheet name here'!B2:D4)", 3, 3, 3, 3, "cross sheet", "sheet name here");
			Assert.AreEqual("SUM('sheet name here'!B2:G7)", result);
		}

		[TestMethod]
		public void UpdateFormulaReferencesFullyQualifiedReferenceOnADifferentSheet()
		{
			var formulaManager = new FormulaManager();
			var result = formulaManager.UpdateFormulaReferences("'updated sheet'!C3", 3, 3, 2, 2, "boring sheet", "updated sheet");
			Assert.AreEqual("'updated sheet'!F6", result);
		}

		[TestMethod]
		public void UpdateFormulaReferencesReferencingADifferentSheetIsNotUpdated()
		{
			var formulaManager = new FormulaManager();
			var result = formulaManager.UpdateFormulaReferences("'boring sheet'!C3", 3, 3, 2, 2, "boring sheet", "updated sheet");
			Assert.AreEqual("'boring sheet'!C3", result);
		}

		[TestMethod]
		public void UpdateFormulaReferencesPreservesEscapedQuotes()
		{
			var formulaManager = new FormulaManager();
			Assert.AreEqual("\"Hello,\"\" World\"&\"!\"", formulaManager.UpdateFormulaReferences("\"Hello,\"\" World\"&\"!\"", 1, 1, 8, 2, "Sheet", "Sheet"));
			Assert.AreEqual("FUNCTION(1,\"Hello World\",\"My name is \"\"Bob\"\"\",16)", formulaManager.UpdateFormulaReferences("FUNCTION(1, \"Hello World\", \"My name is \"\"Bob\"\"\", 16)", 1, 1, 8, 2, "Sheet", "Sheet"));
			Assert.AreEqual("FUNCTION(\"This is an example of \"\" Nested \"\"\"\" Quotes \"\".\")", formulaManager.UpdateFormulaReferences("FUNCTION(\"This is an example of \"\" Nested \"\"\"\" Quotes \"\".\")", 1, 1, 8, 2, "Sheet", "Sheet"));
		}
		#endregion

		#region UpdateFormulaSheetReferences Tests
		[TestMethod]
		public void UpdateFormulaSheetReferences()
		{
			var formulaManager = new FormulaManager();
			var result = formulaManager.UpdateFormulaSheetReferences("5+'OldSheet'!$G3+'Some Other Sheet'!C3+SUM(1,2,3)", "OldSheet", "NewSheet");
			Assert.AreEqual("5+'NewSheet'!$G3+'Some Other Sheet'!C3+SUM(1,2,3)", result);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void UpdateFormulaSheetReferencesNullOldSheetThrowsException()
		{
			var formulaManager = new FormulaManager();
			formulaManager.UpdateFormulaSheetReferences("formula", null, "sheet2");
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void UpdateFormulaSheetReferencesEmptyOldSheetThrowsException()
		{
			var formulaManager = new FormulaManager();
			formulaManager.UpdateFormulaSheetReferences("formula", string.Empty, "sheet2");
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void UpdateFormulaSheetReferencesNullNewSheetThrowsException()
		{
			var formulaManager = new FormulaManager();
			formulaManager.UpdateFormulaSheetReferences("formula", "sheet1", null);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void UpdateFormulaSheetReferencesEmptyNewSheetThrowsException()
		{
			var formulaManager = new FormulaManager();
			formulaManager.UpdateFormulaSheetReferences("formula", "sheet1", string.Empty);
		}
		#endregion

		#region UpdateFormulaDeletedSheetReferences Tests
		[TestMethod]
		public void UpdateFormulaDeletedSheetReference()
		{
			var formulaManager = new FormulaManager();
			string actualFormula = formulaManager.UpdateFormulaDeletedSheetReferences("CONCATENATE(Sheet1!B2, Sheet2!C3)", "sheet1");
			Assert.AreEqual("CONCATENATE(#REF!B2,'Sheet2'!C3)", actualFormula);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void UpdateFormulaDeletedSheetReferenceNullSheetNameThrowsException()
		{
			var formulaManager = new FormulaManager();
			string actualFormula = formulaManager.UpdateFormulaDeletedSheetReferences("CONCATENATE(Sheet1!B2, Sheet2!C3)", null);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void UpdateFormulaDeletedSheetReferenceEmptySheetNameThrowsException()
		{
			var formulaManager = new FormulaManager();
			string actualFormula = formulaManager.UpdateFormulaDeletedSheetReferences("CONCATENATE(Sheet1!B2, Sheet2!C3)", string.Empty);
		}
		#endregion
	}
}
