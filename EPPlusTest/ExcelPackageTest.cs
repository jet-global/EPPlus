using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest
{
	[TestClass]
	public class ExcelPackageTest
	{
		#region Configure Tests
		[TestMethod]
		public void ConfigureAppliesGivenIFormulaManager()
		{
			var formulaManager = new FormulaManager();
			var excelPackage = new ExcelPackage();
			Assert.AreNotEqual(formulaManager, excelPackage.FormulaManager);
			excelPackage.Configure(formulaManager);
			Assert.AreEqual(formulaManager, excelPackage.FormulaManager);
		}

		[TestMethod]
		public void ConfigureHandlesNullIFormulaManager()
		{
			var excelPackage = new ExcelPackage();
			Assert.IsInstanceOfType(excelPackage.FormulaManager, typeof(FormulaManager));
			excelPackage.Configure((IFormulaManager)null);
			Assert.IsInstanceOfType(excelPackage.FormulaManager, typeof(FormulaManager));
		}
		#endregion
	}
}
