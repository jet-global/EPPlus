using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest
{
	[TestClass]
	public class WorkbookTests
	{
		#region ExternalReference Tests
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\ExternalLinks.xlsx")]
		public void WorkbookReferenceInNamedRangeDoesNotCorruptWorkbook()
		{
			var file = new FileInfo(@"ExternalLinks.xlsx");
			var newFile = new FileInfo("test.xlsx");
			if (newFile.Exists)
				newFile.Delete();
			Assert.IsTrue(file.Exists);
			try
			{
				using (var package = new ExcelPackage(file))
				{
					Assert.AreEqual("'[F:\\Cgypt\\Desktop\\test%20cases\\Demo%20Waterfall%20Chart.xlsx]Chart'!$A$1", package.Workbook.Names.First().Formula);
					package.SaveAs(newFile);
				}
				using (var package = new ExcelPackage(newFile))
				{
					Assert.AreEqual("'[F:\\Cgypt\\Desktop\\test%20cases\\Demo%20Waterfall%20Chart.xlsx]Chart'!$A$1", package.Workbook.Names.First().Formula);
				}

			}
			finally
			{
				if (newFile.Exists)
					newFile.Delete();
			}
		}

		[TestMethod]
		public void ExternalWorkbookReferenceIsLoadedWithIdAndName()
		{
			var testFile = new FileInfo(@"..\..\Workbooks\externalreferences.xlsx");
			var tempFile = new FileInfo(Path.GetTempFileName());
			if (tempFile.Exists)
				tempFile.Delete();
			testFile.CopyTo(tempFile.FullName);
			try
			{
				using (var package = new ExcelPackage(tempFile))
				{
					Assert.AreEqual(1, package.Workbook.ExternalReferences.References.Count);
					Assert.AreEqual(1, package.Workbook.ExternalReferences.References[0].Id);
					Assert.AreEqual(@"file:///E:\Source\Jet\ExternalWorkbook.xlam", package.Workbook.ExternalReferences.References[0].Name);
					package.Workbook.ExternalReferences.DeleteReference(1);
					Assert.AreEqual(0, package.Workbook.ExternalReferences.References.Count);
					package.Save();
				}
				// If no links exist then the entire collection must be removed for Excel to not corrupt it.
				// The ExternalReferences collection is null in this case.
				using (var package = new ExcelPackage(tempFile))
				{
					Assert.IsNull(package.Workbook.ExternalReferences);
				}
			}
			finally
			{
				tempFile.Delete();
			}
		}
		#endregion

		#region NamedRange Tests
		[TestMethod]
		public void SavePreservesRelativeWorkbookLevelNamedRanges()
		{
			var file = new FileInfo(Path.GetTempFileName());
			file.Delete();
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				package.Workbook.Names.Add("MyNamedRange", new ExcelRangeBase(sheet, "$C1"));
				Assert.AreEqual("'Sheet'!$C1", package.Workbook.Names["MyNamedRange"].FullAddress);
				package.SaveAs(file);
			}
			using (var package = new ExcelPackage(file))
			{
				Assert.AreEqual("'Sheet'!$C1", package.Workbook.Names["MyNamedRange"].FullAddress);
			}
		}
		#endregion
	}
}
