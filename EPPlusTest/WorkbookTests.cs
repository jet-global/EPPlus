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
					Assert.AreEqual("[1]Chart!$A$1", package.Workbook.Names.ElementAt(0).NameFormula);
					Assert.AreEqual("[1]Sheet1!XFA1048568", package.Workbook.Names.ElementAt(1).NameFormula);
					// Verify that references containing external links are never updated.
					package.Workbook.Worksheets.First().InsertRow(1, 5);
					package.Workbook.Worksheets.First().InsertColumn(1, 5);
					Assert.AreEqual("[1]Chart!$A$1", package.Workbook.Names.ElementAt(0).NameFormula);
					var name = package.Workbook.Names.ElementAt(1);
					Assert.AreEqual("[1]Sheet1!XFA1048568", name.NameFormula);
					Assert.AreEqual("[1]Sheet1!XFA1048568", string.Join(string.Empty, name.GetRelativeNameFormula(3, 3).Select(n => n.Value)));
					package.SaveAs(newFile);
				}
				using (var package = new ExcelPackage(newFile))
				{
					Assert.AreEqual("[1]Chart!$A$1", package.Workbook.Names.ElementAt(0).NameFormula);
					Assert.AreEqual("[1]Sheet1!XFA1048568", package.Workbook.Names.ElementAt(1).NameFormula);
				}
			}
			finally
			{
				if (newFile.Exists)
					newFile.Delete();
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\externalreferences.xlsx")]
		public void ExternalWorkbookReferenceIsLoadedWithIdAndName()
		{
			var testFile = new FileInfo(@"externalreferences.xlsx");
			var tempFile = new FileInfo(Path.GetTempFileName());
			if (tempFile.Exists)
				tempFile.Delete();
			testFile.CopyTo(tempFile.FullName);
			try
			{
				using (var package = new ExcelPackage(tempFile))
				{
					Assert.AreEqual(2, package.Workbook.ExternalReferences.References.Count);
					Assert.AreEqual(1, package.Workbook.ExternalReferences.References[0].Id);
					Assert.AreEqual(@"/Source/Jet/ExternalWorkbook.xlam", package.Workbook.ExternalReferences.References[0].Name);
					package.Workbook.ExternalReferences.DeleteReference(1);
					Assert.AreEqual(1, package.Workbook.ExternalReferences.References.Count);
					package.Save();
				}
				// If no links exist then the entire collection must be removed for Excel to not corrupt it.
				// The ExternalReferences collection is null in this case.
				using (var package = new ExcelPackage(tempFile))
				{
					Assert.AreEqual(1, package.Workbook.ExternalReferences.References.Count);
					Assert.AreEqual(1, package.Workbook.ExternalReferences.References[0].Id);
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

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\externalreferences.xlsx")]
		public void ExternalWorkbookReferencesPoundRefAsFunctions()
		{
			var testFile = new FileInfo(@"externalreferences.xlsx");
			var tempFile = new FileInfo(Path.GetTempFileName());
			if (tempFile.Exists)
				tempFile.Delete();
			testFile.CopyTo(tempFile.FullName);
			try
			{
				using (var package = new ExcelPackage(tempFile))
				{
					Assert.AreEqual(2, package.Workbook.ExternalReferences.References.Count);
					var sheet = package.Workbook.Worksheets.First();
					sheet.Cells["F9"].Calculate();
					Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Ref), sheet.Cells["F9"].Value);
				}
			}
			finally
			{
				tempFile.Delete();
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\externalreferences.xlsx")]
		public void ExternalWorkbookReferencesPoundRefAsNamedRanges()
		{
			var testFile = new FileInfo(@"externalreferences.xlsx");
			var tempFile = new FileInfo(Path.GetTempFileName());
			if (tempFile.Exists)
				tempFile.Delete();
			testFile.CopyTo(tempFile.FullName);
			try
			{
				using (var package = new ExcelPackage(tempFile))
				{
					Assert.AreEqual(2, package.Workbook.ExternalReferences.References.Count);
					var sheet = package.Workbook.Worksheets.First();
					sheet.Cells["F15"].Calculate();
					Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Ref), sheet.Cells["F15"].Value);
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
			if (file.Exists)
				file.Delete();
			try
			{
				using (var package = new ExcelPackage())
				{
					var sheet = package.Workbook.Worksheets.Add("Sheet");
					package.Workbook.Names.Add("MyNamedRange", new ExcelRangeBase(sheet, "$C1"));
					Assert.AreEqual("'Sheet'!$C1", package.Workbook.Names["MyNamedRange"].NameFormula);
					package.SaveAs(file);
				}
				using (var package = new ExcelPackage(file))
				{
					Assert.AreEqual("'Sheet'!$C1", package.Workbook.Names["MyNamedRange"].NameFormula);
				}
			}
			finally
			{
				file.Delete();
			}
		}
		#endregion
	}
}
