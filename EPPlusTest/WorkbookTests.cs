using System;
using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest
{
	[TestClass]
	public class WorkbookTests
	{
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
	}
}
