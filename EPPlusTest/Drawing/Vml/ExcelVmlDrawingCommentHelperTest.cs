using System;
using System.IO;
using System.Xml;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Vml;

namespace EPPlusTest.Drawing.Vml
{
	[TestClass]
	public class ExcelVmlDrawingCommentHelperTest
	{
		#region Test Methods
		[TestMethod]
		public void AddCommentDrawingsToEmptyPackage()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet1");
				var vmlDrawingsUri = XmlHelper.GetNewUri(sheet.Package.Package, @"/xl/drawings/vmlDrawing{0}.vml");
				Assert.IsFalse(sheet.Package.Package.TryGetPart(vmlDrawingsUri, out var vmlDrawingsPart));
				var commentCollection = new ExcelCommentCollection(package, sheet, sheet.NameSpaceManager)
				{
					{ sheet.Cells[2, 2], "commenttext1", "author" },
					{ sheet.Cells[3, 2], "commenttext2", "author" },
					{ sheet.Cells[4, 2], "commenttext3", "author" }
				};
				ExcelVmlDrawingCommentHelper.AddCommentDrawings(sheet, commentCollection);
				Assert.IsTrue(sheet.Package.Package.TryGetPart(vmlDrawingsUri, out vmlDrawingsPart));
			}
		}

		[TestMethod]
		public void AddCommentDrawingsToPackageWithExistingVml()
		{
			var tempFile = new FileInfo(Path.GetTempFileName());
			if (tempFile.Exists)
				tempFile.Delete();
			try
			{
				using (var package = new ExcelPackage())
				{
					var sheet = package.Workbook.Worksheets.Add("Sheet1");
					var vmlDrawingsUri = XmlHelper.GetNewUri(sheet.Package.Package, @"/xl/drawings/vmlDrawing{0}.vml");
					var commentCollection = new ExcelCommentCollection(package, sheet, sheet.NameSpaceManager)
					{
						{ sheet.Cells[2, 2], "commenttext1", "author" },
						{ sheet.Cells[3, 2], "commenttext2", "author" },
						{ sheet.Cells[4, 2], "commenttext3", "author" }
					};
					ExcelVmlDrawingCommentHelper.AddCommentDrawings(sheet, commentCollection);
					Assert.IsTrue(sheet.Package.Package.TryGetPart(vmlDrawingsUri, out var vmlDrawingsPart));
					package.SaveAs(tempFile);
				}
				using (var package = new ExcelPackage(tempFile))
				{
					var sheet = package.Workbook.Worksheets["Sheet1"];
					var vmlDrawingsUri = XmlHelper.GetNewUri(sheet.Package.Package, @"/xl/drawings/vmlDrawing{0}.vml");
					var commentCollection = new ExcelCommentCollection(package, sheet, sheet.NameSpaceManager)
					{
						{ sheet.Cells[2, 2], "newcommenttext1", "author" },
						{ sheet.Cells[3, 2], "newcommenttext2", "author" }
					};
					ExcelVmlDrawingCommentHelper.AddCommentDrawings(sheet, commentCollection);
					Assert.IsTrue(sheet.Package.Package.TryGetPart(vmlDrawingsUri, out var vmlDrawingsPart));
					var xmlDoc = new XmlDocument();
					xmlDoc.Load(vmlDrawingsPart.GetStream());
					var nsmgr = new XmlNamespaceManager(new NameTable());
					nsmgr.AddNamespace("v", "urn:schemas-microsoft-com:vml");
					var nodes = xmlDoc.SelectNodes("/xml/v:shape", nsmgr);
					Assert.AreEqual(2, nodes.Count);
				}
			}
			finally
			{
				if (tempFile.Exists)
					tempFile.Delete();
			}
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void AddCommentDrawingsNullWorksheetThrowsException()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet1");
				var vmlDrawingsUri = XmlHelper.GetNewUri(sheet.Package.Package, @"/xl/drawings/vmlDrawing{0}.vml");
				var commentCollection = new ExcelCommentCollection(package, sheet, sheet.NameSpaceManager);
				ExcelVmlDrawingCommentHelper.AddCommentDrawings(null, commentCollection);
			}
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void AddCommentDrawingsNullCommentCollectionThrowsException()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet1");
				ExcelVmlDrawingCommentHelper.AddCommentDrawings(sheet, null);
			}
		}
		#endregion
	}
}
