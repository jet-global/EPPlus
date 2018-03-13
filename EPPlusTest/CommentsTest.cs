using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest
{
	[TestClass]
	public class CommentsTest
	{
		#region AddComment Tests
		[TestMethod]
		public void AddCommentTest()
		{
			var tempFile = new FileInfo(Path.GetTempFileName());
			if (tempFile.Exists)
				tempFile.Delete();
			try
			{
				using (var package = new ExcelPackage())
				{
					var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
					var sheet2 = package.Workbook.Worksheets.Add("Sheet2");
					sheet1.Cells[2, 2].Value = "testvalue1";
					sheet1.Cells[2, 2].AddComment("Comment text", "an author");

					sheet2.Cells[2, 2].Value = "testvalue2";
					sheet2.Cells[2, 2].AddComment("Comment text 2", "another author");
					Assert.AreEqual("Comment text", sheet1.Cells[2, 2].Comment.Text);
					Assert.AreEqual("an author", sheet1.Cells[2, 2].Comment.Author);
					Assert.AreEqual("Comment text 2", sheet2.Cells[2, 2].Comment.Text);
					Assert.AreEqual("another author", sheet2.Cells[2, 2].Comment.Author);
					package.SaveAs(tempFile);
				}
				using (var package = new ExcelPackage(tempFile))
				{
					var sheet1 = package.Workbook.Worksheets["Sheet1"];
					var sheet2 = package.Workbook.Worksheets["Sheet2"];
					Assert.AreEqual("Comment text", sheet1.Cells[2,2].Comment.Text);
					Assert.AreEqual("an author", sheet1.Cells[2,2].Comment.Author);
					Assert.AreEqual("Comment text 2", sheet2.Cells[2, 2].Comment.Text);
					Assert.AreEqual("another author", sheet2.Cells[2, 2].Comment.Author);
				}
			}
			finally
			{
				if (tempFile.Exists)
					tempFile.Delete();
			}
		}

		[TestMethod]
		public void AddCommentAfterRowInsert()
		{
			var tempFile = new FileInfo(Path.GetTempFileName());
			if (tempFile.Exists)
				tempFile.Delete();
			try
			{
				using (var package = new ExcelPackage())
				{
					var sheet = package.Workbook.Worksheets.Add("Sheet1");
					sheet.Cells[2, 2].Value = "testdata";
					sheet.Cells[2, 2].AddComment("testMessage1", "an author");
					sheet.InsertRow(2, 1);
					sheet.Cells[2, 2].AddComment("testMessage2", "another author");
					package.SaveAs(tempFile);
				}
				using (var package = new ExcelPackage(tempFile))
				{
					var sheet = package.Workbook.Worksheets["Sheet1"];
					Assert.AreEqual("testMessage2", sheet.Cells[2, 2].Comment.Text);
					Assert.AreEqual("another author", sheet.Cells[2, 2].Comment.Author);
					Assert.AreEqual("testMessage1", sheet.Cells[3, 2].Comment.Text);
					Assert.AreEqual("an author", sheet.Cells[3, 2].Comment.Author);
				}
			}
			finally
			{
				if (tempFile.Exists)
					tempFile.Delete();
			}
		}

		[TestMethod]
		public void AddCommentAfterRowDelete()
		{
			var tempFile = new FileInfo(Path.GetTempFileName());
			if (tempFile.Exists)
				tempFile.Delete();
			try
			{
				using (var package = new ExcelPackage())
				{
					var sheet = package.Workbook.Worksheets.Add("Sheet1");
					sheet.Cells[3, 3].Value = "testdata";
					sheet.Cells[3, 3].AddComment("testMessage1", "an author");
					sheet.DeleteRow(2, 1);
					sheet.Cells[3, 3].AddComment("testMessage2", "another author");
					package.SaveAs(tempFile);
				}
				using (var package = new ExcelPackage(tempFile))
				{
					var sheet = package.Workbook.Worksheets["Sheet1"];
					Assert.AreEqual("testMessage1", sheet.Cells[2, 3].Comment.Text);
					Assert.AreEqual("an author", sheet.Cells[2, 3].Comment.Author);
					Assert.AreEqual("testMessage2", sheet.Cells[3, 3].Comment.Text);
					Assert.AreEqual("another author", sheet.Cells[3, 3].Comment.Author);
				}
			}
			finally
			{
				if (tempFile.Exists)
					tempFile.Delete();
			}
		}

		#endregion
		#region Copy Comment Tests
		[TestMethod]
		public void CopyCommentSameWorksheet()
		{
			var tempFile = new FileInfo(Path.GetTempFileName());
			if (tempFile.Exists)
				tempFile.Delete();
			try
			{
				using (var package = new ExcelPackage())
				{
					var sheet = package.Workbook.Worksheets.Add("Sheet1");
					sheet.Cells[2, 2].AddComment("testMessage1", "an author");
					sheet.Cells[2, 2].Copy(sheet.Cells[3, 2]);
					Assert.AreEqual("testMessage1", sheet.Cells[2, 2].Comment.Text);
					Assert.AreEqual("an author", sheet.Cells[2, 2].Comment.Author);
					Assert.AreEqual("testMessage1", sheet.Cells[3, 2].Comment.Text);
					Assert.AreEqual("an author", sheet.Cells[3, 2].Comment.Author);
					package.SaveAs(tempFile);
				}
				using (var package = new ExcelPackage(tempFile))
				{
					var sheet = package.Workbook.Worksheets["Sheet1"];
					Assert.AreEqual("testMessage1", sheet.Cells[2, 2].Comment.Text);
					Assert.AreEqual("an author", sheet.Cells[2, 2].Comment.Author);
					Assert.AreEqual("testMessage1", sheet.Cells[3, 2].Comment.Text);
					Assert.AreEqual("an author", sheet.Cells[3, 2].Comment.Author);
				}
			}
			finally
			{
				if (tempFile.Exists)
					tempFile.Delete();
			}
		}

		[TestMethod]
		public void CopyCommentDifferentWorksheet()
		{
			var tempFile = new FileInfo(Path.GetTempFileName());
			if (tempFile.Exists)
				tempFile.Delete();
			try
			{
				using (var package = new ExcelPackage())
				{
					var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
					var sheet2 = package.Workbook.Worksheets.Add("Sheet2");
					sheet1.Cells[2, 2].AddComment("testMessage1", "an author");
					sheet1.Cells[2, 2].Copy(sheet2.Cells[3, 2]);
					Assert.AreEqual("testMessage1", sheet1.Cells[2, 2].Comment.Text);
					Assert.AreEqual("an author", sheet1.Cells[2, 2].Comment.Author);
					Assert.AreEqual("testMessage1", sheet2.Cells[3, 2].Comment.Text);
					Assert.AreEqual("an author", sheet2.Cells[3, 2].Comment.Author);
					package.SaveAs(tempFile);
				}
				using (var package = new ExcelPackage(tempFile))
				{
					var sheet1 = package.Workbook.Worksheets["Sheet1"];
					var sheet2 = package.Workbook.Worksheets["Sheet2"];
					Assert.AreEqual("testMessage1", sheet1.Cells[2, 2].Comment.Text);
					Assert.AreEqual("an author", sheet1.Cells[2, 2].Comment.Author);
					Assert.AreEqual("testMessage1", sheet2.Cells[3, 2].Comment.Text);
					Assert.AreEqual("an author", sheet2.Cells[3, 2].Comment.Author);
				}
			}
			finally
			{
				if (tempFile.Exists)
					tempFile.Delete();
			}
		}

		[TestMethod]
		public void CopyWorksheetCopiesComments()
		{
			var tempFile = new FileInfo(Path.GetTempFileName());
			if (tempFile.Exists)
				tempFile.Delete();
			try
			{
				using (var package = new ExcelPackage())
				{
					var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
					sheet1.Cells[2, 2].AddComment("testMessage1", "an author");
					sheet1.Cells[3, 2].AddComment("testMessage2", "another author");
					var sheet1Copy = package.Workbook.Worksheets.Add("sheet1 copy", sheet1);
					Assert.AreEqual("testMessage1", sheet1.Cells[2, 2].Comment.Text);
					Assert.AreEqual("an author", sheet1.Cells[2, 2].Comment.Author);
					Assert.AreEqual("testMessage2", sheet1.Cells[3, 2].Comment.Text);
					Assert.AreEqual("another author", sheet1.Cells[3, 2].Comment.Author);

					Assert.AreEqual("testMessage1", sheet1Copy.Cells[2, 2].Comment.Text);
					Assert.AreEqual("an author", sheet1Copy.Cells[2, 2].Comment.Author);
					Assert.AreEqual("testMessage2", sheet1Copy.Cells[3, 2].Comment.Text);
					Assert.AreEqual("another author", sheet1Copy.Cells[3, 2].Comment.Author);
					package.SaveAs(tempFile);
				}
				using (var package = new ExcelPackage(tempFile))
				{
					var sheet1 = package.Workbook.Worksheets["Sheet1"];
					var sheet1Copy = package.Workbook.Worksheets["Sheet1 copy"];
					Assert.AreEqual("testMessage1", sheet1.Cells[2, 2].Comment.Text);
					Assert.AreEqual("an author", sheet1.Cells[2, 2].Comment.Author);
					Assert.AreEqual("testMessage2", sheet1.Cells[3, 2].Comment.Text);
					Assert.AreEqual("another author", sheet1.Cells[3, 2].Comment.Author);
					Assert.AreEqual("testMessage1", sheet1Copy.Cells[2, 2].Comment.Text);
					Assert.AreEqual("an author", sheet1Copy.Cells[2, 2].Comment.Author);
					Assert.AreEqual("testMessage2", sheet1Copy.Cells[3, 2].Comment.Text);
					Assert.AreEqual("another author", sheet1Copy.Cells[3, 2].Comment.Author);
				}
			}
			finally
			{
				if (tempFile.Exists)
					tempFile.Delete();
			}
		}
		#endregion
	}
}
