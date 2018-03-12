using System;
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

		#region Integration Tests
		[TestMethod]
		public void VisibilityComments()
		{
			var xlsxName = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsx");
			try
			{
				using (var ms = File.Open(xlsxName, FileMode.OpenOrCreate))
				using (var pkg = new ExcelPackage(ms))
				{
					var ws = pkg.Workbook.Worksheets.Add("Comment");
					var a1 = ws.Cells["A1"];
					a1.Value = "Justin Dearing";
					a1.AddComment("I am A1s comment", "JD");
					Assert.IsFalse(a1.Comment.Visible); // Comments are by default invisible 
					a1.Comment.Visible = true;
					a1.Comment.Visible = false;
					Assert.IsNotNull(a1.Comment);
					//check style attribute
					var stylesDict = new System.Collections.Generic.Dictionary<string, string>();
					string[] styles = a1.Comment.Style
						 .Split(new[] { ";" }, StringSplitOptions.RemoveEmptyEntries);
					Array.ForEach(styles, s =>
					{
						string[] split = s.Split(':');
						if (split.Length == 2)
						{
							var k = (split[0] ?? "").Trim().ToLower();
							var v = (split[1] ?? "").Trim().ToLower();
							stylesDict[k] = v;
						}
					});
					Assert.IsTrue(stylesDict.ContainsKey("visibility"));
					Assert.AreEqual("hidden", stylesDict["visibility"]);
					Assert.IsFalse(a1.Comment.Visible);
					pkg.Save();
					ms.Close();
				}
			}
			finally
			{
				//open results file in program for view xlsx.
				//comments of cell A1 must be hidden.
				//System.Diagnostics.Process.Start(Path.GetDirectoryName(xlsxName));
				File.Delete(xlsxName);
			}
		}
		#endregion
	}
}
