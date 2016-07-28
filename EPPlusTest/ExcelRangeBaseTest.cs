using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest
{
    [TestClass]
    public class ExcelRangeBaseTest : TestBase
    {
        [TestMethod]
        public void CopyCopiesCommentsFromSingleCellRanges()
        {
            InitBase();
            var pck = new ExcelPackage();
            var ws1 = pck.Workbook.Worksheets.Add("CommentCopying");
            var sourceExcelRange = ws1.Cells[3, 3];
            Assert.IsNull(sourceExcelRange.Comment);
            sourceExcelRange.AddComment("Testing comment 1", "test1");
            Assert.AreEqual("test1", sourceExcelRange.Comment.Author);
            Assert.AreEqual("Testing comment 1", sourceExcelRange.Comment.Text);
            var destinationExcelRange = ws1.Cells[5, 5];
            Assert.IsNull(destinationExcelRange.Comment);
            sourceExcelRange.Copy(destinationExcelRange);
            // Assert the original comment is intact.
            Assert.AreEqual("test1", sourceExcelRange.Comment.Author);
            Assert.AreEqual("Testing comment 1", sourceExcelRange.Comment.Text);
            // Assert the comment was copied.
            Assert.AreEqual("test1", destinationExcelRange.Comment.Author);
            Assert.AreEqual("Testing comment 1", destinationExcelRange.Comment.Text);
        }

        [TestMethod]
        public void CopyCopiesCommentsFromMultiCellRanges()
        {
            InitBase();
            var pck = new ExcelPackage();
            var ws1 = pck.Workbook.Worksheets.Add("CommentCopying");
            var sourceExcelRangeC3 = ws1.Cells[3, 3];
            var sourceExcelRangeD3 = ws1.Cells[3, 4];
            var sourceExcelRangeE3 = ws1.Cells[3, 5];
            Assert.IsNull(sourceExcelRangeC3.Comment);
            Assert.IsNull(sourceExcelRangeD3.Comment);
            Assert.IsNull(sourceExcelRangeE3.Comment);
            sourceExcelRangeC3.AddComment("Testing comment 1", "test1");
            sourceExcelRangeD3.AddComment("Testing comment 2", "test1");
            sourceExcelRangeE3.AddComment("Testing comment 3", "test1");
            Assert.AreEqual("test1", sourceExcelRangeC3.Comment.Author);
            Assert.AreEqual("Testing comment 1", sourceExcelRangeC3.Comment.Text);
            Assert.AreEqual("test1", sourceExcelRangeD3.Comment.Author);
            Assert.AreEqual("Testing comment 2", sourceExcelRangeD3.Comment.Text);
            Assert.AreEqual("test1", sourceExcelRangeE3.Comment.Author);
            Assert.AreEqual("Testing comment 3", sourceExcelRangeE3.Comment.Text);
            // Copy the full row to capture each cell at once.
            Assert.IsNull(ws1.Cells[5, 3].Comment);
            Assert.IsNull(ws1.Cells[5, 4].Comment);
            Assert.IsNull(ws1.Cells[5, 5].Comment);
            ws1.Cells["3:3"].Copy(ws1.Cells["5:5"]);
            // Assert the original comments are intact.
            Assert.AreEqual("test1", sourceExcelRangeC3.Comment.Author);
            Assert.AreEqual("Testing comment 1", sourceExcelRangeC3.Comment.Text);
            Assert.AreEqual("test1", sourceExcelRangeD3.Comment.Author);
            Assert.AreEqual("Testing comment 2", sourceExcelRangeD3.Comment.Text);
            Assert.AreEqual("test1", sourceExcelRangeE3.Comment.Author);
            Assert.AreEqual("Testing comment 3", sourceExcelRangeE3.Comment.Text);
            // Assert the comments were copied.
            var destinationExcelRangeC5 = ws1.Cells[5, 3];
            var destinationExcelRangeD5 = ws1.Cells[5, 4];
            var destinationExcelRangeE5 = ws1.Cells[5, 5];
            Assert.AreEqual("test1", destinationExcelRangeC5.Comment.Author);
            Assert.AreEqual("Testing comment 1", destinationExcelRangeC5.Comment.Text);
            Assert.AreEqual("test1", destinationExcelRangeD5.Comment.Author);
            Assert.AreEqual("Testing comment 2", destinationExcelRangeD5.Comment.Text);
            Assert.AreEqual("test1", destinationExcelRangeE5.Comment.Author);
            Assert.AreEqual("Testing comment 3", destinationExcelRangeE5.Comment.Text);
        }

        [TestMethod]
        public void SettingAddressHandlesMultiAddresses()
        {
            using (ExcelPackage package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");
                var name = package.Workbook.Names.Add("Test", worksheet.Cells[3, 3]);
                name.Address = "Sheet1!C3";
                name.Address = "Sheet1!D3";
                Assert.IsNull(name.Addresses);
                name.Address = "C3:D3,E3:F3";
                Assert.IsNotNull(name.Addresses);
                name.Address = "Sheet1!C3";
                Assert.IsNull(name.Addresses);
            }
        }

        [TestMethod]
        public void IfWithArray()
        {
            using (ExcelPackage package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");
                worksheet.Cells[3, 3].Formula = "IF(FALSE,\"true\",{\"false\"})";
                worksheet.Cells[3, 3].Calculate();
                CollectionAssert.AreEqual(new List<object> { "false" }, (List<object>)worksheet.Cells[3, 3].Value);
            }
        }

        [TestMethod]
        public void ArrayEquality()
        {
            using (ExcelPackage package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");
                worksheet.Cells[3, 3].Formula = "{\"test\"}=\"test\"";
                worksheet.Cells[3, 4].Formula = "{\"test\"}=\"\"\"test\"\"\"";
                worksheet.Cells[3, 5].Formula = "{\"test1\",\"test2\"}=\"test1\"";
                worksheet.Cells[3, 6].Formula = "{\"test1\",\"test2\"}={\"test1\"}";
                worksheet.Cells[3, 7].Formula = "{\"test1\",\"test2\"}={\"test1\",\"testB\"}";
                worksheet.Cells[3, 8].Formula = "{1,2,3}={1,2,3}";
                worksheet.Cells[3, 9].Formula = "{1,2,3}={1}";
                worksheet.Cells[3, 10].Formula = "{1,2,3}+4";
                worksheet.Calculate();
                Assert.IsTrue((bool)worksheet.Cells[3, 3].Value);
                Assert.IsFalse((bool)worksheet.Cells[3, 4].Value);
                Assert.IsTrue((bool)worksheet.Cells[3, 5].Value);
                Assert.IsTrue((bool)worksheet.Cells[3, 6].Value);
                Assert.IsTrue((bool)worksheet.Cells[3, 7].Value);
                Assert.IsTrue((bool)worksheet.Cells[3, 7].Value);
                Assert.IsTrue((bool)worksheet.Cells[3, 8].Value);
                Assert.IsTrue((bool)worksheet.Cells[3, 9].Value);
                Assert.AreEqual(5d, worksheet.Cells[3, 10].Value);
            }
        }

        [TestMethod]
        public void ArrayCell()
        {
            using (ExcelPackage package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");
                worksheet.Cells[3, 3].Formula = "{\"test1\",\"test2\"}";
                worksheet.Calculate();
                CollectionAssert.AreEqual(new List<object> { "test1", "test2" }, (List<object>)worksheet.Cells[3, 3].Value);
            }
        }

    }
}
