using System;
using System.Collections.Generic;
using System.IO;
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
        public void OverwrittenSharedFormulaRowsAreRespected()
        {
            // In Excel, a cell in a shared formula range can have its formula replaced by something
            // that is not the shared formula. EP Plus overwrites this formula when splitting shared
            // formulas.
            FileInfo newFile = new FileInfo(Path.GetTempFileName());
            if (newFile.Exists)
                newFile.Delete();
            try
            {
                var dir = AppDomain.CurrentDomain.BaseDirectory;
                using (var pck = new ExcelPackage(new FileInfo(Path.Combine(dir, "Workbooks", "SharedFormulasRows.xlsx"))))
                {
                    var sheet = pck.Workbook.Worksheets["Sheet1"];
                    // This formula is in the shared formula range, but was explicitly overwritten in Excel.
                    var nonSharedFormulaOriginal = sheet.Cells["C5"].Formula;
                    // Set some other cell's formula in the shared formula range to trigger a split.
                    sheet.Cells["C4"].Formula = "SUM(1,2)";
                    // Verify that the explicit formula in the shared formula range was NOT overwritten.
                    var nonSharedFormulaUpdated = sheet.Cells["C5"].Formula;
                    Assert.AreEqual(nonSharedFormulaOriginal, nonSharedFormulaUpdated);
                    // Ensure that the shared formula was propagated to the region.
                    var sharedFormulaUpdated = sheet.Cells["C6"].Formula;
                    Assert.AreEqual("B6+D6", sharedFormulaUpdated);
                    pck.SaveAs(newFile);
                }
                // Ensure the integrity of the package is maintained.
                using (var pck = new ExcelPackage(newFile))
                {
                    var sheet = pck.Workbook.Worksheets["Sheet1"];
                    Assert.AreEqual("B3+D3", sheet.Cells["C3"].Formula);
                    Assert.AreEqual("SUM(1,2)", sheet.Cells["C4"].Formula);
                    Assert.AreEqual("B5*D5", sheet.Cells["C5"].Formula);
                    Assert.AreEqual("B6+D6", sheet.Cells["C6"].Formula);
                }
            }
            finally
            {
                if (newFile.Exists)
                    newFile.Delete();
            }
        }

        [TestMethod]
        public void OverwrittenSharedFormulaColumnsAreRespected()
        {
            // In Excel, a cell in a shared formula range can have its formula replaced by something
            // that is not the shared formula. EP Plus overwrites this formula when splitting shared
            // formulas.
            FileInfo newFile = new FileInfo(Path.GetTempFileName());
            if (newFile.Exists)
                newFile.Delete();
            try
            {
                var dir = AppDomain.CurrentDomain.BaseDirectory;
                using (var pck = new ExcelPackage(new FileInfo(Path.Combine(dir, "Workbooks", "SharedFormulasColumns.xlsx"))))
                {
                    var sheet = pck.Workbook.Worksheets["Sheet1"];
                    // This formula is in the shared formula range, but was explicitly overwritten in Excel.
                    var nonSharedFormulaOriginal = sheet.Cells["E3"].Formula;
                    // Set some other cell's formula in the shared formula range to trigger a split.
                    sheet.Cells["D3"].Formula = "SUM(1,2)";
                    // Verify that the explicit formula in the shared formula range was NOT overwritten.
                    var nonSharedFormulaUpdated = sheet.Cells["E3"].Formula;
                    Assert.AreEqual(nonSharedFormulaOriginal, nonSharedFormulaUpdated);
                    // Ensure that the shared formula was propagated to the region.
                    var sharedFormulaUpdated = sheet.Cells["F3"].Formula;
                    Assert.AreEqual("F2+F4", sharedFormulaUpdated);
                    pck.SaveAs(newFile);
                }
                // Ensure the integrity of the package is maintained.
                using (var pck = new ExcelPackage(newFile))
                {
                    var sheet = pck.Workbook.Worksheets["Sheet1"];
                    Assert.AreEqual("C2+C4", sheet.Cells["C3"].Formula);
                    Assert.AreEqual("SUM(1,2)", sheet.Cells["D3"].Formula);
                    Assert.AreEqual("E2*E4", sheet.Cells["E3"].Formula);
                    Assert.AreEqual("F2+F4", sheet.Cells["F3"].Formula);
                }
            }
            finally
            {
                if (newFile.Exists)
                    newFile.Delete();
            }
        }

        [TestMethod]
        public void FullyOverwrittenSharedRowFormulaIsRemovedCorrectly()
        {
            // There is a bug where Shared Formulas that are overwritten from the bottom up cause
            // an exception on the final cell in the Shared Formula range.
            FileInfo newFile = new FileInfo(Path.GetTempFileName());
            if (newFile.Exists)
                newFile.Delete();
            try
            {
                var dir = AppDomain.CurrentDomain.BaseDirectory;
                using (var pck = new ExcelPackage())
                {
                    var sheet = pck.Workbook.Worksheets.Add("Sheet1");
                    sheet.Cells["C3:C4"].Formula = "B3+B3";
                    Assert.AreEqual("B3+B3", sheet.Cells[3, 3].Formula);
                    Assert.AreEqual("B4+B4", sheet.Cells[4, 3].Formula);
                    sheet.Cells[4, 3].Formula = "B4-B4";
                    Assert.AreEqual("B3+B3", sheet.Cells[3, 3].Formula);
                    Assert.AreEqual("B4-B4", sheet.Cells[4, 3].Formula);
                    sheet.Cells[3, 3].Formula = "B3-B3";
                    Assert.AreEqual("B3-B3", sheet.Cells[3, 3].Formula);
                    Assert.AreEqual("B4-B4", sheet.Cells[4, 3].Formula);
                    pck.SaveAs(newFile);
                }
                // Ensure the integrity of the package is maintained.
                using (var pck = new ExcelPackage(newFile))
                {
                    var sheet = pck.Workbook.Worksheets["Sheet1"];
                    Assert.AreEqual("B3-B3", sheet.Cells[3, 3].Formula);
                    Assert.AreEqual("B4-B4", sheet.Cells[4, 3].Formula);
                }
            }
            finally
            {
                if (newFile.Exists)
                    newFile.Delete();
            }
        }

        [TestMethod]
        public void FullyOverwrittenSharedColumnFormulaIsRemovedCorrectly()
        {
            // There is a bug where Shared Formulas that are overwritten from the bottom up cause
            // an exception on the final cell in the Shared Formula range.
            FileInfo newFile = new FileInfo(Path.GetTempFileName());
            if (newFile.Exists)
                newFile.Delete();
            try
            {
                var dir = AppDomain.CurrentDomain.BaseDirectory;
                using (var pck = new ExcelPackage())
                {
                    var sheet = pck.Workbook.Worksheets.Add("Sheet1");
                    sheet.Cells["C3:D3"].Formula = "C4+C4";
                    Assert.AreEqual("C4+C4", sheet.Cells[3, 3].Formula);
                    Assert.AreEqual("D4+D4", sheet.Cells[3, 4].Formula);
                    sheet.Cells[3, 4].Formula = "D4-D4";
                    Assert.AreEqual("C4+C4", sheet.Cells[3, 3].Formula);
                    Assert.AreEqual("D4-D4", sheet.Cells[3, 4].Formula);
                    sheet.Cells[3, 3].Formula = "C4-C4";
                    Assert.AreEqual("C4-C4", sheet.Cells[3, 3].Formula);
                    Assert.AreEqual("D4-D4", sheet.Cells[3, 4].Formula);
                    pck.SaveAs(newFile);
                }
                // Ensure the integrity of the package is maintained.
                using (var pck = new ExcelPackage(newFile))
                {
                    var sheet = pck.Workbook.Worksheets["Sheet1"];
                    Assert.AreEqual("C4-C4", sheet.Cells[3, 3].Formula);
                    Assert.AreEqual("D4-D4", sheet.Cells[3, 4].Formula);
                }
            }
            finally
            {
                if (newFile.Exists)
                    newFile.Delete();
            }
        }

        [TestMethod]
        public void SetFormulaRemovesLeadingEquals()
        {
            using (var pkg = new ExcelPackage())
            {
                var sheet = pkg.Workbook.Worksheets.Add("Sheet");
                sheet.Cells[3, 3].Formula = "=SUM(1,2)";
                Assert.AreEqual("SUM(1,2)", sheet.Cells[3, 3].Formula);
            }
        }

        [TestMethod]
        public void SetSharedFormulaRemovesLeadingEquals()
        {
            using (var pkg = new ExcelPackage())
            {
                var sheet = pkg.Workbook.Worksheets.Add("Sheet");
                sheet.Cells[3, 3, 5, 5].Formula = "=SUM(1,2)";
                Assert.AreEqual("SUM(1,2)", sheet.Cells[3, 3].Formula);
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
        public void EPPlusRegressionColumnAutoFitResizesColumnToFitContentsWithDefaultRowHeight()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet");
                sheet.Cells[1, 2].Value = "Header to set width";
                sheet.Cells[1, 2].Style.WrapText = true;
                var contents = "=0";
                sheet.Cells[2, 2].Formula = contents;
                sheet.Cells[2, 2].Style.Numberformat.Format = "_(* #,##0_);_(* (#,##0);_(* \" - \"??_);_(@_)";
                sheet.Column(2).Width = 2;
                Assert.AreEqual(2, sheet.Column(2).Width);
                sheet.Cells[2, 3].Formula = "=\"Next cell contents.\"";
                sheet.Cells[2, 2].Calculate();
                sheet.Column(2).AutoFit();
                var actualWidth = sheet.Column(2).Width;
                Assert.AreEqual(19, Math.Round(sheet.Column(2).Width));
            }
        }

        [TestMethod]
        public void EPPlusRegressionColumnAutoFitResizesColumnToFitContentsWithNoSpacesRegardlessOfRowHeight()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet");
                sheet.Row(1).Height = 33; // Just a bit taller than two lines of text.
                sheet.Cells[1, 2].Value = "HeaderWithNoSpacesShouldNotBreak";
                sheet.Cells[1, 2].Style.WrapText = true;
                var contents = "=0";
                sheet.Cells[2, 2].Formula = contents;
                sheet.Cells[2, 2].Style.Numberformat.Format = "_(* #,##0_);_(* (#,##0);_(* \" - \"??_);_(@_)";
                sheet.Column(2).Width = 2;
                Assert.AreEqual(2, sheet.Column(2).Width);
                sheet.Cells[2, 3].Formula = "=\"Next cell contents.\"";
                sheet.Cells[2, 2].Calculate();
                sheet.Column(2).AutoFit();
                Assert.AreEqual(35, Math.Round(sheet.Column(2).Width));
            }
        }

        [TestMethod]
        public void EPPlusRegressionColumnAutoFitResizesColumnToFitContentsWithSpecifiedRowHeight()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet");
                sheet.Row(1).Height = 33; // Just a bit taller than two lines of text.
                sheet.Cells[1, 2].Value = "Header to set width";
                sheet.Cells[1, 2].Style.WrapText = true;
                var contents = "=0";
                sheet.Cells[2, 2].Formula = contents;
                sheet.Cells[2, 2].Style.Numberformat.Format = "_(* #,##0_);_(* (#,##0);_(* \" - \"??_);_(@_)";
                sheet.Column(2).Width = 2;
                Assert.AreEqual(2, sheet.Column(2).Width);
                sheet.Cells[2, 3].Formula = "=\"Next cell contents.\"";
                sheet.Cells[2, 2].Calculate();
                sheet.Column(2).AutoFit();
                Assert.AreEqual(10, Math.Round(sheet.Column(2).Width));

            }
        }

        [TestMethod]
        public void EPPlusRegressionColumnAutoFitResizesColumnToFitContentsWithSpecifiedRowHeightForFiveLines()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("Sheet");
                sheet.Row(1).Height = 76; // Just a bit taller than five lines of text.
                sheet.Cells[1, 2].Value = "Set width to five lines";
                sheet.Cells[1, 2].Style.WrapText = true;
                var contents = "=0";
                sheet.Cells[2, 2].Formula = contents;
                sheet.Cells[2, 2].Style.Numberformat.Format = "_(* #,##0_);_(* (#,##0);_(* \" - \"??_);_(@_)";
                sheet.Column(2).Width = 2;
                Assert.AreEqual(2, sheet.Column(2).Width);
                sheet.Cells[2, 3].Formula = "=\"Next cell contents.\"";
                sheet.Cells[2, 2].Calculate();
                sheet.Column(2).AutoFit();
                Assert.AreEqual(9, Math.Round(sheet.Column(2).Width));

            }
        }
    }
}
