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
        public void WorkbookReferenceInNamedRangeDoesNotCorruptWorkbook()
        {
            var file = new FileInfo(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\..\Workbooks\ExternalLinks.xlsx"));
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
    }
}
