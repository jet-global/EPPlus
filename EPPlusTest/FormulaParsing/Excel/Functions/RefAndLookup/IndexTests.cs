using EPPlusTest.FormulaParsing.TestHelpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using System.IO;

namespace EPPlusTest.FormulaParsing.Excel.Functions.RefAndLookup
{
    [TestClass]
    public class IndexTests
    {
        #region Properties
        private ParsingContext ParsingContext { get; set; }
        private ExcelPackage Package { get; set; }
        private ExcelWorksheet Worksheet { get; set; }
        #endregion

        #region TestInitialize/TestCleanup
        [TestInitialize]
        public void Initialize()
        {
            this.ParsingContext = ParsingContext.Create();
            this.Package = new ExcelPackage(new MemoryStream());
            this.Worksheet = this.Package.Workbook.Worksheets.Add("test");
        }

        [TestCleanup]
        public void Cleanup()
        {
            this.Package.Dispose();
        }
        #endregion

        #region Index Tests
        [TestMethod]
		public void IndexReturnsPoundValueWhenTooFewArgumentsAreSupplied()
		{
			this.Worksheet.Cells["A1"].Value = 1d;
			this.Worksheet.Cells["A2"].Value = 3d;
			this.Worksheet.Cells["A3"].Value = 5d;
			this.Worksheet.Cells["A4"].Formula = "INDEX(A1:A3,213)";
			this.Worksheet.Calculate();
			Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), this.Worksheet.Cells["A4"].Value);
		}

		[TestMethod]
        public void IndexReturnsValueByIndex()
        {
            var func = new Index();
            var result = func.Execute(FunctionsHelper.CreateArgs(FunctionsHelper.CreateArgs(1, 2, 5), 3), this.ParsingContext);
            Assert.AreEqual(5, result.Result);
        }

        [TestMethod]
        public void IndexHandlesSingleRange()
        {
            this.Worksheet.Cells["A1"].Value = 1d;
            this.Worksheet.Cells["A2"].Value = 3d;
            this.Worksheet.Cells["A3"].Value = 5d;
            this.Worksheet.Cells["A4"].Formula = "INDEX(A1:A3,3)";
            this.Worksheet.Calculate();
            Assert.AreEqual(5d, this.Worksheet.Cells["A4"].Value);
        }

        [TestMethod]
        public void IndexHandlesNAError()
        {
            this.Worksheet.Cells["A1"].Value = 1d;
            this.Worksheet.Cells["A2"].Value = 3d;
            this.Worksheet.Cells["A3"].Value = 5d;
            this.Worksheet.Cells["A4"].Value = ExcelErrorValue.Create(eErrorType.NA);
            this.Worksheet.Cells["A5"].Formula = "INDEX(A1:A3,A4)";
            this.Worksheet.Calculate();
            Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), this.Worksheet.Cells["A5"].Value);
        }
        #endregion
    }
}
