using System.IO;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Slicers;

namespace EPPlusTest.Drawing.Slicers
{
	[TestClass]
	public class ExcelSlicerDrawingTest
	{
		#region Constructor Tests
		[TestMethod]
		[DeploymentItem(@"Workbooks\SlicerWithSpecialCharacterName.xlsx")]
		public void ConstructorTestWithLineFeedInSlicerName()
		{
			var file = new FileInfo("SlicerWithSpecialCharacterName.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				ExcelWorksheet worksheet = package.Workbook.Worksheets.First();
				ExcelSlicer expectedSlicer = worksheet.Slicers.Slicers.First();
				Assert.AreEqual("Nombre _x000a_Cliente", expectedSlicer.Name);
				ExcelSlicerDrawing excelSlicerDrawing = (ExcelSlicerDrawing)(worksheet.Drawings.First());
				Assert.AreEqual("Nombre \nCliente", excelSlicerDrawing.Name);
				Assert.AreEqual(expectedSlicer, excelSlicerDrawing.Slicer);
			}
		}
		#endregion
	}
}
