using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;

namespace EPPlusTest.Drawing
{
	[TestClass]
	public class ExcelDrawingsTest
	{
		#region Test Methods
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\OfPieChart.xlsx")]
		public void ExcelDrawingsWithExcelOfPieChartTest()
		{
			using (var package = new ExcelPackage(new FileInfo("OfPieChart.xlsx")))
			{
				var drawings = new ExcelDrawings(package, package.Workbook.Worksheets["Dashboard"]);
				var ofPieChart = (ExcelOfPieChart)drawings["Chart 1"];
				Assert.AreEqual(eChartType.PieOfPie, ofPieChart.ChartType);
				Assert.AreEqual(ePieType.Pie, ofPieChart.OfPieType);
			}
		}
		#endregion
	}
}