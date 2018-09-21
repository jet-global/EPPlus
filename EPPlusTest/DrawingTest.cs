using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Xml;
using EPPlusTest.Properties;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Slicers;
using OfficeOpenXml.Style;

namespace EPPlusTest
{
	/// <summary>
	/// Unit Tests for various kinds of Excel Drawings. 
	/// </summary>
	[TestClass]
	public class DrawingTest : TestBase
	{
		#region Original EPPlus Comprehensive Drawing Tests (One Test to Test them All)
		[TestMethod]
		public void ComprehensiveDrawingTests()
		{
			BarChart();
			Column();
			Cone();
			Dougnut();
			Drawings();
			Line();
			LineMarker();
			PieChart();
			PieChart3D();
			PieOfChart();
			Pyramid();
			Scatter();
			Bubble();
			Radar();
			Surface();
			Line2Test();
			MultiChartSeries();
			Picture();
			DrawingRowheightDynamic();
			DrawingSizingAndPositioning();
			DeleteDrawing();

			SaveWorksheet("Drawing.xlsx");

			ReadDocument();
			ReadDrawing();
		}

		public void ReadDrawing()
		{
			using (ExcelPackage pck = new ExcelPackage(new FileInfo(_worksheetPath + @"Drawing.xlsx")))
			{
				var ws = pck.Workbook.Worksheets["Pyramid"];
				Assert.AreEqual(ws.Cells["V24"].Value, 104D);
				ws = pck.Workbook.Worksheets["Scatter"];
				var cht = ws.Drawings["ScatterChart1"] as ExcelScatterChart;
				Assert.AreEqual(cht.Title.Text, "Header  Text");
				cht.Title.Text = "Test";
				Assert.AreEqual(cht.Title.Text, "Test");
			}
		}

		public void Picture()
		{
			var ws = base._pck.Workbook.Worksheets.Add("Picture");
			var pic = ws.Drawings.AddPicture("Pic1", Resources.Test1);

			pic = ws.Drawings.AddPicture("Pic2", Resources.Test1);
			pic.SetPosition(150, 200);
			pic.Border.LineStyle = eLineStyle.Solid;
			pic.Border.Fill.Color = Color.DarkCyan;
			pic.Fill.Style = eFillStyle.SolidFill;
			pic.Fill.Color = Color.White;
			pic.Fill.Transparancy = 50;

			pic = ws.Drawings.AddPicture("Pic3", Resources.Test1);
			pic.SetPosition(400, 200);
			pic.SetSize(150);

			pic = ws.Drawings.AddPicture("Pic5", new FileInfo(Path.Combine(_clipartPath, "BitmapImage.gif")));
			pic.SetPosition(400, 200);
			pic.SetSize(150);

			ws.Column(1).Width = 53;
			ws.Column(4).Width = 58;

			pic = ws.Drawings.AddPicture("Pic6öäå", new FileInfo(Path.Combine(_clipartPath, "BitmapImage.gif")));
			pic.SetPosition(400, 400);
			pic.SetSize(100);

			pic = ws.Drawings.AddPicture("PicPixelSized", Resources.Test1);
			pic.SetPosition(800, 800);
			pic.SetSize(568 * 2, 66 * 2);
			var ws2 = base._pck.Workbook.Worksheets.Add("Picture2");
			var fi = new FileInfo(Path.Combine(_clipartPath, "BitmapImage.gif"));
			if (fi.Exists)
			{
				pic = ws2.Drawings.AddPicture("Pic7", fi);
			}
			else
			{
				TestContext.WriteLine("AG00021_.GIF does not exists. Skipping Pic7.");
			}

			var wsCopy = base._pck.Workbook.Worksheets.Add("Picture3", ws2);
		}

		public void DrawingSizingAndPositioning()
		{
			var ws = base._pck.Workbook.Worksheets.Add("DrawingPosSize");

			var pic = ws.Drawings.AddPicture("Pic1", Resources.Test1);
			pic.SetPosition(1, 0, 1, 0);

			pic = ws.Drawings.AddPicture("Pic2", Resources.Test1);
			pic.EditAs = eEditAs.Absolute;
			pic.SetPosition(10, 5, 1, 4);

			pic = ws.Drawings.AddPicture("Pic3", Resources.Test1);
			pic.EditAs = eEditAs.TwoCell;
			pic.SetPosition(20, 5, 2, 4);


			ws.Column(1).Width = 100;
			ws.Column(3).Width = 100;
		}

		public void BarChart()
		{
			var ws = base._pck.Workbook.Worksheets.Add("BarChart");
			var chrt = ws.Drawings.AddChart("barChart", eChartType.BarClustered) as ExcelBarChart;
			chrt.SetPosition(50, 50);
			chrt.SetSize(800, 300);
			AddTestSerie(ws, chrt);
			chrt.VaryColors = true;
			chrt.XAxis.Orientation = eAxisOrientation.MaxMin;
			chrt.XAxis.MajorTickMark = eAxisTickMark.In;
			chrt.XAxis.Format = "yyyy-MM";
			chrt.YAxis.Orientation = eAxisOrientation.MaxMin;
			chrt.YAxis.MinorTickMark = eAxisTickMark.Out;
			chrt.ShowHiddenData = true;
			chrt.DisplayBlanksAs = eDisplayBlanksAs.Zero;
			chrt.Title.RichText.Text = "Barchart Test";
			chrt.GapWidth = 5;
			Assert.IsTrue(chrt.ChartType == eChartType.BarClustered, "Invalid Charttype");
			Assert.IsTrue(chrt.Direction == eDirection.Bar, "Invalid Bardirection");
			Assert.IsTrue(chrt.Grouping == eGrouping.Clustered, "Invalid Grouping");
			Assert.IsTrue(chrt.Shape == eShape.Box, "Invalid Shape");
		}

		public void PieChart()
		{
			var ws = base._pck.Workbook.Worksheets.Add("PieChart");
			var chrt = ws.Drawings.AddChart("pieChart", eChartType.Pie) as ExcelPieChart;

			AddTestSerie(ws, chrt);

			chrt.To.Row = 25;
			chrt.To.Column = 12;

			chrt.DataLabel.ShowPercent = true;
			chrt.Legend.Font.Color = Color.SteelBlue;
			chrt.Title.Border.Fill.Style = eFillStyle.SolidFill;
			chrt.Legend.Position = eLegendPosition.TopRight;
			Assert.IsTrue(chrt.ChartType == eChartType.Pie, "Invalid Charttype");
			Assert.IsTrue(chrt.VaryColors);
			chrt.Title.Text = "Piechart";
		}

		public void PieOfChart()
		{
			var ws = base._pck.Workbook.Worksheets.Add("PieOfChart");
			var chrt = ws.Drawings.AddChart("pieOfChart", eChartType.BarOfPie) as ExcelOfPieChart;

			AddTestSerie(ws, chrt);

			chrt.To.Row = 25;
			chrt.To.Column = 12;

			chrt.DataLabel.ShowPercent = true;
			chrt.Legend.Font.Color = Color.SteelBlue;
			chrt.Title.Border.Fill.Style = eFillStyle.SolidFill;
			chrt.Legend.Position = eLegendPosition.TopRight;
			Assert.IsTrue(chrt.ChartType == eChartType.BarOfPie, "Invalid Charttype");
			chrt.Title.Text = "Piechart";
		}

		public void PieChart3D()
		{
			var ws = base._pck.Workbook.Worksheets.Add("PieChart3d");
			var chrt = ws.Drawings.AddChart("pieChart3d", eChartType.Pie3D) as ExcelPieChart;
			AddTestSerie(ws, chrt);

			chrt.To.Row = 25;
			chrt.To.Column = 12;

			chrt.DataLabel.ShowValue = true;
			chrt.Legend.Position = eLegendPosition.Left;
			chrt.ShowHiddenData = false;
			chrt.DisplayBlanksAs = eDisplayBlanksAs.Gap;
			chrt.Title.RichText.Add("Pie RT Title add");
			Assert.IsTrue(chrt.ChartType == eChartType.Pie3D, "Invalid Charttype");
			Assert.IsTrue(chrt.VaryColors);

		}

		public void Scatter()
		{
			var ws = base._pck.Workbook.Worksheets.Add("Scatter");
			var chrt = ws.Drawings.AddChart("ScatterChart1", eChartType.XYScatterSmoothNoMarkers) as ExcelScatterChart;
			AddTestSerie(ws, chrt);
			chrt.To.Row = 23;
			chrt.To.Column = 12;
			var r1 = chrt.Title.RichText.Add("Header");
			r1.Bold = true;
			var r2 = chrt.Title.RichText.Add("  Text");
			r2.UnderLine = eUnderLineType.WavyHeavy;

			chrt.Title.Fill.Style = eFillStyle.SolidFill;
			chrt.Title.Fill.Color = Color.LightBlue;
			chrt.Title.Fill.Transparancy = 50;
			chrt.VaryColors = true;
			ExcelScatterChartSerie ser = chrt.Series[0] as ExcelScatterChartSerie;
			ser.DataLabel.Position = eLabelPosition.Center;
			ser.DataLabel.ShowValue = true;
			ser.DataLabel.ShowCategory = true;
			ser.DataLabel.Fill.Color = Color.BlueViolet;
			ser.DataLabel.Font.Color = Color.White;
			ser.DataLabel.Font.Italic = true;
			ser.DataLabel.Font.SetFromFont(new Font("bookman old style", 8));
			Assert.IsTrue(chrt.ChartType == eChartType.XYScatterSmoothNoMarkers, "Invalid Charttype");
			chrt.Series[0].Header = "Test serie";
			chrt = ws.Drawings.AddChart("ScatterChart2", eChartType.XYScatterSmooth) as ExcelScatterChart;
			chrt.Series.Add("U19:U24", "V19:V24");

			chrt.From.Column = 0;
			chrt.From.Row = 25;
			chrt.To.Row = 53;
			chrt.To.Column = 12;
			chrt.Legend.Position = eLegendPosition.Bottom;
		}

		public void Bubble()
		{
			var ws = base._pck.Workbook.Worksheets.Add("Bubble");
			var chrt = ws.Drawings.AddChart("Bubble", eChartType.Bubble) as ExcelBubbleChart;
			AddTestData(ws);

			chrt.Series.Add("V19:V24", "U19:U24");

			chrt = ws.Drawings.AddChart("Bubble3d", eChartType.Bubble3DEffect) as ExcelBubbleChart;
			ws.Cells["W19"].Value = 1;
			ws.Cells["W20"].Value = 1;
			ws.Cells["W21"].Value = 2;
			ws.Cells["W22"].Value = 2;
			ws.Cells["W23"].Value = 3;
			ws.Cells["W24"].Value = 4;

			chrt.Series.Add("V19:V24", "U19:U24", "W19:W24");
			chrt.Style = eChartStyle.Style25;

			chrt.From.Row = 23;
			chrt.From.Column = 12;
			chrt.To.Row = 33;
			chrt.To.Column = 22;
			chrt.Title.Text = "Header Text";


		}

		public void Radar()
		{
			var ws = base._pck.Workbook.Worksheets.Add("Radar");
			AddTestData(ws);

			var chrt = ws.Drawings.AddChart("Radar1", eChartType.Radar) as ExcelRadarChart;
			var s = chrt.Series.Add("V19:V24", "U19:U24");
			s.Header = "serie1";
			chrt.From.Row = 23;
			chrt.From.Column = 12;
			chrt.To.Row = 38;
			chrt.To.Column = 22;
			chrt.Title.Text = "Radar Chart 1";

			chrt = ws.Drawings.AddChart("Radar2", eChartType.RadarFilled) as ExcelRadarChart;
			s = chrt.Series.Add("V19:V24", "U19:U24");
			s.Header = "serie1";
			chrt.From.Row = 43;
			chrt.From.Column = 12;
			chrt.To.Row = 58;
			chrt.To.Column = 22;
			chrt.Title.Text = "Radar Chart 2";

			chrt = ws.Drawings.AddChart("Radar3", eChartType.RadarMarkers) as ExcelRadarChart;
			var rs = (ExcelRadarChartSerie)chrt.Series.Add("V19:V24", "U19:U24");
			rs.Header = "serie1";
			rs.Marker = eMarkerStyle.Star;
			rs.MarkerSize = 14;

			chrt.From.Row = 63;
			chrt.From.Column = 12;
			chrt.To.Row = 78;
			chrt.To.Column = 22;
			chrt.Title.Text = "Radar Chart 3";
		}

		public void Surface()
		{
			var ws = base._pck.Workbook.Worksheets.Add("Surface");
			AddTestData(ws);

			var chrt = ws.Drawings.AddChart("Surface1", eChartType.Surface) as ExcelSurfaceChart;
			var s = chrt.Series.Add("V19:V24", "U19:U24");
			var s2 = chrt.Series.Add("W19:W24", "U19:U24");
			s.Header = "serie1";
			chrt.From.Row = 23;
			chrt.From.Column = 12;
			chrt.To.Row = 38;
			chrt.To.Column = 22;
			chrt.Title.Text = "Surface Chart 1";
		}

		public void Pyramid()
		{
			var ws = base._pck.Workbook.Worksheets.Add("Pyramid");
			var chrt = ws.Drawings.AddChart("Pyramid1", eChartType.PyramidCol) as ExcelBarChart;
			AddTestSerie(ws, chrt);
			chrt.VaryColors = true;
			chrt.To.Row = 23;
			chrt.To.Column = 12;
			chrt.Title.Text = "Header Text";
			chrt.Title.Fill.Style = eFillStyle.SolidFill;
			chrt.Title.Fill.Color = Color.DarkBlue;
			chrt.DataLabel.ShowValue = true;
			chrt.Border.LineCap = eLineCap.Round;
			chrt.Border.LineStyle = eLineStyle.LongDashDotDot;
			chrt.Border.Fill.Style = eFillStyle.SolidFill;
			chrt.Border.Fill.Color = Color.Blue;

			chrt.Fill.Color = Color.LightCyan;
			chrt.PlotArea.Fill.Color = Color.White;
			chrt.PlotArea.Border.Fill.Style = eFillStyle.SolidFill;
			chrt.PlotArea.Border.Fill.Color = Color.Beige;
			chrt.PlotArea.Border.LineStyle = eLineStyle.LongDash;

			chrt.Legend.Fill.Color = Color.Aquamarine;
			chrt.Legend.Position = eLegendPosition.Top;
			chrt.Axis[0].Fill.Style = eFillStyle.SolidFill;
			chrt.Axis[0].Fill.Color = Color.Black;
			chrt.Axis[0].Font.Color = Color.White;

			chrt.Axis[1].Fill.Style = eFillStyle.SolidFill;
			chrt.Axis[1].Fill.Color = Color.LightSlateGray;
			chrt.Axis[1].Font.Color = Color.DarkRed;

			chrt.DataLabel.Font.Bold = true;
			chrt.DataLabel.Fill.Color = Color.LightBlue;
			chrt.DataLabel.Border.Fill.Style = eFillStyle.SolidFill;
			chrt.DataLabel.Border.Fill.Color = Color.Black;
			chrt.DataLabel.Border.LineStyle = eLineStyle.Solid;
		}

		public void Cone()
		{
			var ws = base._pck.Workbook.Worksheets.Add("Cone");
			var chrt = ws.Drawings.AddChart("Cone1", eChartType.ConeBarClustered) as ExcelBarChart;
			AddTestSerie(ws, chrt);
			chrt.VaryColors = true;
			chrt.SetSize(200);
			chrt.Title.Text = "Cone bar";
			chrt.Series[0].Header = "Serie 1";
			chrt.Legend.Position = eLegendPosition.Right;
			chrt.Axis[1].DisplayUnit = 100000;
			Assert.AreEqual(chrt.Axis[1].DisplayUnit, 100000);
		}

		public void Column()
		{
			var ws = base._pck.Workbook.Worksheets.Add("Column");
			var chrt = ws.Drawings.AddChart("Column1", eChartType.ColumnClustered3D) as ExcelBarChart;
			AddTestSerie(ws, chrt);
			chrt.VaryColors = true;
			chrt.View3D.RightAngleAxes = true;
			chrt.View3D.DepthPercent = 99;
			chrt.View3D.RightAngleAxes = true;
			chrt.SetSize(200);
			chrt.Title.Text = "Column";
			chrt.Series[0].Header = "Serie 1";
			chrt.Locked = false;
			chrt.Print = false;
			chrt.EditAs = eEditAs.TwoCell;
			chrt.Axis[1].DisplayUnit = 10020;
			Assert.AreEqual(chrt.Axis[1].DisplayUnit, 10020);
		}

		public void Dougnut()
		{
			var ws = base._pck.Workbook.Worksheets.Add("Dougnut");
			var chrt = ws.Drawings.AddChart("Dougnut1", eChartType.DoughnutExploded) as ExcelDoughnutChart;
			AddTestSerie(ws, chrt);
			chrt.SetSize(200);
			chrt.Title.Text = "Doughnut Exploded";
			chrt.Series[0].Header = "Serie 1";
			chrt.EditAs = eEditAs.Absolute;
		}

		public void Line()
		{
			var ws = base._pck.Workbook.Worksheets.Add("Line");
			var chrt = ws.Drawings.AddChart("Line1", eChartType.Line) as ExcelLineChart;
			AddTestSerie(ws, chrt);
			chrt.SetSize(150);
			chrt.VaryColors = true;
			chrt.Smooth = false;
			chrt.Title.Text = "Line 3D";
			chrt.Series[0].Header = "Line serie 1";
			var tl = chrt.Series[0].TrendLines.Add(eTrendLine.Polynomial);
			tl.Name = "Test";
			tl.DisplayRSquaredValue = true;
			tl.DisplayEquation = true;
			tl.Forward = 15;
			tl.Backward = 1;
			tl.Intercept = 6;
			tl.Order = 5;

			tl = chrt.Series[0].TrendLines.Add(eTrendLine.MovingAvgerage);
			chrt.Fill.Color = Color.LightSteelBlue;
			chrt.Border.LineStyle = eLineStyle.Dot;
			chrt.Border.Fill.Color = Color.Black;

			chrt.Legend.Font.Color = Color.Red;
			chrt.Legend.Font.Strike = eStrikeType.Double;
			chrt.Title.Font.Color = Color.DarkGoldenrod;
			chrt.Title.Font.LatinFont = "Arial";
			chrt.Title.Font.Bold = true;
			chrt.Title.Fill.Color = Color.White;
			chrt.Title.Border.Fill.Style = eFillStyle.SolidFill;
			chrt.Title.Border.LineStyle = eLineStyle.LongDashDotDot;
			chrt.Title.Border.Fill.Color = Color.Tomato;
			chrt.DataLabel.ShowSeriesName = true;
			chrt.DataLabel.ShowLeaderLines = true;
			chrt.EditAs = eEditAs.OneCell;
			chrt.DisplayBlanksAs = eDisplayBlanksAs.Span;
			chrt.Axis[0].Title.Text = "Axis 0";
			chrt.Axis[0].Title.Rotation = 90;
			chrt.Axis[0].Title.Overlay = true;
			chrt.Axis[1].Title.Text = "Axis 1";
			chrt.Axis[1].Title.AnchorCtr = true;
			chrt.Axis[1].Title.TextVertical = eTextVerticalType.Vertical270;
			chrt.Axis[1].Title.Border.LineStyle = eLineStyle.LongDashDotDot;

		}

		public void LineMarker()
		{
			var ws = base._pck.Workbook.Worksheets.Add("LineMarker1");
			var chrt = ws.Drawings.AddChart("Line1", eChartType.LineMarkers) as ExcelLineChart;
			AddTestSerie(ws, chrt);
			chrt.SetSize(150);
			chrt.Title.Text = "Line Markers";
			chrt.Series[0].Header = "Line serie 1";
			((ExcelLineChartSerie)chrt.Series[0]).Marker = eMarkerStyle.Plus;

			var chrt2 = ws.Drawings.AddChart("Line2", eChartType.LineMarkers) as ExcelLineChart;
			AddTestSerie(ws, chrt2);
			chrt2.SetPosition(500, 0);
			chrt2.SetSize(150);
			chrt2.Title.Text = "Line Markers";
			var serie = (ExcelLineChartSerie)chrt2.Series[0];
			serie.Marker = eMarkerStyle.X;

		}

		public void Drawings()
		{
			var ws = base._pck.Workbook.Worksheets.Add("Shapes");
			int y = 100, i = 1;
			foreach (eShapeStyle style in Enum.GetValues(typeof(eShapeStyle)))
			{
				var shape = ws.Drawings.AddShape("shape" + i.ToString(), style);
				shape.SetPosition(y, 100);
				shape.SetSize(300, 300);
				y += 400;
				shape.Text = style.ToString();
				i++;
			}

			(ws.Drawings["shape1"] as ExcelShape).TextAnchoring = eTextAnchoringType.Top;
			var rt = (ws.Drawings["shape1"] as ExcelShape).RichText.Add("Added formated richtext");
			(ws.Drawings["shape1"] as ExcelShape).LockText = false;
			rt.Bold = true;
			rt.Color = Color.Aquamarine;
			rt.Italic = true;
			rt.Size = 17;
			(ws.Drawings["shape2"] as ExcelShape).TextVertical = eTextVerticalType.Vertical;
			rt = (ws.Drawings["shape2"] as ExcelShape).RichText.Add("\r\nAdded formated richtext");
			rt.Bold = true;
			rt.Color = Color.DarkGoldenrod;
			rt.SetFromFont(new Font("Times new roman", 18, FontStyle.Underline));
			rt.UnderLineColor = Color.Green;


			(ws.Drawings["shape3"] as ExcelShape).TextAnchoring = eTextAnchoringType.Bottom;
			(ws.Drawings["shape3"] as ExcelShape).TextAnchoringControl = true;

			(ws.Drawings["shape4"] as ExcelShape).TextVertical = eTextVerticalType.Vertical270;
			(ws.Drawings["shape4"] as ExcelShape).TextAnchoring = eTextAnchoringType.Top;

			(ws.Drawings["shape5"] as ExcelShape).Fill.Style = eFillStyle.SolidFill;
			(ws.Drawings["shape5"] as ExcelShape).Fill.Color = Color.Red;
			(ws.Drawings["shape5"] as ExcelShape).Fill.Transparancy = 50;

			(ws.Drawings["shape6"] as ExcelShape).Fill.Style = eFillStyle.NoFill;
			(ws.Drawings["shape6"] as ExcelShape).Font.Color = Color.Black;
			(ws.Drawings["shape6"] as ExcelShape).Border.Fill.Color = Color.Black;

			(ws.Drawings["shape7"] as ExcelShape).Fill.Style = eFillStyle.SolidFill;
			(ws.Drawings["shape7"] as ExcelShape).Fill.Color = Color.Gray;
			(ws.Drawings["shape7"] as ExcelShape).Border.Fill.Style = eFillStyle.SolidFill;
			(ws.Drawings["shape7"] as ExcelShape).Border.Fill.Color = Color.Black;
			(ws.Drawings["shape7"] as ExcelShape).Border.Fill.Transparancy = 43;
			(ws.Drawings["shape7"] as ExcelShape).Border.LineCap = eLineCap.Round;
			(ws.Drawings["shape7"] as ExcelShape).Border.LineStyle = eLineStyle.LongDash;
			(ws.Drawings["shape7"] as ExcelShape).Font.UnderLineColor = Color.Blue;
			(ws.Drawings["shape7"] as ExcelShape).Font.Color = Color.Black;
			(ws.Drawings["shape7"] as ExcelShape).Font.Bold = true;
			(ws.Drawings["shape7"] as ExcelShape).Font.LatinFont = "Arial";
			(ws.Drawings["shape7"] as ExcelShape).Font.ComplexFont = "Arial";
			(ws.Drawings["shape7"] as ExcelShape).Font.Italic = true;
			(ws.Drawings["shape7"] as ExcelShape).Font.UnderLine = eUnderLineType.Dotted;

			(ws.Drawings["shape8"] as ExcelShape).Fill.Style = eFillStyle.SolidFill;
			(ws.Drawings["shape8"] as ExcelShape).Font.LatinFont = "Miriam";
			(ws.Drawings["shape8"] as ExcelShape).Font.UnderLineColor = Color.CadetBlue;
			(ws.Drawings["shape8"] as ExcelShape).Font.UnderLine = eUnderLineType.Single;

			(ws.Drawings["shape9"] as ExcelShape).TextAlignment = eTextAlignment.Right;

			(ws.Drawings["shape120"] as ExcelShape).LineEnds.TailEnd = eEndStyle.Oval;
			(ws.Drawings["shape120"] as ExcelShape).LineEnds.TailEndSizeWidth = eEndSize.Large;
			(ws.Drawings["shape120"] as ExcelShape).LineEnds.TailEndSizeHeight = eEndSize.Large;
			(ws.Drawings["shape120"] as ExcelShape).LineEnds.HeadEnd = eEndStyle.Arrow;
			(ws.Drawings["shape120"] as ExcelShape).LineEnds.HeadEndSizeHeight = eEndSize.Small;
			(ws.Drawings["shape120"] as ExcelShape).LineEnds.HeadEndSizeWidth = eEndSize.Small;
		}

		public void Line2Test()
		{
			ExcelWorksheet worksheet = base._pck.Workbook.Worksheets.Add("LineIssue");

			ExcelChart chart = worksheet.Drawings.AddChart("LineChart", eChartType.Line);

			worksheet.Cells["A1"].Value = 1;
			worksheet.Cells["A2"].Value = 2;
			worksheet.Cells["A3"].Value = 3;
			worksheet.Cells["A4"].Value = 4;
			worksheet.Cells["A5"].Value = 5;
			worksheet.Cells["A6"].Value = 6;

			worksheet.Cells["B1"].Value = 10000;
			worksheet.Cells["B2"].Value = 10100;
			worksheet.Cells["B3"].Value = 10200;
			worksheet.Cells["B4"].Value = 10150;
			worksheet.Cells["B5"].Value = 10250;
			worksheet.Cells["B6"].Value = 10200;

			chart.Series.Add(ExcelRange.GetAddress(1, 2, worksheet.Dimension.End.Row, 2),
							 ExcelRange.GetAddress(1, 1, worksheet.Dimension.End.Row, 1));

			var Series = chart.Series[0];

			chart.Series[0].Header = "Blah";
		}

		public void MultiChartSeries()
		{
			ExcelWorksheet worksheet = base._pck.Workbook.Worksheets.Add("MultiChartTypes");

			ExcelChart chart = worksheet.Drawings.AddChart("chtPie", eChartType.LineMarkers);
			chart.SetPosition(100, 100);
			chart.SetSize(800, 600);
			AddTestSerie(worksheet, chart);
			chart.Series[0].Header = "Serie5";
			chart.Style = eChartStyle.Style27;
			worksheet.Cells["W19"].Value = 120;
			worksheet.Cells["W20"].Value = 122;
			worksheet.Cells["W21"].Value = 121;
			worksheet.Cells["W22"].Value = 123;
			worksheet.Cells["W23"].Value = 125;
			worksheet.Cells["W24"].Value = 124;

			worksheet.Cells["X19"].Value = 90;
			worksheet.Cells["X20"].Value = 52;
			worksheet.Cells["X21"].Value = 88;
			worksheet.Cells["X22"].Value = 75;
			worksheet.Cells["X23"].Value = 77;
			worksheet.Cells["X24"].Value = 99;

			var cs2 = chart.PlotArea.ChartTypes.Add(eChartType.ColumnClustered);
			var s = cs2.Series.Add(worksheet.Cells["W19:W24"], worksheet.Cells["U19:U24"]);
			s.Header = "Serie4";
			cs2.YAxis.MaxValue = 300;
			cs2.YAxis.MinValue = -5.5;
			var cs3 = chart.PlotArea.ChartTypes.Add(eChartType.Line);
			s = cs3.Series.Add(worksheet.Cells["X19:X24"], worksheet.Cells["U19:U24"]);
			s.Header = "Serie1";
			cs3.UseSecondaryAxis = true;

			cs3.XAxis.Deleted = false;
			cs3.XAxis.MajorUnit = 20;
			cs3.XAxis.MinorUnit = 3;

			cs3.XAxis.TickLabelPosition = eTickLabelPosition.High;
			cs3.YAxis.LogBase = 10.2;

			var chart2 = worksheet.Drawings.AddChart("scatter1", eChartType.XYScatterSmooth);
			s = chart2.Series.Add(worksheet.Cells["W19:W24"], worksheet.Cells["U19:U24"]);
			s.Header = "Serie2";

			var c2ct2 = chart2.PlotArea.ChartTypes.Add(eChartType.XYScatterSmooth);
			s = c2ct2.Series.Add(worksheet.Cells["X19:X24"], worksheet.Cells["V19:V24"]);
			s.Header = "Serie3";
			s = c2ct2.Series.Add(worksheet.Cells["W19:W24"], worksheet.Cells["V19:V24"]);
			s.Header = "Serie4";

			c2ct2.UseSecondaryAxis = true;
			c2ct2.XAxis.Deleted = false;
			c2ct2.XAxis.TickLabelPosition = eTickLabelPosition.High;

			ExcelChart chart3 = worksheet.Drawings.AddChart("chart", eChartType.LineMarkers);
			chart3.SetPosition(300, 1000);
			var s31 = chart3.Series.Add(worksheet.Cells["W19:W24"], worksheet.Cells["U19:U24"]);
			s31.Header = "Serie1";

			var c3ct2 = chart3.PlotArea.ChartTypes.Add(eChartType.LineMarkers);
			var c32 = c3ct2.Series.Add(worksheet.Cells["X19:X24"], worksheet.Cells["V19:V24"]);
			c3ct2.UseSecondaryAxis = true;
			c32.Header = "Serie2";

			XmlNamespaceManager ns = new XmlNamespaceManager(new NameTable());
			ns.AddNamespace("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
			var element = chart.ChartXml.SelectSingleNode("//c:plotVisOnly", ns);
			if (element != null) element.ParentNode.RemoveChild(element);
		}

		public void DeleteDrawing()
		{
			var ws = base._pck.Workbook.Worksheets.Add("DeleteDrawing1");
			var chart1 = ws.Drawings.AddChart("Chart1", eChartType.Line);
			var chart2 = ws.Drawings.AddChart("Chart2", eChartType.Line);
			var shape1 = ws.Drawings.AddShape("Shape1", eShapeStyle.ActionButtonBackPrevious);
			var pic1 = ws.Drawings.AddPicture("Pic1", Resources.Test1);
			ws.Drawings.Remove(2);
			ws.Drawings.Remove(chart2);
			ws.Drawings.Remove("Pic1");

			ws = base._pck.Workbook.Worksheets.Add("DeleteDrawing2");
			chart1 = ws.Drawings.AddChart("Chart1", eChartType.Line);
			chart2 = ws.Drawings.AddChart("Chart2", eChartType.Line);
			shape1 = ws.Drawings.AddShape("Shape1", eShapeStyle.ActionButtonBackPrevious);
			pic1 = ws.Drawings.AddPicture("Pic1", Resources.Test1);

			ws.Drawings.Remove("chart1");

			ws = base._pck.Workbook.Worksheets.Add("ClearDrawing2");
			chart1 = ws.Drawings.AddChart("Chart1", eChartType.Line);
			chart2 = ws.Drawings.AddChart("Chart2", eChartType.Line);
			shape1 = ws.Drawings.AddShape("Shape1", eShapeStyle.ActionButtonBackPrevious);
			pic1 = ws.Drawings.AddPicture("Pic1", Resources.Test1);
			ws.Drawings.Clear();
		}

		public void ReadDocument()
		{
			var fi = new FileInfo(_worksheetPath + "drawing.xlsx");
			if (!fi.Exists)
			{
				Assert.Inconclusive("Drawing.xlsx is not created. Skippng");
			}
			var pck = new ExcelPackage(fi, true);

			foreach (var ws in pck.Workbook.Worksheets)
			{
				foreach (ExcelDrawing d in pck.Workbook.Worksheets[1].Drawings)
				{
					if (d is ExcelChart)
					{
						TestContext.WriteLine(((ExcelChart)d).ChartType.ToString());
					}
				}
			}
			pck.Dispose();
		}

		public void DrawingRowheightDynamic()
		{
			var ws = base._pck.Workbook.Worksheets.Add("PicResize");
			ws.Cells["A1"].Value = "test";
			ws.Cells["A1"].Style.Font.Name = "Symbol";
			ws.Cells["A1"].Style.Font.Size = 39;
			ws.Workbook.Styles.NamedStyles[0].Style.Font.Name = "Symbol";
			ws.Workbook.Styles.NamedStyles[0].Style.Font.Size = 16;
			var pic = ws.Drawings.AddPicture("Pic1", Resources.Test1);
			pic.SetPosition(10, 12);
		}
		#endregion

		#region Additional Drawings Tests
		[TestMethod]
		public void GetDrawingByNameGetsCorrectDrawing()
		{
			var worksheet = base._pck.Workbook.Worksheets.Add("sheet");
			var lineChartDrawing = worksheet.Drawings.AddChart("LineChart", eChartType.Line);
			var hexagonDrawing = worksheet.Drawings.AddShape("Hexagon", eShapeStyle.Hexagon);
			var pictureDrawing = worksheet.Drawings.AddPicture("Picture", Resources.Test1);

			Assert.AreEqual(3, worksheet.Drawings.Count);
			ExcelDrawing retrievedDrawing = worksheet.Drawings["LineChart"];
			Assert.AreEqual(lineChartDrawing, retrievedDrawing);
			retrievedDrawing = worksheet.Drawings["Hexagon"];
			Assert.AreEqual(hexagonDrawing, retrievedDrawing);
			retrievedDrawing = worksheet.Drawings["Picture"];
			Assert.AreEqual(pictureDrawing, retrievedDrawing);
			Assert.AreEqual(3, worksheet.Drawings.Count);
			retrievedDrawing = worksheet.Drawings["NonExistent Drawing"];
			Assert.AreEqual(null, retrievedDrawing);
		}

		[TestMethod]
		public void GetDrawingFromEmptyWorkbook()
		{
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("sheet");
				Assert.AreEqual(0, worksheet.Drawings.Count);
				var retrievedDrawing = worksheet.Drawings["NonExistent Drawing"];
				Assert.AreEqual(null, retrievedDrawing);
			}
		}

		[TestMethod]
		public void MultipleDrawingsWithTheSameNameGetsCorrectDrawing()
		{
			var worksheet = base._pck.Workbook.Worksheets.Add("sheet");
			var picture1 = worksheet.Drawings.AddPicture("drawing", Resources.Test1);
			Assert.AreEqual(picture1, worksheet.Drawings["drawing"]);
			var picture2 = worksheet.Drawings.AddPicture("drawing", Resources.Test1);
			Assert.AreEqual(picture1, worksheet.Drawings["drawing"]);
			var shape1 = worksheet.Drawings.AddShape("drawing", eShapeStyle.Hexagon);
			Assert.AreEqual(picture1, worksheet.Drawings["drawing"]);
			var shape2 = worksheet.Drawings.AddShape("drawing", eShapeStyle.Hexagon);
			Assert.AreEqual(picture1, worksheet.Drawings["drawing"]);
			worksheet.Drawings.Remove(picture1);
			Assert.AreEqual(picture2, worksheet.Drawings["drawing"]);
			worksheet.Drawings.Remove(picture2);
			Assert.AreEqual(shape1, worksheet.Drawings["drawing"]);
			worksheet.Drawings.Remove(shape1);
			Assert.AreEqual(shape2, worksheet.Drawings["drawing"]);
			worksheet.Drawings.Remove(shape2);
			Assert.AreEqual(null, worksheet.Drawings["drawing"]);
		}

		[TestMethod]
		public void RemoveDrawings()
		{
			var worksheet = base._pck.Workbook.Worksheets.Add("sheet");
			var lineChartDrawing = worksheet.Drawings.AddChart("LineChart", eChartType.Line);
			var hexagonDrawing = worksheet.Drawings.AddShape("Hexagon", eShapeStyle.Hexagon);
			var pictureDrawing = worksheet.Drawings.AddPicture("Picture", Resources.Test1);

			Assert.AreEqual(3, worksheet.Drawings.Count);
			Assert.IsTrue(worksheet.Drawings.Contains(pictureDrawing));
			worksheet.Drawings.Remove(2);
			Assert.IsFalse(worksheet.Drawings.Contains(pictureDrawing));

			Assert.IsTrue(worksheet.Drawings.Contains(hexagonDrawing));
			worksheet.Drawings.Remove(hexagonDrawing);
			Assert.IsFalse(worksheet.Drawings.Contains(hexagonDrawing));

			Assert.IsTrue(worksheet.Drawings.Contains(lineChartDrawing));
			worksheet.Drawings.Remove("LineChart");
			Assert.IsFalse(worksheet.Drawings.Contains(lineChartDrawing));
			Assert.AreEqual(0, worksheet.Drawings.Count);
		}

		[TestMethod]
		public void ClearDrawings()
		{
			var worksheet = base._pck.Workbook.Worksheets.Add("sheet");
			var lineChartDrawing = worksheet.Drawings.AddChart("LineChart", eChartType.Line);
			var hexagonDrawing = worksheet.Drawings.AddShape("Hexagon", eShapeStyle.Hexagon);
			var pictureDrawing = worksheet.Drawings.AddPicture("Picture", Resources.Test1);
			Assert.AreEqual(3, worksheet.Drawings.Count);
			worksheet.Drawings.ClearDrawings();
			Assert.AreEqual(0, worksheet.Drawings.Count);
		}

		[TestMethod]
		public void SaveCloseAndLoadDrawingsIntoWorkbook()
		{
			string sheetName = "DrawingSheet";
			FileInfo fileInfo = new FileInfo(base._worksheetPath + "SaveAndLoadTest.xlsx");
			if (fileInfo.Exists)
				fileInfo.Delete();
			using (var package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add(sheetName);
				Assert.AreEqual(0, worksheet.Drawings.Count);
				package.SaveAs(fileInfo);
			}

			using (var package = new ExcelPackage(fileInfo))
			{
				var worksheet = package.Workbook.Worksheets[sheetName];
				Assert.AreEqual(0, worksheet.Drawings.Count);
				worksheet.Drawings.AddChart("LineChart", eChartType.Line);
				worksheet.Drawings.AddShape("Hexagon", eShapeStyle.Hexagon);
				worksheet.Drawings.AddPicture("Picture", Resources.Test1);
				Assert.AreEqual(3, worksheet.Drawings.Count);
				package.Save();
			}

			using (var package = new ExcelPackage(fileInfo))
			{
				var worksheet = package.Workbook.Worksheets[sheetName];
				Assert.IsTrue(worksheet.Drawings.Contains(worksheet.Drawings["LineChart"]));
				Assert.IsTrue(worksheet.Drawings.Contains(worksheet.Drawings["Hexagon"]));
				Assert.IsTrue(worksheet.Drawings.Contains(worksheet.Drawings["Picture"]));
				worksheet.Drawings.Remove(worksheet.Drawings["LineChart"]);
				worksheet.Drawings.Remove(worksheet.Drawings["Hexagon"]);
				worksheet.Drawings.Remove(worksheet.Drawings["Picture"]);
				Assert.AreEqual(0, worksheet.Drawings.Count);
				package.Save();
			}

			using (var package = new ExcelPackage(fileInfo))
			{
				var worksheet = package.Workbook.Worksheets[sheetName];
				Assert.AreEqual(0, worksheet.Drawings.Count);
				package.Save();
			}

			if (fileInfo.Exists)
				fileInfo.Delete();
		}

		[TestMethod]
		public void SaveCloseAndLoadDrawingsIntoWorkbookWithABMPImage()
		{
			string sheetName = "DrawingSheet";
			FileInfo fileInfo = new FileInfo(base._worksheetPath + "SaveAndLoadTest.xlsx");
			if (fileInfo.Exists)
				fileInfo.Delete();
			try
			{ 
				using (var package = new ExcelPackage())
				{
					var worksheet = package.Workbook.Worksheets.Add(sheetName);
					Assert.AreEqual(0, worksheet.Drawings.Count);
					package.SaveAs(fileInfo);
				}

				using (var package = new ExcelPackage(fileInfo))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					Assert.AreEqual(0, worksheet.Drawings.Count);
					worksheet.Drawings.AddChart("LineChart", eChartType.Line);
					worksheet.Drawings.AddShape("Hexagon", eShapeStyle.Hexagon);
					worksheet.Drawings.AddPicture("Picture", Resources.bmpTestResource);
					Assert.AreEqual(3, worksheet.Drawings.Count);
					package.Save();
				}

				using (var package = new ExcelPackage(fileInfo))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					Assert.IsTrue(worksheet.Drawings.Contains(worksheet.Drawings["LineChart"]));
					Assert.IsTrue(worksheet.Drawings.Contains(worksheet.Drawings["Hexagon"]));
					Assert.IsTrue(worksheet.Drawings.Contains(worksheet.Drawings["Picture"]));
					worksheet.Drawings.Remove(worksheet.Drawings["LineChart"]);
					worksheet.Drawings.Remove(worksheet.Drawings["Hexagon"]);
					worksheet.Drawings.Remove(worksheet.Drawings["Picture"]);
					Assert.AreEqual(0, worksheet.Drawings.Count);
					package.Save();
				}

				using (var package = new ExcelPackage(fileInfo))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					Assert.AreEqual(0, worksheet.Drawings.Count);
					package.Save();
				}

			}
			finally
			{
				if (fileInfo.Exists)
					fileInfo.Delete();
			}
		}

		[TestMethod]
		public void SaveCloseAndLoadDrawingsIntoWorkbookWithAGIFImage()
		{
			string sheetName = "DrawingSheet";
			FileInfo fileInfo = new FileInfo(base._worksheetPath + "SaveAndLoadTest.xlsx");
			if (fileInfo.Exists)
				fileInfo.Delete();
			try
			{ 
				using (var package = new ExcelPackage())
				{
					var worksheet = package.Workbook.Worksheets.Add(sheetName);
					Assert.AreEqual(0, worksheet.Drawings.Count);
					package.SaveAs(fileInfo);
				}

				using (var package = new ExcelPackage(fileInfo))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					Assert.AreEqual(0, worksheet.Drawings.Count);
					worksheet.Drawings.AddChart("LineChart", eChartType.Line);
					worksheet.Drawings.AddShape("Hexagon", eShapeStyle.Hexagon);
					worksheet.Drawings.AddPicture("Picture", Resources.gifTestResource);
					Assert.AreEqual(3, worksheet.Drawings.Count);
					package.Save();
				}

				using (var package = new ExcelPackage(fileInfo))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					Assert.IsTrue(worksheet.Drawings.Contains(worksheet.Drawings["LineChart"]));
					Assert.IsTrue(worksheet.Drawings.Contains(worksheet.Drawings["Hexagon"]));
					Assert.IsTrue(worksheet.Drawings.Contains(worksheet.Drawings["Picture"]));
					worksheet.Drawings.Remove(worksheet.Drawings["LineChart"]);
					worksheet.Drawings.Remove(worksheet.Drawings["Hexagon"]);
					worksheet.Drawings.Remove(worksheet.Drawings["Picture"]);
					Assert.AreEqual(0, worksheet.Drawings.Count);
					package.Save();
				}

				using (var package = new ExcelPackage(fileInfo))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					Assert.AreEqual(0, worksheet.Drawings.Count);
					package.Save();
				}
			}
			finally
			{
				if (fileInfo.Exists)
					fileInfo.Delete();
			}
		}

[TestMethod]
		public void SaveCloseAndLoadDrawingsIntoWorkbookWithAJPEGImage()
		{
			string sheetName = "DrawingSheet";
			FileInfo fileInfo = new FileInfo(base._worksheetPath + "SaveAndLoadTest.xlsx");
			if (fileInfo.Exists)
				fileInfo.Delete();
			try { 
				using (var package = new ExcelPackage())
				{
					var worksheet = package.Workbook.Worksheets.Add(sheetName);
					Assert.AreEqual(0, worksheet.Drawings.Count);
					package.SaveAs(fileInfo);
				}

				using (var package = new ExcelPackage(fileInfo))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					Assert.AreEqual(0, worksheet.Drawings.Count);
					worksheet.Drawings.AddChart("LineChart", eChartType.Line);
					worksheet.Drawings.AddShape("Hexagon", eShapeStyle.Hexagon);
					worksheet.Drawings.AddPicture("Picture", Resources.jpegTestResource);
					Assert.AreEqual(3, worksheet.Drawings.Count);
					package.Save();
				}

				using (var package = new ExcelPackage(fileInfo))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					Assert.IsTrue(worksheet.Drawings.Contains(worksheet.Drawings["LineChart"]));
					Assert.IsTrue(worksheet.Drawings.Contains(worksheet.Drawings["Hexagon"]));
					Assert.IsTrue(worksheet.Drawings.Contains(worksheet.Drawings["Picture"]));
					worksheet.Drawings.Remove(worksheet.Drawings["LineChart"]);
					worksheet.Drawings.Remove(worksheet.Drawings["Hexagon"]);
					worksheet.Drawings.Remove(worksheet.Drawings["Picture"]);
					Assert.AreEqual(0, worksheet.Drawings.Count);
					package.Save();
				}

				using (var package = new ExcelPackage(fileInfo))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					Assert.AreEqual(0, worksheet.Drawings.Count);
					package.Save();
				}

			}
			finally
			{
				if (fileInfo.Exists)
					fileInfo.Delete();
			}
		}

		[TestMethod]
		public void SaveCloseAndLoadDrawingsIntoWorkbookWithAPNGImage()
		{
			string sheetName = "DrawingSheet";
			FileInfo fileInfo = new FileInfo(base._worksheetPath + "SaveAndLoadTest.xlsx");
			if (fileInfo.Exists)
				fileInfo.Delete();
			try
			{
				using (var package = new ExcelPackage())
				{
					var worksheet = package.Workbook.Worksheets.Add(sheetName);
					Assert.AreEqual(0, worksheet.Drawings.Count);
					package.SaveAs(fileInfo);
				}

				using (var package = new ExcelPackage(fileInfo))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					Assert.AreEqual(0, worksheet.Drawings.Count);
					worksheet.Drawings.AddChart("LineChart", eChartType.Line);
					worksheet.Drawings.AddShape("Hexagon", eShapeStyle.Hexagon);
					worksheet.Drawings.AddPicture("Picture", Resources.pngTestResource);
					Assert.AreEqual(3, worksheet.Drawings.Count);
					package.Save();
				}

				using (var package = new ExcelPackage(fileInfo))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					Assert.IsTrue(worksheet.Drawings.Contains(worksheet.Drawings["LineChart"]));
					Assert.IsTrue(worksheet.Drawings.Contains(worksheet.Drawings["Hexagon"]));
					Assert.IsTrue(worksheet.Drawings.Contains(worksheet.Drawings["Picture"]));
					worksheet.Drawings.Remove(worksheet.Drawings["LineChart"]);
					worksheet.Drawings.Remove(worksheet.Drawings["Hexagon"]);
					worksheet.Drawings.Remove(worksheet.Drawings["Picture"]);
					Assert.AreEqual(0, worksheet.Drawings.Count);
					package.Save();
				}

				using (var package = new ExcelPackage(fileInfo))
				{
					var worksheet = package.Workbook.Worksheets[sheetName];
					Assert.AreEqual(0, worksheet.Drawings.Count);
					package.Save();
				}
			}
			finally
			{
				if (fileInfo.Exists)
					fileInfo.Delete();
			}
		}

		[TestMethod]
		public void AddPicturePNGAndTestItsLocation()
		{
			string sheetName = "DrawingSheet";
			FileInfo fileInfo = new FileInfo(base._worksheetPath + "SaveAndLoadTest.xlsx");
			if (fileInfo.Exists)
				fileInfo.Delete();
			try
			{
				using (var package = new ExcelPackage(fileInfo))
				{
					var worksheet = package.Workbook.Worksheets.Add(sheetName);
					var test = worksheet.Drawings.AddPicture("Picture", Resources.pngTestResource);
					test.SetPosition(4, 0, 7, 0);
					Assert.AreEqual(4, test.From.Row);
					Assert.AreEqual(7, test.From.Column);
					package.Save();
				}

				using (var package = new ExcelPackage(fileInfo))
				{
					var sheet = package.Workbook.Worksheets[sheetName];
					var picture = sheet.Drawings["Picture"];
					Assert.AreEqual(4, picture.From.Row);
					Assert.AreEqual(7, picture.From.Column);
				}
			}
			finally
			{
				if (fileInfo.Exists)
					fileInfo.Delete();
			}
		}

		[TestMethod]
		public void AddPictureJpegAndTestItsLocation()
		{
			string sheetName = "DrawingSheet";
			FileInfo fileInfo = new FileInfo(base._worksheetPath + "SaveAndLoadTest.xlsx");
			if (fileInfo.Exists)
				fileInfo.Delete();
			try
			{
				using (var package = new ExcelPackage(fileInfo))
				{
					var worksheet = package.Workbook.Worksheets.Add(sheetName);
					var test = worksheet.Drawings.AddPicture("Picture", Resources.jpegTestResource);
					test.SetPosition(4, 0, 7, 0);
					Assert.AreEqual(4, test.From.Row);
					Assert.AreEqual(7, test.From.Column);
					package.Save();
				}

				using (var package = new ExcelPackage(fileInfo))
				{
					var sheet = package.Workbook.Worksheets[sheetName];
					var picture = sheet.Drawings["Picture"];
					Assert.AreEqual(4, picture.From.Row);
					Assert.AreEqual(7, picture.From.Column);
				}
			}
			finally
			{
				if (fileInfo.Exists)
					fileInfo.Delete();
			}
		}

		[TestMethod]
		public void AddPictureGIFAndTestItsLocation()
		{
			string sheetName = "DrawingSheet";
			FileInfo fileInfo = new FileInfo(base._worksheetPath + "SaveAndLoadTest.xlsx");
			if (fileInfo.Exists)
				fileInfo.Delete();
			try
			{
				using (var package = new ExcelPackage(fileInfo))
				{
					var worksheet = package.Workbook.Worksheets.Add(sheetName);
					var test = worksheet.Drawings.AddPicture("Picture", Resources.gifTestResource);
					test.SetPosition(4, 0, 7, 0);
					Assert.AreEqual(4, test.From.Row);
					Assert.AreEqual(7, test.From.Column);
					package.Save();
				}

				using (var package = new ExcelPackage(fileInfo))
				{
					var sheet = package.Workbook.Worksheets[sheetName];
					var picture = sheet.Drawings["Picture"];
					Assert.AreEqual(4, picture.From.Row);
					Assert.AreEqual(7, picture.From.Column);
				}
			}
			finally
			{
				if (fileInfo.Exists)
					fileInfo.Delete();
			}
		}

		[TestMethod]
		public void AddPictureBMPAndTestItsLocation()
		{
			string sheetName = "DrawingSheet";
			FileInfo fileInfo = new FileInfo(base._worksheetPath + "SaveAndLoadTest.xlsx");
			if (fileInfo.Exists)
				fileInfo.Delete();
			try
			{
				using (var package = new ExcelPackage(fileInfo))
				{
					var worksheet = package.Workbook.Worksheets.Add(sheetName);
					var test = worksheet.Drawings.AddPicture("Picture", Resources.bmpTestResource);
					test.SetPosition(4, 0, 7, 0);
					Assert.AreEqual(4, test.From.Row);
					Assert.AreEqual(7, test.From.Column);
					package.Save();
				}

				using (var package = new ExcelPackage(fileInfo))
				{
					var sheet = package.Workbook.Worksheets[sheetName];
					var picture = sheet.Drawings["Picture"];
					Assert.AreEqual(4, picture.From.Row);
					Assert.AreEqual(7, picture.From.Column);
				}
			}
			finally
			{
				if (fileInfo.Exists)
					fileInfo.Delete();
			}
		}

		[TestMethod]
		public void AddPictureBMPAndTestItsLocationOneCellAnchor()
		{
			string sheetName = "DrawingSheet";
			FileInfo fileInfo = new FileInfo(base._worksheetPath + "SaveAndLoadTest.xlsx");
			if (fileInfo.Exists)
				fileInfo.Delete();
			try
			{
				using (var package = new ExcelPackage(fileInfo))
				{
					var worksheet = package.Workbook.Worksheets.Add(sheetName);
					var test = worksheet.Drawings.AddPicture("Picture", Resources.bmpTestResource);
					test.EditAs = eEditAs.OneCell;
					test.SetPosition(4, 0, 7, 0);
					test.SetSize(100, 500);
					Assert.AreEqual(4, test.From.Row);
					Assert.AreEqual(7, test.From.Column);
					Assert.AreEqual(100, test.GetPixelWidth());
					Assert.AreEqual(500, test.GetPixelHeight());
					package.Save();
				}

				using (var package = new ExcelPackage(fileInfo))
				{
					var sheet = package.Workbook.Worksheets[sheetName];
					var picture = sheet.Drawings["Picture"];
					Assert.AreEqual(4, picture.From.Row);
					Assert.AreEqual(7, picture.From.Column);
					var PixelWidth2 = picture.GetPixelWidth();
					var PixelHeight2 = picture.GetPixelHeight();
					Assert.AreEqual(100, picture.GetPixelWidth());
					Assert.AreEqual(500, picture.GetPixelHeight());
				}
			}
			finally
			{
				if (fileInfo.Exists)
					fileInfo.Delete();
			}
		}

		[TestMethod]
		public void AddPictureBMPAndTestItsLocationTwoCellAnchor()
		{
			string sheetName = "DrawingSheet";
			FileInfo fileInfo = new FileInfo(base._worksheetPath + "SaveAndLoadTest.xlsx");
			if (fileInfo.Exists)
				fileInfo.Delete();
			try
			{
				using (var package = new ExcelPackage(fileInfo))
				{
					var worksheet = package.Workbook.Worksheets.Add(sheetName);
					var test = worksheet.Drawings.AddPicture("Picture", Resources.bmpTestResource);
					test.SetPosition(4, 0, 7, 0);
					Assert.AreEqual(4, test.From.Row);
					Assert.AreEqual(7, test.From.Column);
					test.EditAs = eEditAs.TwoCell;
					package.Save();
				}

				using (var package = new ExcelPackage(fileInfo))
				{
					var sheet = package.Workbook.Worksheets[sheetName];
					var picture = sheet.Drawings["Picture"];
					Assert.AreEqual(4, picture.From.Row);
					Assert.AreEqual(7, picture.From.Column);
				}
			}
			finally
			{
				if (fileInfo.Exists)
					fileInfo.Delete();
			}
		}

		[TestMethod]
		public void AddPictureBMPAndTestItsLocationAbsoluteCellAnchor()
		{
			string sheetName = "DrawingSheet";
			FileInfo fileInfo = new FileInfo(base._worksheetPath + "SaveAndLoadTest.xlsx");
			if (fileInfo.Exists)
				fileInfo.Delete();
			try
			{
				using (var package = new ExcelPackage(fileInfo))
				{
					var worksheet = package.Workbook.Worksheets.Add(sheetName);
					var test = worksheet.Drawings.AddPicture("Picture", Resources.bmpTestResource);
					test.EditAs = eEditAs.Absolute;
					test.SetPixelLeft(450);
					test.SetPixelTop(90);
					Assert.AreEqual(4, test.From.Row);
					Assert.AreEqual(7, test.From.Column);
					package.Save();
				}

				using (var package = new ExcelPackage(fileInfo))
				{
					var sheet = package.Workbook.Worksheets[sheetName];
					var picture = sheet.Drawings["Picture"];
					Assert.AreEqual(4, picture.From.Row);
					Assert.AreEqual(7, picture.From.Column);
				}
			}
			finally
			{
				if (fileInfo.Exists)
					fileInfo.Delete();
			}
		}
		#endregion

		#region Read Chart Tests
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\ScatterFromExcel.xlsx")]
		public void ReadExcelScatterChart()
		{
			var file = new FileInfo("ScatterFromExcel.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var sheet = package.Workbook.Worksheets[1];
				var drawing = sheet.Drawings[0] as ExcelScatterChart;
				Assert.IsNotNull(drawing);
				Assert.AreEqual("Sheet1!$C$20:$C$42", drawing.Series[0].Series);
				Assert.AreEqual("Sheet1!$B$20:$B$42", drawing.Series[0].XSeries);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\BubbleFromExcel.xlsx")]
		public void ReadExcelBubbleChart()
		{
			var file = new FileInfo("BubbleFromExcel.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var sheet = package.Workbook.Worksheets[1];
				var drawing = sheet.Drawings[0] as ExcelBubbleChart;
				Assert.IsNotNull(drawing);
				Assert.AreEqual("Sheet1!$C$20:$C$42", drawing.Series[0].Series);
				Assert.AreEqual("Sheet1!$B$20:$B$42", drawing.Series[0].XSeries);
				Assert.AreEqual("Sheet1!$D$20:$D$42", ((ExcelBubbleChartSerie)drawing.Series[0]).BubbleSize);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\3DBubbleFromExcel.xlsx")]
		public void ReadExcel3DBubbleChart()
		{
			var file = new FileInfo("3DBubbleFromExcel.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var sheet = package.Workbook.Worksheets[1];
				var drawing = sheet.Drawings[0] as ExcelBubbleChart;
				Assert.IsNotNull(drawing);
				Assert.AreEqual("Sheet1!$C$20:$C$42", drawing.Series[0].Series);
				Assert.AreEqual("Sheet1!$B$20:$B$42", drawing.Series[0].XSeries);
				Assert.AreEqual("Sheet1!$D$20:$D$42", ((ExcelBubbleChartSerie)drawing.Series[0]).BubbleSize);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\BarFromExcel.xlsx")]
		public void ReadExcelBarChart()
		{
			var file = new FileInfo("BarFromExcel.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var sheet = package.Workbook.Worksheets[1];
				var drawing = sheet.Drawings[0] as ExcelBarChart;
				Assert.IsNotNull(drawing);
				Assert.AreEqual("Sheet1!$B$20:$B$42", drawing.Series[0].Series);
				Assert.AreEqual(string.Empty, drawing.Series[0].XSeries);
				Assert.AreEqual("Sheet1!$C$20:$C$42", drawing.Series[1].Series);
				Assert.AreEqual(string.Empty, drawing.Series[1].XSeries);
				Assert.AreEqual("Sheet1!$D$20:$D$42", drawing.Series[2].Series);
				Assert.AreEqual(string.Empty, drawing.Series[2].XSeries);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\LineChartFromExcel.xlsx")]

		public void ReadExcelLineChart()
		{
			var file = new FileInfo("LineChartFromExcel.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var sheet = package.Workbook.Worksheets[1];
				var drawing = sheet.Drawings[0] as ExcelLineChart;
				Assert.IsNotNull(drawing);
				Assert.AreEqual("Sheet1!$B$3:$B$20", drawing.Series[0].Series);
				Assert.AreEqual(string.Empty, drawing.Series[0].XSeries);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\BoxAndWhiskerFromExcel.xlsx")]
		public void ReadExcelBoxAndWhiskerChart()
		{
			var file = new FileInfo("BoxAndWhiskerFromExcel.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				Assert.AreEqual(6, package.Workbook.Names.Count);
				Assert.AreEqual("Sheet1!$B$20:$B$42", package.Workbook.Names["_xlchart.v1.0"].NameFormula);
				Assert.AreEqual("Sheet1!$C$20:$C$42", package.Workbook.Names["_xlchart.v1.1"].NameFormula);
				Assert.AreEqual("Sheet1!$D$20:$D$42", package.Workbook.Names["_xlchart.v1.2"].NameFormula);
				Assert.AreEqual("Sheet1!$B$20:$B$42", package.Workbook.Names["_xlchart.v1.3"].NameFormula);
				Assert.AreEqual("Sheet1!$C$20:$C$42", package.Workbook.Names["_xlchart.v1.4"].NameFormula);
				Assert.AreEqual("Sheet1!$D$20:$D$42", package.Workbook.Names["_xlchart.v1.5"].NameFormula);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\ColumnFromExcel.xlsx")]
		public void ReadExcelColumnChart()
		{
			var file = new FileInfo("ColumnFromExcel.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var sheet = package.Workbook.Worksheets[1];
				var drawing = sheet.Drawings[0] as ExcelBarChart;
				Assert.AreEqual(eChartType.ColumnStacked, drawing.ChartType);
				Assert.IsNotNull(drawing);
				Assert.AreEqual("Sheet1!$B$20:$B$42", drawing.Series[0].Series);
				Assert.AreEqual(string.Empty, drawing.Series[0].XSeries);
				Assert.AreEqual("Sheet1!$C$20:$C$42", drawing.Series[1].Series);
				Assert.AreEqual(string.Empty, drawing.Series[1].XSeries);
				Assert.AreEqual("Sheet1!$D$20:$D$42", drawing.Series[2].Series);
				Assert.AreEqual(string.Empty, drawing.Series[2].XSeries);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\ComboFromExcel.xlsx")]
		public void ReadExcelComboChart()
		{
			var file = new FileInfo("ComboFromExcel.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var sheet = package.Workbook.Worksheets[1];
				var drawing = sheet.Drawings[0] as ExcelChart;
				Assert.IsNotNull(drawing);
				Assert.AreEqual("Sheet1!$C$20:$C$42", drawing.Series[1].Series);
				Assert.AreEqual("Sheet1!$B$20:$B$42", drawing.Series[0].Series);
				Assert.AreEqual("Sheet1!$D$20:$D$42", drawing.PlotArea.ChartTypes[2].Series[0].Series);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\HistogramFromExcel.xlsx")]
		public void ReadExcelHistogramChart()
		{
			var file = new FileInfo(@"HistogramFromExcel.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				// Excel 2016 Chart Series are stored in named ranges.
				Assert.AreEqual(9, package.Workbook.Names.Count);
				Assert.AreEqual("Sheet1!$B$20:$B$42", package.Workbook.Names["_xlchart.v1.0"].NameFormula);
				Assert.AreEqual("Sheet1!$C$20:$C$42", package.Workbook.Names["_xlchart.v1.1"].NameFormula);
				Assert.AreEqual("Sheet1!$D$20:$D$42", package.Workbook.Names["_xlchart.v1.2"].NameFormula);
				Assert.AreEqual("Sheet1!$B$20:$B$42", package.Workbook.Names["_xlchart.v1.3"].NameFormula);
				Assert.AreEqual("Sheet1!$C$20:$C$42", package.Workbook.Names["_xlchart.v1.4"].NameFormula);
				Assert.AreEqual("Sheet1!$D$20:$D$42", package.Workbook.Names["_xlchart.v1.5"].NameFormula);
				Assert.AreEqual("Sheet1!$B$20:$B$42", package.Workbook.Names["_xlchart.v1.6"].NameFormula);
				Assert.AreEqual("Sheet1!$C$20:$C$42", package.Workbook.Names["_xlchart.v1.7"].NameFormula);
				Assert.AreEqual("Sheet1!$D$20:$D$42", package.Workbook.Names["_xlchart.v1.8"].NameFormula);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\PieFromExcel.xlsx")]
		public void ReadExcelPieChart()
		{
			var file = new FileInfo( @"PieFromExcel.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var sheet = package.Workbook.Worksheets[1];
				var drawing = sheet.Drawings[0] as ExcelPieChart;
				Assert.AreEqual(eChartType.Pie3D, drawing.ChartType);
				Assert.IsNotNull(drawing);
				Assert.AreEqual("Sheet1!$B$20:$B$42", drawing.Series[0].Series);
				Assert.AreEqual(string.Empty, drawing.Series[0].XSeries);
				Assert.AreEqual("Sheet1!$C$20:$C$42", drawing.Series[1].Series);
				Assert.AreEqual(string.Empty, drawing.Series[1].XSeries);
				Assert.AreEqual("Sheet1!$D$20:$D$42", drawing.Series[2].Series);
				Assert.AreEqual(string.Empty, drawing.Series[2].XSeries);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\DoughnutFromExcel.xlsx")]
		public void ReadExcelDoughnutChart()
		{
			var file = new FileInfo(@"DoughnutFromExcel.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var sheet = package.Workbook.Worksheets[1];
				var doughnutChart = sheet.Drawings[0] as ExcelDoughnutChart;
				Assert.AreEqual(eChartType.Doughnut, doughnutChart.ChartType);
				Assert.IsNotNull(doughnutChart);
				Assert.AreEqual("Sheet1!$B$20:$B$42", doughnutChart.Series[0].Series);
				Assert.AreEqual(string.Empty, doughnutChart.Series[0].XSeries);
				Assert.AreEqual("Sheet1!$C$20:$C$42", doughnutChart.Series[1].Series);
				Assert.AreEqual(string.Empty, doughnutChart.Series[1].XSeries);
				Assert.AreEqual("Sheet1!$D$20:$D$42", doughnutChart.Series[2].Series);
				Assert.AreEqual(string.Empty, doughnutChart.Series[2].XSeries);

				var explodedDoughnutChart = sheet.Drawings[1] as ExcelDoughnutChart;
				Assert.AreEqual(eChartType.DoughnutExploded, explodedDoughnutChart.ChartType);
				Assert.IsNotNull(explodedDoughnutChart);
				Assert.AreEqual("Sheet1!$B$20:$B$42", explodedDoughnutChart.Series[0].Series);
				Assert.AreEqual(string.Empty, explodedDoughnutChart.Series[0].XSeries);
				Assert.AreEqual("Sheet1!$C$20:$C$42", explodedDoughnutChart.Series[1].Series);
				Assert.AreEqual(string.Empty, explodedDoughnutChart.Series[1].XSeries);
				Assert.AreEqual("Sheet1!$D$20:$D$42", explodedDoughnutChart.Series[2].Series);
				Assert.AreEqual(string.Empty, explodedDoughnutChart.Series[2].XSeries);
			}
		}

		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\SlicerFromExcel.xlsx")]
		public void ReadExcelSlicer()
		{
			var file = new FileInfo(@"SlicerFromExcel.xlsx");
			Assert.IsTrue(file.Exists);
			using (var package = new ExcelPackage(file))
			{
				var sheet = package.Workbook.Worksheets["PivotTable"];
				var drawing = sheet.Drawings[0] as ExcelSlicerDrawing;
				Assert.IsNotNull(drawing);
				Assert.AreEqual("Description", drawing.Name);
				Assert.IsNotNull(drawing.Slicer);
				Assert.AreEqual("Description", drawing.Slicer.Name);
				Assert.IsNotNull(drawing.Slicer.SlicerCache);
				Assert.AreEqual(new Uri("slicerCaches/slicerCache1.xml", UriKind.Relative), drawing.Slicer.SlicerCache.SlicerCacheUri);
				Assert.AreEqual("Slicer_Description", drawing.Slicer.SlicerCache.Name);
			}
		}
		#endregion

		#region Copy Worksheet Copies Drawings Tests
		[TestMethod]
		[DeploymentItem(@"..\..\Workbooks\SlicerFromExcel.xlsx")]
		public void CopyWorksheetCopiesExcelSlicer()
		{
			var tempFile = new FileInfo(Path.GetTempFileName());
			try
			{
				var file = new FileInfo("SlicerFromExcel.xlsx");
				Assert.IsTrue(file.Exists);
				using (var package = new ExcelPackage(file))
				{
					var oldSheet = package.Workbook.Worksheets["PivotTable"];
					var newSheet = package.Workbook.Worksheets.Add("New Sheet", oldSheet);
					var oldDrawing = oldSheet.Drawings[0] as ExcelSlicerDrawing;
					var newDrawing = newSheet.Drawings[0] as ExcelSlicerDrawing;
					Assert.IsNotNull(newDrawing);
					Assert.AreNotSame(oldDrawing, newDrawing);
					Assert.AreEqual("Description", oldDrawing.Name);
					Assert.AreEqual("Description 1", newDrawing.Name);
					Assert.IsNotNull(newDrawing.Slicer);
					Assert.AreEqual("Description", oldDrawing.Slicer.Name);
					Assert.AreEqual("Description 1", newDrawing.Slicer.Name);
					Assert.IsNotNull(newDrawing.Slicer.SlicerCache);
					Assert.AreEqual(new Uri("slicerCaches/slicerCache1.xml", UriKind.Relative), oldDrawing.Slicer.SlicerCache.SlicerCacheUri);
					Assert.AreEqual(new Uri("slicerCaches/slicerCache2.xml", UriKind.Relative), newDrawing.Slicer.SlicerCache.SlicerCacheUri);
					Assert.AreEqual("Slicer_Description", oldDrawing.Slicer.SlicerCache.Name);
					Assert.AreEqual("Slicer_Description1", newDrawing.Slicer.SlicerCache.Name);
					Assert.AreEqual("PivotTable1", oldDrawing.Slicer.SlicerCache.PivotTables[0].PivotTableName);
					Assert.AreEqual("PivotTable2", oldDrawing.Slicer.SlicerCache.PivotTables[1].PivotTableName);
					Assert.AreEqual("PivotTable3", newDrawing.Slicer.SlicerCache.PivotTables[0].PivotTableName);
					Assert.AreEqual("PivotTable4", newDrawing.Slicer.SlicerCache.PivotTables[1].PivotTableName);
					package.SaveAs(tempFile);
				}
				using (var package = new ExcelPackage(tempFile))
				{
					Assert.AreEqual(2, package.Workbook.SlicerCaches.Count);
					var oldSheet = package.Workbook.Worksheets["PivotTable"];
					var newSheet = package.Workbook.Worksheets["New Sheet"];
					Assert.AreEqual(1, oldSheet.Slicers.Slicers.Count);
					Assert.AreEqual(1, newSheet.Slicers.Slicers.Count);
					var oldDrawing = oldSheet.Drawings[0] as ExcelSlicerDrawing;
					var newDrawing = newSheet.Drawings[0] as ExcelSlicerDrawing;
					Assert.IsNotNull(oldDrawing);
					Assert.IsNotNull(newDrawing);
					Assert.AreNotSame(oldDrawing, newDrawing);
					Assert.AreEqual("Description", oldDrawing.Name);
					Assert.AreEqual("Description 1", newDrawing.Name);
					Assert.IsNotNull(oldDrawing.Slicer);
					Assert.IsNotNull(newDrawing.Slicer);
					Assert.AreEqual("Description", oldDrawing.Slicer.Name);
					Assert.AreEqual("Description 1", newDrawing.Slicer.Name);
					Assert.IsNotNull(oldDrawing.Slicer.SlicerCache);
					Assert.IsNotNull(newDrawing.Slicer.SlicerCache);
					Assert.AreEqual(new Uri("slicerCaches/slicerCache1.xml", UriKind.Relative), oldDrawing.Slicer.SlicerCache.SlicerCacheUri);
					Assert.AreEqual(new Uri("slicerCaches/slicerCache2.xml", UriKind.Relative), newDrawing.Slicer.SlicerCache.SlicerCacheUri);
					Assert.AreEqual("Slicer_Description", oldDrawing.Slicer.SlicerCache.Name);
					Assert.AreEqual("Slicer_Description1", newDrawing.Slicer.SlicerCache.Name);
					Assert.AreEqual("PivotTable1", oldDrawing.Slicer.SlicerCache.PivotTables[0].PivotTableName);
					Assert.AreEqual("PivotTable2", oldDrawing.Slicer.SlicerCache.PivotTables[1].PivotTableName);
					Assert.AreEqual("PivotTable3", newDrawing.Slicer.SlicerCache.PivotTables[0].PivotTableName);
					Assert.AreEqual("PivotTable4", newDrawing.Slicer.SlicerCache.PivotTables[1].PivotTableName);
				}
			}
			finally
			{
				tempFile.Delete();
			}
		}

		[TestMethod]
		public void CopyWorksheetTwiceCopiesExcelSlicer()
		{
			var tempFile = new FileInfo(Path.GetTempFileName());
			try
			{
				var file = new FileInfo("SlicerFromExcel.xlsx");
				Assert.IsTrue(file.Exists);
				using (var package = new ExcelPackage(file))
				{
					var oldSheet = package.Workbook.Worksheets["PivotTable"];
					var newSheet = package.Workbook.Worksheets.Add("New Sheet", oldSheet);
					var newSheet2 = package.Workbook.Worksheets.Add("New Sheet 2", oldSheet);
					var oldDrawing = oldSheet.Drawings[0] as ExcelSlicerDrawing;
					var newDrawing = newSheet.Drawings[0] as ExcelSlicerDrawing;
					var newDrawing2 = newSheet2.Drawings[0] as ExcelSlicerDrawing;
					Assert.IsNotNull(newDrawing);
					Assert.IsNotNull(newDrawing2);
					Assert.AreEqual("Description", oldDrawing.Name);
					Assert.AreEqual("Description 1", newDrawing.Name);
					Assert.AreEqual("Description 2", newDrawing2.Name);
					Assert.IsNotNull(newDrawing.Slicer);
					Assert.IsNotNull(newDrawing2.Slicer);
					Assert.AreEqual("Description", oldDrawing.Slicer.Name);
					Assert.AreEqual("Description 1", newDrawing.Slicer.Name);
					Assert.AreEqual("Description 2", newDrawing2.Slicer.Name);
					Assert.IsNotNull(newDrawing.Slicer.SlicerCache);
					Assert.IsNotNull(newDrawing2.Slicer.SlicerCache);
					Assert.AreEqual(new Uri("slicerCaches/slicerCache1.xml", UriKind.Relative), oldDrawing.Slicer.SlicerCache.SlicerCacheUri);
					Assert.AreEqual(new Uri("slicerCaches/slicerCache2.xml", UriKind.Relative), newDrawing.Slicer.SlicerCache.SlicerCacheUri);
					Assert.AreEqual(new Uri("slicerCaches/slicerCache3.xml", UriKind.Relative), newDrawing2.Slicer.SlicerCache.SlicerCacheUri);
					Assert.AreEqual("Slicer_Description", oldDrawing.Slicer.SlicerCache.Name);
					Assert.AreEqual("Slicer_Description1", newDrawing.Slicer.SlicerCache.Name);
					Assert.AreEqual("Slicer_Description2", newDrawing2.Slicer.SlicerCache.Name);
					Assert.AreEqual("PivotTable1", oldDrawing.Slicer.SlicerCache.PivotTables[0].PivotTableName);
					Assert.AreEqual("PivotTable2", oldDrawing.Slicer.SlicerCache.PivotTables[1].PivotTableName);
					Assert.AreEqual("PivotTable3", newDrawing.Slicer.SlicerCache.PivotTables[0].PivotTableName);
					Assert.AreEqual("PivotTable4", newDrawing.Slicer.SlicerCache.PivotTables[1].PivotTableName);
					Assert.AreEqual("PivotTable5", newDrawing2.Slicer.SlicerCache.PivotTables[0].PivotTableName);
					Assert.AreEqual("PivotTable6", newDrawing2.Slicer.SlicerCache.PivotTables[1].PivotTableName);
					package.SaveAs(tempFile);
				}
				using (var package = new ExcelPackage(tempFile))
				{
					Assert.AreEqual(3, package.Workbook.SlicerCaches.Count);
					var oldSheet = package.Workbook.Worksheets["PivotTable"];
					var newSheet = package.Workbook.Worksheets["New Sheet"];
					var newSheet2 = package.Workbook.Worksheets["New Sheet 2"];
					Assert.AreEqual(1, oldSheet.Slicers.Slicers.Count);
					Assert.AreEqual(1, newSheet.Slicers.Slicers.Count);
					Assert.AreEqual(1, newSheet2.Slicers.Slicers.Count);
					var oldDrawing = oldSheet.Drawings[0] as ExcelSlicerDrawing;
					var newDrawing = newSheet.Drawings[0] as ExcelSlicerDrawing;
					var newDrawing2 = newSheet2.Drawings[0] as ExcelSlicerDrawing;
					Assert.IsNotNull(newDrawing);
					Assert.IsNotNull(newDrawing2);
					Assert.AreEqual("Description", oldDrawing.Name);
					Assert.AreEqual("Description 1", newDrawing.Name);
					Assert.AreEqual("Description 2", newDrawing2.Name);
					Assert.IsNotNull(newDrawing.Slicer);
					Assert.IsNotNull(newDrawing2.Slicer);
					Assert.AreEqual("Description", oldDrawing.Slicer.Name);
					Assert.AreEqual("Description 1", newDrawing.Slicer.Name);
					Assert.AreEqual("Description 2", newDrawing2.Slicer.Name);
					Assert.IsNotNull(newDrawing.Slicer.SlicerCache);
					Assert.IsNotNull(newDrawing2.Slicer.SlicerCache);
					Assert.AreEqual(new Uri("slicerCaches/slicerCache1.xml", UriKind.Relative), oldDrawing.Slicer.SlicerCache.SlicerCacheUri);
					Assert.AreEqual(new Uri("slicerCaches/slicerCache2.xml", UriKind.Relative), newDrawing.Slicer.SlicerCache.SlicerCacheUri);
					Assert.AreEqual(new Uri("slicerCaches/slicerCache3.xml", UriKind.Relative), newDrawing2.Slicer.SlicerCache.SlicerCacheUri);
					Assert.AreEqual("Slicer_Description", oldDrawing.Slicer.SlicerCache.Name);
					Assert.AreEqual("Slicer_Description1", newDrawing.Slicer.SlicerCache.Name);
					Assert.AreEqual("Slicer_Description2", newDrawing2.Slicer.SlicerCache.Name);
					Assert.AreEqual("PivotTable1", oldDrawing.Slicer.SlicerCache.PivotTables[0].PivotTableName);
					Assert.AreEqual("PivotTable2", oldDrawing.Slicer.SlicerCache.PivotTables[1].PivotTableName);
					Assert.AreEqual("PivotTable3", newDrawing.Slicer.SlicerCache.PivotTables[0].PivotTableName);
					Assert.AreEqual("PivotTable4", newDrawing.Slicer.SlicerCache.PivotTables[1].PivotTableName);
					Assert.AreEqual("PivotTable5", newDrawing2.Slicer.SlicerCache.PivotTables[0].PivotTableName);
					Assert.AreEqual("PivotTable6", newDrawing2.Slicer.SlicerCache.PivotTables[1].PivotTableName);
				}
			}
			finally
			{
				tempFile.Delete();
			}
		}
		#endregion

		#region Private Static Methods
		private static void AddTestSerie(ExcelWorksheet ws, ExcelChart chrt)
		{
			AddTestData(ws);
			chrt.Series.Add("'" + ws.Name + "'!V19:V24", "'" + ws.Name + "'!U19:U24");
		}

		private static void AddTestData(ExcelWorksheet ws)
		{
			ws.Cells["U19"].Value = new DateTime(2009, 12, 31);
			ws.Cells["U20"].Value = new DateTime(2010, 1, 1);
			ws.Cells["U21"].Value = new DateTime(2010, 1, 2);
			ws.Cells["U22"].Value = new DateTime(2010, 1, 3);
			ws.Cells["U23"].Value = new DateTime(2010, 1, 4);
			ws.Cells["U24"].Value = new DateTime(2010, 1, 5);
			ws.Cells["U19:U24"].Style.Numberformat.Format = "yyyy-mm-dd";

			ws.Cells["V19"].Value = 100;
			ws.Cells["V20"].Value = 102;
			ws.Cells["V21"].Value = 101;
			ws.Cells["V22"].Value = 103;
			ws.Cells["V23"].Value = 105;
			ws.Cells["V24"].Value = 104;

			ws.Cells["W19"].Value = 105;
			ws.Cells["W20"].Value = 108;
			ws.Cells["W21"].Value = 104;
			ws.Cells["W22"].Value = 121;
			ws.Cells["W23"].Value = 103;
			ws.Cells["W24"].Value = 109;


			ws.Cells["X19"].Value = "öäå";
			ws.Cells["X20"].Value = "ÖÄÅ";
			ws.Cells["X21"].Value = "üÛ";
			ws.Cells["X22"].Value = "&%#¤";
			ws.Cells["X23"].Value = "ÿ";
			ws.Cells["X24"].Value = "û";
		}
		#endregion
	}
}
