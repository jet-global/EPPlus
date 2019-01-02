/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See http://www.codeplex.com/EPPlus for details.
 *
 * Copyright (C) 2011  Jan Källman
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
 * Author							Change						Date
 *******************************************************************************
 * Jan Källman		Added		2009-10-01
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Xml;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.Drawing.Chart
{
	#region Chart Enums
	/// <summary>
	/// Chart type
	/// </summary>
	public enum eChartType
	{
		Area3D = -4098,
		AreaStacked3D = 78,
		AreaStacked1003D = 79,
		BarClustered3D = 60,
		BarStacked3D = 61,
		BarStacked1003D = 62,
		Column3D = -4100,
		ColumnClustered3D = 54,
		ColumnStacked3D = 55,
		ColumnStacked1003D = 56,
		Line3D = -4101,
		Pie3D = -4102,
		PieExploded3D = 70,
		Area = 1,
		AreaStacked = 76,
		AreaStacked100 = 77,
		BarClustered = 57,
		BarOfPie = 71,
		BarStacked = 58,
		BarStacked100 = 59,
		Bubble = 15,
		Bubble3DEffect = 87,
		ColumnClustered = 51,
		ColumnStacked = 52,
		ColumnStacked100 = 53,
		ConeBarClustered = 102,
		ConeBarStacked = 103,
		ConeBarStacked100 = 104,
		ConeCol = 105,
		ConeColClustered = 99,
		ConeColStacked = 100,
		ConeColStacked100 = 101,
		CylinderBarClustered = 95,
		CylinderBarStacked = 96,
		CylinderBarStacked100 = 97,
		CylinderCol = 98,
		CylinderColClustered = 92,
		CylinderColStacked = 93,
		CylinderColStacked100 = 94,
		Doughnut = -4120,
		DoughnutExploded = 80,
		Line = 4,
		LineMarkers = 65,
		LineMarkersStacked = 66,
		LineMarkersStacked100 = 67,
		LineStacked = 63,
		LineStacked100 = 64,
		Pie = 5,
		PieExploded = 69,
		PieOfPie = 68,
		PyramidBarClustered = 109,
		PyramidBarStacked = 110,
		PyramidBarStacked100 = 111,
		PyramidCol = 112,
		PyramidColClustered = 106,
		PyramidColStacked = 107,
		PyramidColStacked100 = 108,
		Radar = -4151,
		RadarFilled = 82,
		RadarMarkers = 81,
		StockHLC = 88,
		StockOHLC = 89,
		StockVHLC = 90,
		StockVOHLC = 91,
		Surface = 83,
		SurfaceTopView = 85,
		SurfaceTopViewWireframe = 86,
		SurfaceWireframe = 84,
		XYScatter = -4169,
		XYScatterLines = 74,
		XYScatterLinesNoMarkers = 75,
		XYScatterSmooth = 72,
		XYScatterSmoothNoMarkers = 73
	}
	/// <summary>
	/// Bar or column
	/// </summary>
	public enum eDirection
	{
		Column,
		Bar
	}
	/// <summary>
	/// How the series are grouped
	/// </summary>
	public enum eGrouping
	{
		Standard,
		Clustered,
		Stacked,
		PercentStacked
	}
	/// <summary>
	/// Shape for bar charts
	/// </summary>
	public enum eShape
	{
		Box,
		Cone,
		ConeToMax,
		Cylinder,
		Pyramid,
		PyramidToMax
	}
	/// <summary>
	/// Smooth or lines markers
	/// </summary>
	public enum eScatterStyle
	{
		LineMarker,
		SmoothMarker,
	}
	public enum eRadarStyle
	{
		/// <summary>
		/// Specifies that the radar chart shall be filled and have lines but no markers.
		/// </summary>
		Filled,
		/// <summary>
		/// Specifies that the radar chart shall have lines and markers but no fill.
		/// </summary>
		Marker,
		/// <summary>
		/// Specifies that the radar chart shall have lines but no markers and no fill.
		/// </summary>
		Standard
	}
	/// <summary>
	/// Bar or pie
	/// </summary>
	public enum ePieType
	{
		Bar,
		Pie
	}
	/// <summary>
	/// Position of the labels
	/// </summary>
	public enum eLabelPosition
	{
		BestFit,
		Left,
		Right,
		Center,
		Top,
		Bottom,
		InBase,
		InEnd,
		OutEnd
	}
	/// <summary>
	/// Axis label position
	/// </summary>
	public enum eTickLabelPosition
	{
		High,
		Low,
		NextTo,
		None
	}
	/// <summary>
	/// Markerstyle
	/// </summary>
	public enum eMarkerStyle
	{
		Circle,
		Dash,
		Diamond,
		Dot,
		None,
		Picture,
		Plus,
		Square,
		Star,
		Triangle,
		X,
	}
	/// <summary>
	/// The time unit of major and minor datetime axis values
	/// </summary>
	public enum eTimeUnit
	{
		Years,
		Months,
		Days,
	}
	/// <summary>
	/// The build in style of the chart.
	/// </summary>
	public enum eChartStyle
	{
		None,
		Style1,
		Style2,
		Style3,
		Style4,
		Style5,
		Style6,
		Style7,
		Style8,
		Style9,
		Style10,
		Style11,
		Style12,
		Style13,
		Style14,
		Style15,
		Style16,
		Style17,
		Style18,
		Style19,
		Style20,
		Style21,
		Style22,
		Style23,
		Style24,
		Style25,
		Style26,
		Style27,
		Style28,
		Style29,
		Style30,
		Style31,
		Style32,
		Style33,
		Style34,
		Style35,
		Style36,
		Style37,
		Style38,
		Style39,
		Style40,
		Style41,
		Style42,
		Style43,
		Style44,
		Style45,
		Style46,
		Style47,
		Style48
	}
	/// <summary>
	/// Type of Trendline for a chart
	/// </summary>
	public enum eTrendLine
	{
		/// <summary>
		/// Specifies the trendline shall be an exponential curve in the form
		/// </summary>
		Exponential,
		/// <summary>
		/// Specifies the trendline shall be a logarithmic curve in the form , where log is the natural
		/// </summary>
		Linear,
		/// <summary>
		/// Specifies the trendline shall be a logarithmic curve in the form , where log is the natural
		/// </summary>
		Logarithmic,
		/// <summary>
		/// Specifies the trendline shall be a moving average of period Period
		/// </summary>
		MovingAvgerage,
		/// <summary>
		/// Specifies the trendline shall be a polynomial curve of order Order in the form 
		/// </summary>
		Polynomial,
		/// <summary>
		/// Specifies the trendline shall be a power curve in the form
		/// </summary>
		Power
	}
	/// <summary>
	/// Specifies the possible ways to display blanks
	/// </summary>
	public enum eDisplayBlanksAs
	{
		/// <summary>
		/// Blank values shall be left as a gap
		/// </summary>
		Gap,
		/// <summary>
		/// Blank values shall be spanned with a line (Line charts)
		/// </summary>
		Span,
		/// <summary>
		/// Blank values shall be treated as zero
		/// </summary>
		Zero
	}
	public enum eSizeRepresents
	{
		/// <summary>
		/// Specifies the area of the bubbles shall be proportional to the bubble size value.
		/// </summary>
		Area,
		/// <summary>
		/// Specifies the radius of the bubbles shall be proportional to the bubble size value.
		/// </summary>
		Width
	}
	#endregion

	/// <summary>
	/// Base class for Chart object.
	/// </summary>
	public class ExcelChart : ExcelDrawing
	{
		#region Constants
		private const string RootPath = "c:chartSpace/c:chart/c:plotArea";
		private const string DisplayBlanksAsPath = "../../c:dispBlanksAs/@val";
		private const string PlotVisibleOnlyPath = "../../c:plotVisOnly/@val";
		private const string ShowDLblsOverMax = "../../c:showDLblsOverMax/@val";
		private const string GroupingPath = "c:grouping/@val";
		private const string VaryColorsPath = "c:varyColors/@val";
		private const string ChartSpaceChartPath = "c:chartSpace/c:chart";
		private const string ChartSpacePath = "c:chartSpace";
		#endregion

		#region Class Variables
		private ExcelChartPlotArea myPlotArea = null;
		private ExcelChartLegend myLegend = null;
		private ExcelDrawingBorder myBorder = null;
		private ExcelDrawingFill myFill = null;
		private ExcelChartTitle myTitle = null;
		private bool mySecondaryAxis = false;
		#endregion

		#region Properties
		/// <summary>
		/// The c:chartSpace/c:chart/c:plotArea chart XmlNode.
		/// </summary>
		protected internal XmlNode ChartNode { get; set; } = null;
		
		/// <summary>
		/// The Chart XmlHelper.
		/// </summary>
		protected XmlHelper ChartXmlHelper { get; set; }

		/// <summary>
		/// The Excel ChartSeries.
		/// </summary>
		protected internal ExcelChartSeries ChartSeries { get; set; }

		/// <summary>
		/// The reference to the worksheet.
		/// </summary>
		public ExcelWorksheet WorkSheet { get; internal set; }

		/// <summary>
		/// The chart xml document.
		/// </summary>
		public XmlDocument ChartXml { get; internal set; }

		/// <summary>
		/// The title of the chart.
		/// </summary>
		public ExcelChartTitle Title
		{
			get
			{
				if (this.myTitle == null)
				{
					this.myTitle = new ExcelChartTitle(this.NameSpaceManager, this.ChartXml.SelectSingleNode(ExcelChart.ChartSpaceChartPath, this.NameSpaceManager));
				}
				return this.myTitle;
			}
		}

		/// <summary>
		/// The type of chart.
		/// </summary>
		public eChartType ChartType { get; internal set; }

		/// <summary>
		/// The chart series.
		/// </summary>
		public virtual ExcelChartSeries Series
		{
			get
			{
				return this.ChartSeries;
			}
		}
		
		/// <summary>
		/// An array containg all axis of all Chart types.
		/// </summary>
		public ExcelChartAxis[] Axis { get; private set; }

		/// <summary>
		/// The X axis.
		/// </summary>
		public ExcelChartAxis XAxis { get; private set; }

		/// <summary>
		/// The Y axis.
		/// </summary>
		public ExcelChartAxis YAxis { get; private set; }
		
		/// <summary>
		/// If true the charttype will use the secondary axis.
		/// The chart must contain a least one other charttype that uses the primary axis.
		/// </summary>
		public bool UseSecondaryAxis
		{
			get
			{
				return this.mySecondaryAxis;
			}
			set
			{
				if (this.mySecondaryAxis != value)
				{
					if (value)
					{
						if (this.IsTypePieDoughnut())
							throw (new Exception("Pie charts do not support axis"));
						else if (this.HasPrimaryAxis() == false)
							throw (new Exception("Can't set to secondary axis when no serie uses the primary axis"));
						if (this.Axis.Length == 2)
							this.AddAxis();
						var nodeList = this.ChartNode.SelectNodes("c:axId", this.NameSpaceManager);
						nodeList[0].Attributes["val"].Value = this.Axis[2].Id;
						nodeList[1].Attributes["val"].Value = this.Axis[3].Id;
						this.XAxis = this.Axis[2];
						this.YAxis = this.Axis[3];
					}
					else
					{
						var nodeList = this.ChartNode.SelectNodes("c:axId", this.NameSpaceManager);
						nodeList[0].Attributes["val"].Value = this.Axis[0].Id;
						nodeList[1].Attributes["val"].Value = this.Axis[1].Id;
						this.XAxis = this.Axis[0];
						this.YAxis = this.Axis[1];
					}
					this.mySecondaryAxis = value;
				}
			}
		}

		/// <summary>
		/// The build-in chart styles. 
		/// </summary>
		public eChartStyle Style
		{
			get
			{
				XmlNode node = this.ChartXml.SelectSingleNode("c:chartSpace/c:style/@val", this.NameSpaceManager);
				if (node == null)
					return eChartStyle.None;
				else
				{
					if (int.TryParse(node.Value, out int v))
						return (eChartStyle)v;
					else
						return eChartStyle.None;
				}
			}
			set
			{
				if (value == eChartStyle.None)
				{
					if (this.ChartXml.SelectSingleNode("c:chartSpace/c:style", this.NameSpaceManager) is XmlElement element)
						element.ParentNode.RemoveChild(element);
				}
				else
				{
					XmlElement element = this.ChartXml.CreateElement("c:style", ExcelPackage.schemaChart);
					element.SetAttribute("val", ((int)value).ToString());
					XmlElement parent = this.ChartXml.SelectSingleNode(ExcelChart.ChartSpacePath, this.NameSpaceManager) as XmlElement;
					parent.InsertBefore(element, parent.SelectSingleNode("c:chart", this.NameSpaceManager));
				}
			}
		}

		/// <summary>
		/// Show data in hidden rows and columns.
		/// </summary>
		public bool ShowHiddenData
		{
			get
			{
				// !!Inverted value!!
				return !this.ChartXmlHelper.GetXmlNodeBool(ExcelChart.PlotVisibleOnlyPath);
			}
			set
			{
				// !!Inverted value!!
				this.ChartXmlHelper.SetXmlNodeBool(ExcelChart.PlotVisibleOnlyPath, !value);
			}
		}

		/// <summary>
		/// The ExcelPivotTable source.
		/// </summary>
		public ExcelPivotTable PivotTableSource { get; private set; }

		/// <summary>
		/// Specifies the possible ways to display blanks.
		/// </summary>
		public eDisplayBlanksAs DisplayBlanksAs
		{
			get
			{
				string value = this.ChartXmlHelper.GetXmlNodeString(ExcelChart.DisplayBlanksAsPath);
				if (string.IsNullOrEmpty(value))
					return eDisplayBlanksAs.Zero; // Issue 14715 Changed in Office 2010-?
				else
					return (eDisplayBlanksAs)Enum.Parse(typeof(eDisplayBlanksAs), value, true);
			}
			set
			{
				this.ChartSeries.SetXmlNodeString(ExcelChart.DisplayBlanksAsPath, value.ToString().ToLower(CultureInfo.InvariantCulture));
			}
		}

		/// <summary>
		/// Specifies data labels over the maximum of the chart shall be shown
		/// </summary>
		public bool ShowDataLabelsOverMaximum
		{
			get
			{
				return this.ChartXmlHelper.GetXmlNodeBool(ExcelChart.ShowDLblsOverMax, true);
			}
			set
			{
				this.ChartXmlHelper.SetXmlNodeBool(ExcelChart.ShowDLblsOverMax, value, true);
			}
		}
		
		/// <summary>
		/// The chart PlotArea.
		/// </summary>
		public ExcelChartPlotArea PlotArea
		{
			get
			{
				if (this.myPlotArea == null)
					this.myPlotArea = new ExcelChartPlotArea(this.NameSpaceManager, this.ChartXml.SelectSingleNode(ExcelChart.RootPath, this.NameSpaceManager), this);
				return this.myPlotArea;
			}
			private set
			{
				myPlotArea = value;
			}
		}

		/// <summary>
		/// Legend
		/// </summary>
		public ExcelChartLegend Legend
		{
			get
			{
				if (this.myLegend == null)
					this.myLegend = new ExcelChartLegend(this.NameSpaceManager, this.ChartXml.SelectSingleNode("c:chartSpace/c:chart/c:legend", this.NameSpaceManager), this);
				return this.myLegend;
			}
		}

		/// <summary>
		/// Border
		/// </summary>
		public ExcelDrawingBorder Border
		{
			get
			{
				if (this.myBorder == null)
					this.myBorder = new ExcelDrawingBorder(this.NameSpaceManager, this.ChartXml.SelectSingleNode(ExcelChart.ChartSpacePath, this.NameSpaceManager), "c:spPr/a:ln");
				return this.myBorder;
			}
		}

		/// <summary>
		/// Fill
		/// </summary>
		public ExcelDrawingFill Fill
		{
			get
			{
				if (this.myFill == null)
					this.myFill = new ExcelDrawingFill(this.NameSpaceManager, this.ChartXml.SelectSingleNode(ExcelChart.ChartSpacePath, this.NameSpaceManager), "c:spPr");
				return this.myFill;
			}
		}

		/// <summary>
		/// 3D-settings
		/// </summary>
		public ExcelView3D View3D
		{
			get
			{
				if (IsType3D())
					return new ExcelView3D(this.NameSpaceManager, this.ChartXml.SelectSingleNode("//c:view3D", this.NameSpaceManager));
				else
					throw (new Exception("Charttype does not support 3D"));
			}
		}

		public eGrouping Grouping
		{
			get
			{
				return GetGroupingEnum(this.ChartXmlHelper?.GetXmlNodeString(ExcelChart.GroupingPath));
			}
			internal set
			{
				this.ChartXmlHelper.SetXmlNodeString(ExcelChart.GroupingPath, this.GetGroupingText(value));
			}
		}
		
		/// <summary>
		/// If the chart has only one serie this varies the colors for each point.
		/// </summary>
		public bool VaryColors
		{
			get
			{
				return this.ChartXmlHelper.GetXmlNodeBool(ExcelChart.VaryColorsPath);
			}
			set
			{
				if (value)
					this.ChartXmlHelper.SetXmlNodeString(ExcelChart.VaryColorsPath, "1");
				else
					this.ChartXmlHelper.SetXmlNodeString(ExcelChart.VaryColorsPath, "0");
			}
		}

		internal Packaging.ZipPackagePart Part { get; set; }

		/// <summary>
		/// Package internal URI
		/// </summary>
		internal Uri UriChart { get; set; }

		internal new string Id
		{
			get { return string.Empty; }
		}
		#endregion

		#region Constructors
		internal ExcelChart(ExcelDrawings drawings, XmlNode node, eChartType type, bool isPivot) :
			 base(drawings, node, "xdr:graphicFrame/xdr:nvGraphicFramePr/xdr:cNvPr/@name")
		{
			this.ChartType = type;
			this.CreateNewChart(drawings, type, null);

			this.Init(drawings, this.ChartNode);

			this.ChartSeries = new ExcelChartSeries(this, drawings.NameSpaceManager, this.ChartNode, isPivot);

			this.SetTypeProperties();
			this.LoadAxis();
		}

		internal ExcelChart(ExcelDrawings drawings, XmlNode node, eChartType type, ExcelChart topChart, ExcelPivotTable PivotTableSource) :
			 base(drawings, node, "xdr:graphicFrame/xdr:nvGraphicFramePr/xdr:cNvPr/@name")
		{
			this.ChartType = type;
			this.CreateNewChart(drawings, type, topChart);

			this.Init(drawings, this.ChartNode);

			this.ChartSeries = new ExcelChartSeries(this, drawings.NameSpaceManager, this.ChartNode, PivotTableSource != null);
			if (PivotTableSource != null)
				this.SetPivotSource(PivotTableSource);

			this.SetTypeProperties();
			if (topChart == null)
				this.LoadAxis();
			else
			{
				this.Axis = topChart.Axis;
				if (this.Axis.Length > 0)
				{
					this.XAxis = this.Axis[0];
					this.YAxis = this.Axis[1];
				}
			}
		}

		internal ExcelChart(ExcelDrawings drawings, XmlNode node, Uri uriChart, ZipPackagePart part, XmlDocument chartXml, XmlNode chartNode) :
			 base(drawings, node, "xdr:graphicFrame/xdr:nvGraphicFramePr/xdr:cNvPr/@name")
		{
			this.UriChart = uriChart;
			this.Part = part;
			this.ChartXml = chartXml;
			this.ChartNode = chartNode;
			// Get preliminary chart type so that chart series can be initialized correctly.
			this.ChartType = this.GetChartType(chartNode.LocalName);
			this.InitChartLoad(drawings, chartNode);
			// Set precise chart type based on observed chart series.
			this.ChartType = this.GetChartType(chartNode.LocalName);
		}

		internal ExcelChart(ExcelChart topChart, XmlNode chartNode) :
			 base(topChart._drawings, topChart.TopNode, "xdr:graphicFrame/xdr:nvGraphicFramePr/xdr:cNvPr/@name")
		{
			this.UriChart = topChart.UriChart;
			this.Part = topChart.Part;
			this.ChartXml = topChart.ChartXml;
			this.PlotArea = topChart.PlotArea;
			this.ChartNode = chartNode;

			this.InitChartLoad(topChart._drawings, chartNode);
		}

		private void InitChartLoad(ExcelDrawings drawings, XmlNode chartNode)
		{
			bool isPivot = false;
			this.Init(drawings, chartNode);
			this.ChartSeries = new ExcelChartSeries(this, drawings.NameSpaceManager, this.ChartNode, isPivot);
			this.LoadAxis();
		}

		private void Init(ExcelDrawings drawings, XmlNode chartNode)
		{
			this.ChartXmlHelper = XmlHelperFactory.Create(drawings.NameSpaceManager, chartNode);
			this.ChartXmlHelper.SchemaNodeOrder = new string[] { "ofPieType", "title", "pivotFmt", "autoTitleDeleted", "view3D", "floor", "sideWall", "backWall", "plotArea", "wireframe", "barDir", "grouping", "scatterStyle", "radarStyle", "varyColors", "ser", "dLbls", "bubbleScale", "showNegBubbles", "dropLines", "upDownBars", "marker", "smooth", "shape", "legend", "plotVisOnly", "dispBlanksAs", "gapWidth", "showDLblsOverMax", "overlap", "bandFmts", "axId", "spPr", "printSettings" };
			this.WorkSheet = drawings.Worksheet;
		}
		#endregion

		#region Private Methods
		private void SetTypeProperties()
		{
			/******* Grouping *******/
			if (this.IsTypeClustered())
				this.Grouping = eGrouping.Clustered;
			else if (this.IsTypeStacked())
				this.Grouping = eGrouping.Stacked;
			else if (this.IsTypePercentStacked())
				this.Grouping = eGrouping.PercentStacked;

			/***** 3D Perspective *****/
			if (this.IsType3D())
			{
				this.View3D.RotY = 20;
				this.View3D.Perspective = 30;  // Default to 30
				if (this.IsTypePieDoughnut())
					this.View3D.RotX = 30;
				else
					this.View3D.RotX = 15;
			}
		}

		private void CreateNewChart(ExcelDrawings drawings, eChartType type, ExcelChart topChart)
		{
			if (topChart == null)
			{
				XmlElement graphFrame = this.TopNode.OwnerDocument.CreateElement("graphicFrame", ExcelPackage.schemaSheetDrawings);
				graphFrame.SetAttribute("macro", string.Empty);
				this.TopNode.AppendChild(graphFrame);
				graphFrame.InnerXml = string.Format("<xdr:nvGraphicFramePr><xdr:cNvPr id=\"{0}\" name=\"Chart 1\" /><xdr:cNvGraphicFramePr /></xdr:nvGraphicFramePr><xdr:xfrm><a:off x=\"0\" y=\"0\" /> <a:ext cx=\"0\" cy=\"0\" /></xdr:xfrm><a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rId1\" />   </a:graphicData>  </a:graphic>", this._id);
				this.TopNode.AppendChild(this.TopNode.OwnerDocument.CreateElement("clientData", ExcelPackage.schemaSheetDrawings));

				var package = drawings.Worksheet.Package.Package;
				this.UriChart = GetNewUri(package, "/xl/charts/chart{0}.xml");

				this.ChartXml = new XmlDocument() { PreserveWhitespace = ExcelPackage.preserveWhitespace };
				LoadXmlSafe(this.ChartXml, this.ChartStartXml(type), Encoding.UTF8);

				// save it to the package
				this.Part = package.CreatePart(this.UriChart, "application/vnd.openxmlformats-officedocument.drawingml.chart+xml", this._drawings.Package.Compression);

				StreamWriter streamChart = new StreamWriter(this.Part.GetStream(FileMode.Create, FileAccess.Write));
				this.ChartXml.Save(streamChart);
				streamChart.Close();
				package.Flush();

				var chartRelation = drawings.Part.CreateRelationship(UriHelper.GetRelativeUri(drawings.UriDrawing, this.UriChart), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/chart");
				graphFrame.SelectSingleNode("a:graphic/a:graphicData/c:chart", this.NameSpaceManager).Attributes["r:id"].Value = chartRelation.Id;
				package.Flush();
				this.ChartNode = this.ChartXml.SelectSingleNode(string.Format("c:chartSpace/c:chart/c:plotArea/{0}", GetChartNodeText()), this.NameSpaceManager);
			}
			else
			{
				this.ChartXml = topChart.ChartXml;
				this.Part = topChart.Part;
				this.PlotArea = topChart.PlotArea;
				this.UriChart = topChart.UriChart;
				this.Axis = topChart.Axis;

				XmlNode preNode = this.PlotArea.ChartTypes[this.PlotArea.ChartTypes.Count - 1].ChartNode;
				this.ChartNode = this.ChartXml.CreateElement(this.GetChartNodeText(), ExcelPackage.schemaChart);
				preNode.ParentNode.InsertAfter(this.ChartNode, preNode);
				if (topChart.Axis.Length == 0)
					this.AddAxis();
				string serieXML = GetChartSerieStartXml(type, int.Parse(topChart.Axis[0].Id), int.Parse(topChart.Axis[1].Id), topChart.Axis.Length > 2 ? int.Parse(topChart.Axis[2].Id) : -1);
				this.ChartNode.InnerXml = serieXML;
			}
			this.GetPositionSize();
		}

		private void LoadAxis()
		{
			XmlNodeList nodeList = this.ChartNode.SelectNodes("c:axId", this.NameSpaceManager);
			List<ExcelChartAxis> excelChartAxes = new List<ExcelChartAxis>();
			foreach (XmlNode node in nodeList)
			{
				string id = node.Attributes["val"].Value;
				var axNode = this.ChartXml.SelectNodes(ExcelChart.RootPath + string.Format("/*/c:axId[@val=\"{0}\"]", id), this.NameSpaceManager);
				if (axNode != null && axNode.Count > 1)
				{
					foreach (XmlNode axn in axNode)
					{
						if (axn.ParentNode.LocalName.EndsWith("Ax"))
						{
							XmlNode axisNode = axNode[1].ParentNode;
							ExcelChartAxis axis = new ExcelChartAxis(this.NameSpaceManager, axisNode);
							excelChartAxes.Add(axis);
						}
					}
				}
			}
			this.Axis = excelChartAxes.ToArray();

			if (this.Axis.Length > 0)
				this.XAxis = this.Axis[0];
			if (this.Axis.Length > 1)
				this.YAxis = this.Axis[1];
		}

		/// <summary>
		/// Remove all axis that are not used any more
		/// </summary>
		/// <param name="excelChartAxis"></param>
		private void CheckRemoveAxis(ExcelChartAxis excelChartAxis)
		{
			if (this.ExistsAxis(excelChartAxis))
			{
				// Remove the axis
				ExcelChartAxis[] newAxis = new ExcelChartAxis[this.Axis.Length - 1];
				int pos = 0;
				foreach (var ax in this.Axis)
				{
					if (ax != excelChartAxis)
						newAxis[pos] = ax;
				}
				// Update all charttypes.
				foreach (ExcelChart chartType in this.PlotArea.ChartTypes)
				{
					chartType.Axis = newAxis;
				}
			}
		}

		private bool ExistsAxis(ExcelChartAxis excelChartAxis)
		{
			foreach (ExcelChart chartType in this.PlotArea.ChartTypes)
			{
				if (chartType != this && 
					(chartType.XAxis.AxisPosition == excelChartAxis.AxisPosition ||
						chartType.YAxis.AxisPosition == excelChartAxis.AxisPosition))
				{
						return true;
				}
			}
			return false;
		}

		private bool HasPrimaryAxis()
		{
			if (this.PlotArea.ChartTypes.Count == 1)
				return false;
			foreach (var chart in this.PlotArea.ChartTypes)
			{
				if (chart != this && chart.UseSecondaryAxis == false && chart.IsTypePieDoughnut() == false)
					return true;
			}
			return false;
		}

		#region Grouping Enum Translation
		private string GetGroupingText(eGrouping grouping)
		{
			switch (grouping)
			{
				case eGrouping.Clustered:
					return "clustered";
				case eGrouping.Stacked:
					return "stacked";
				case eGrouping.PercentStacked:
					return "percentStacked";
				default:
					return "standard";
			}
		}

		private eGrouping GetGroupingEnum(string grouping)
		{
			switch (grouping)
			{
				case "stacked":
					return eGrouping.Stacked;
				case "percentStacked":
					return eGrouping.PercentStacked;
				default: // "clustered":               
					return eGrouping.Clustered;
			}
		}
		#endregion

		#region Xml Init Functions
		private string ChartStartXml(eChartType type)
		{
			StringBuilder xml = new StringBuilder();
			int axID = 1;
			int xAxID = 2;
			int serAxID = this.IsTypeSurface() ? 3 : -1;

			xml.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
			xml.AppendFormat("<c:chartSpace xmlns:c=\"{0}\" xmlns:a=\"{1}\" xmlns:r=\"{2}\">", ExcelPackage.schemaChart, ExcelPackage.schemaDrawings, ExcelPackage.schemaRelationships);
			xml.Append("<c:chart>");
			xml.AppendFormat("{0}{1}<c:plotArea><c:layout/>", this.AddPerspectiveXml(type), this.AddSurfaceXml(type));

			string chartNodeText = this.GetChartNodeText();
			xml.AppendFormat("<{0}>", chartNodeText);
			xml.Append(this.GetChartSerieStartXml(type, axID, xAxID, serAxID));
			xml.AppendFormat("</{0}>", chartNodeText);

			if (!this.IsTypePieDoughnut())
			{
				if (this.IsTypeScatterBubble())
					xml.AppendFormat("<c:valAx><c:axId val=\"{0}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"b\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"{1}\"/><c:crosses val=\"autoZero\"/></c:valAx>", axID, xAxID);
				else
					xml.AppendFormat("<c:catAx><c:axId val=\"{0}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"b\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"{1}\"/><c:crosses val=\"autoZero\"/><c:auto val=\"1\"/><c:lblAlgn val=\"ctr\"/><c:lblOffset val=\"100\"/></c:catAx>", axID, xAxID);
				xml.AppendFormat("<c:valAx><c:axId val=\"{1}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"l\"/><c:majorGridlines/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"{0}\"/><c:crosses val=\"autoZero\"/><c:crossBetween val=\"between\"/></c:valAx>", axID, xAxID);
				if (serAxID == 3) //Surface Chart
					xml.AppendFormat("<c:serAx><c:axId val=\"{0}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\"/><c:axPos val=\"b\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"{1}\"/><c:crosses val=\"autoZero\"/></c:serAx>", serAxID, xAxID);
			}
			xml.AppendFormat("</c:plotArea><c:legend><c:legendPos val=\"r\"/><c:layout/><c:overlay val=\"0\" /></c:legend><c:plotVisOnly val=\"1\"/></c:chart>", axID, xAxID);
			xml.Append("<c:printSettings><c:headerFooter/><c:pageMargins b=\"0.75\" l=\"0.7\" r=\"0.7\" t=\"0.75\" header=\"0.3\" footer=\"0.3\"/><c:pageSetup/></c:printSettings></c:chartSpace>");
			return xml.ToString();
		}

		private string GetChartSerieStartXml(eChartType type, int axID, int xAxID, int serAxID)
		{
			StringBuilder xml = new StringBuilder();
			xml.Append(this.AddScatterType(type));
			xml.Append(this.AddRadarType(type));
			xml.Append(this.AddBarDir(type));
			xml.Append(this.AddGrouping());
			xml.Append(this.AddVaryColors());
			xml.Append(this.AddHasMarker(type));
			xml.Append(this.AddShape(type));
			xml.Append(this.AddFirstSliceAng(type));
			xml.Append(this.AddHoleSize(type));
			if (this.ChartType == eChartType.BarStacked100 ||
				 this.ChartType == eChartType.BarStacked ||
				 this.ChartType == eChartType.ColumnStacked ||
				 this.ChartType == eChartType.ColumnStacked100)
			{
				xml.Append("<c:overlap val=\"100\"/>");
			}
			if (this.IsTypeSurface())
			{
				xml.Append("<c:bandFmts/>");
			}
			xml.Append(this.AddAxisId(axID, xAxID, serAxID));
			return xml.ToString();
		}

		private string AddAxisId(int axID, int xAxID, int serAxID)
		{
			if (!this.IsTypePieDoughnut())
			{
				if (this.IsTypeSurface())
					return string.Format("<c:axId val=\"{0}\"/><c:axId val=\"{1}\"/><c:axId val=\"{2}\"/>", axID, xAxID, serAxID);
				else
					return string.Format("<c:axId val=\"{0}\"/><c:axId val=\"{1}\"/>", axID, xAxID);
			}
			else
				return string.Empty;
		}

		private string AddAxType()
		{
			switch (this.ChartType)
			{
				case eChartType.XYScatter:
				case eChartType.XYScatterLines:
				case eChartType.XYScatterLinesNoMarkers:
				case eChartType.XYScatterSmooth:
				case eChartType.XYScatterSmoothNoMarkers:
				case eChartType.Bubble:
				case eChartType.Bubble3DEffect:
					return "valAx";
				default:
					return "catAx";
			}
		}

		private string AddScatterType(eChartType type)
		{
			if (type == eChartType.XYScatter ||
				 type == eChartType.XYScatterLines ||
				 type == eChartType.XYScatterLinesNoMarkers ||
				 type == eChartType.XYScatterSmooth ||
				 type == eChartType.XYScatterSmoothNoMarkers)
			{
				return "<c:scatterStyle val=\"\" />";
			}
			else
				return string.Empty;
		}

		private string AddRadarType(eChartType type)
		{
			if (type == eChartType.Radar ||
				 type == eChartType.RadarFilled ||
				 type == eChartType.RadarMarkers)
			{
				return "<c:radarStyle val=\"\" />";
			}
			else
				return string.Empty;
		}

		private string AddGrouping()
		{
			if (this.IsTypeShape() || this.IsTypeLine())
				return "<c:grouping val=\"standard\"/>";
			else
				return string.Empty;
		}

		private string AddHoleSize(eChartType type)
		{
			if (type == eChartType.Doughnut ||
				 type == eChartType.DoughnutExploded)
				return "<c:holeSize val=\"50\" />";
			else
				return string.Empty;
		}

		private string AddFirstSliceAng(eChartType type)
		{
			if (type == eChartType.Doughnut ||
				 type == eChartType.DoughnutExploded)
				return "<c:firstSliceAng val=\"0\" />";
			else
				return string.Empty;
		}

		private string AddVaryColors()
		{
			if (this.IsTypePieDoughnut())
				return "<c:varyColors val=\"1\" />";
			else
				return "<c:varyColors val=\"0\" />";
		}

		private string AddHasMarker(eChartType type)
		{
			if (type == eChartType.LineMarkers ||
				 type == eChartType.LineMarkersStacked ||
				 type == eChartType.LineMarkersStacked100 /*||
               type == eChartType.XYScatterLines ||
               type == eChartType.XYScatterSmooth*/)
			{
				return "<c:marker val=\"1\"/>";
			}
			else
				return string.Empty;
		}

		private string AddShape(eChartType type)
		{
			if (this.IsTypeShape())
				return "<c:shape val=\"box\" />";
			else
				return string.Empty;
		}

		private string AddBarDir(eChartType type)
		{
			if (this.IsTypeShape())
				return "<c:barDir val=\"col\" />";
			else
				return string.Empty;
		}

		private string AddPerspectiveXml(eChartType type)
		{
			if (this.IsType3D())
				return "<c:view3D><c:perspective val=\"30\" /></c:view3D>";
			else
				return string.Empty;
		}

		private string AddSurfaceXml(eChartType type)
		{
			if (this.IsTypeSurface())
				return this.AddSurfacePart("floor") + this.AddSurfacePart("sideWall") + this.AddSurfacePart("backWall");
			else
				return string.Empty;
		}

		private string AddSurfacePart(string name)
		{
			return string.Format("<c:{0}><c:thickness val=\"0\"/><c:spPr><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/><a:sp3d/></c:spPr></c:{0}>", name);
		}
		#endregion
		#endregion

		#region Protected Methods
		/// <summary>
		/// Get the name of the chart node
		/// </summary>
		/// <returns>The name</returns>
		protected string GetChartNodeText()
		{
			switch (this.ChartType)
			{
				case eChartType.Area3D:
				case eChartType.AreaStacked3D:
				case eChartType.AreaStacked1003D:
					return "c:area3DChart";
				case eChartType.Area:
				case eChartType.AreaStacked:
				case eChartType.AreaStacked100:
					return "c:areaChart";
				case eChartType.BarClustered:
				case eChartType.BarStacked:
				case eChartType.BarStacked100:
				case eChartType.ColumnClustered:
				case eChartType.ColumnStacked:
				case eChartType.ColumnStacked100:
					return "c:barChart";
				case eChartType.BarClustered3D:
				case eChartType.BarStacked3D:
				case eChartType.BarStacked1003D:
				case eChartType.ColumnClustered3D:
				case eChartType.ColumnStacked3D:
				case eChartType.ColumnStacked1003D:
				case eChartType.ConeBarClustered:
				case eChartType.ConeBarStacked:
				case eChartType.ConeBarStacked100:
				case eChartType.ConeCol:
				case eChartType.ConeColClustered:
				case eChartType.ConeColStacked:
				case eChartType.ConeColStacked100:
				case eChartType.CylinderBarClustered:
				case eChartType.CylinderBarStacked:
				case eChartType.CylinderBarStacked100:
				case eChartType.CylinderCol:
				case eChartType.CylinderColClustered:
				case eChartType.CylinderColStacked:
				case eChartType.CylinderColStacked100:
				case eChartType.PyramidBarClustered:
				case eChartType.PyramidBarStacked:
				case eChartType.PyramidBarStacked100:
				case eChartType.PyramidCol:
				case eChartType.PyramidColClustered:
				case eChartType.PyramidColStacked:
				case eChartType.PyramidColStacked100:
					return "c:bar3DChart";
				case eChartType.Bubble:
				case eChartType.Bubble3DEffect:
					return "c:bubbleChart";
				case eChartType.Doughnut:
				case eChartType.DoughnutExploded:
					return "c:doughnutChart";
				case eChartType.Line:
				case eChartType.LineMarkers:
				case eChartType.LineMarkersStacked:
				case eChartType.LineMarkersStacked100:
				case eChartType.LineStacked:
				case eChartType.LineStacked100:
					return "c:lineChart";
				case eChartType.Line3D:
					return "c:line3DChart";
				case eChartType.Pie:
				case eChartType.PieExploded:
					return "c:pieChart";
				case eChartType.BarOfPie:
				case eChartType.PieOfPie:
					return "c:ofPieChart";
				case eChartType.Pie3D:
				case eChartType.PieExploded3D:
					return "c:pie3DChart";
				case eChartType.Radar:
				case eChartType.RadarFilled:
				case eChartType.RadarMarkers:
					return "c:radarChart";
				case eChartType.XYScatter:
				case eChartType.XYScatterLines:
				case eChartType.XYScatterLinesNoMarkers:
				case eChartType.XYScatterSmooth:
				case eChartType.XYScatterSmoothNoMarkers:
					return "c:scatterChart";
				case eChartType.Surface:
				case eChartType.SurfaceWireframe:
					return "c:surface3DChart";
				case eChartType.SurfaceTopView:
				case eChartType.SurfaceTopViewWireframe:
					return "c:surfaceChart";
				case eChartType.StockHLC:
					return "c:stockChart";
				default:
					throw (new NotImplementedException("Chart type not implemented"));
			}
		}
		#endregion

		#region Internal Methods
		/// <summary>
		/// Add a secondary axis.
		/// </summary>
		internal void AddAxis()
		{
			XmlElement catAx = this.ChartXml.CreateElement(string.Format("c:{0}", this.AddAxType()), ExcelPackage.schemaChart);
			int axID;
			if (this.Axis.Length == 0)
			{
				this.PlotArea.TopNode.AppendChild(catAx);
				axID = 1;
			}
			else
			{
				this.Axis[0].TopNode.ParentNode.InsertAfter(catAx, this.Axis[this.Axis.Length - 1].TopNode);
				axID = int.Parse(this.Axis[0].Id) < int.Parse(this.Axis[1].Id) ? int.Parse(this.Axis[1].Id) + 1 : int.Parse(this.Axis[0].Id) + 1;
			}

			XmlElement valAx = this.ChartXml.CreateElement("c:valAx", ExcelPackage.schemaChart);
			catAx.ParentNode.InsertAfter(valAx, catAx);

			if (this.Axis.Length == 0)
			{
				catAx.InnerXml = string.Format("<c:axId val=\"{0}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\" /><c:axPos val=\"b\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"{1}\"/><c:crosses val=\"autoZero\"/><c:auto val=\"1\"/><c:lblAlgn val=\"ctr\"/><c:lblOffset val=\"100\"/>", axID, axID + 1);
				valAx.InnerXml = string.Format("<c:axId val=\"{1}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\" /><c:axPos val=\"l\"/><c:majorGridlines/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"{0}\"/><c:crosses val=\"autoZero\"/><c:crossBetween val=\"between\"/>", axID, axID + 1);
			}
			else
			{
				catAx.InnerXml = string.Format("<c:axId val=\"{0}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"1\" /><c:axPos val=\"b\"/><c:tickLblPos val=\"none\"/><c:crossAx val=\"{1}\"/><c:crosses val=\"autoZero\"/>", axID, axID + 1);
				valAx.InnerXml = string.Format("<c:axId val=\"{0}\"/><c:scaling><c:orientation val=\"minMax\"/></c:scaling><c:delete val=\"0\" /><c:axPos val=\"r\"/><c:tickLblPos val=\"nextTo\"/><c:crossAx val=\"{1}\"/><c:crosses val=\"max\"/><c:crossBetween val=\"between\"/>", axID + 1, axID);
			}

			if (this.Axis.Length == 0)
				this.Axis = new ExcelChartAxis[2];
			else
			{
				ExcelChartAxis[] newAxis = new ExcelChartAxis[this.Axis.Length + 2];
				Array.Copy(this.Axis, newAxis, this.Axis.Length);
				this.Axis = newAxis;
			}

			this.Axis[this.Axis.Length - 2] = new ExcelChartAxis(this.NameSpaceManager, catAx);
			this.Axis[this.Axis.Length - 1] = new ExcelChartAxis(this.NameSpaceManager, valAx);
			foreach (var chart in this.PlotArea.ChartTypes)
			{
				chart.Axis = this.Axis;
			}
		}

		internal void RemoveSecondaryAxis()
		{
			throw (new NotImplementedException("Not yet implemented"));
		}

		internal static ExcelChart GetChart(ExcelDrawings drawings, XmlNode node)
		{
			XmlNode chartNode = node.SelectSingleNode("xdr:graphicFrame/a:graphic/a:graphicData/c:chart", drawings.NameSpaceManager);
			if (chartNode != null)
			{
				var drawingRelation = drawings.Part.GetRelationship(chartNode.Attributes["r:id"].Value);
				var uriChart = UriHelper.ResolvePartUri(drawings.UriDrawing, drawingRelation.TargetUri);

				var part = drawings.Part.Package.GetPart(uriChart);
				var chartXml = new XmlDocument();
				LoadXmlSafe(chartXml, part.GetStream());

				ExcelChart topChart = null;
				foreach (XmlElement n in chartXml.SelectSingleNode(ExcelChart.RootPath, drawings.NameSpaceManager).ChildNodes)
				{
					if (topChart == null)
					{
						topChart = GetChart(n, drawings, node, uriChart, part, chartXml, null);
						if (topChart != null)
							topChart.PlotArea.ChartTypes.Add(topChart);
					}
					else
					{
						var subChart = GetChart(n, null, null, null, null, null, topChart);
						if (subChart != null)
							topChart.PlotArea.ChartTypes.Add(subChart);
					}
				}
				return topChart;
			}
			else
				return null;
		}

		internal static ExcelChart GetChart(XmlElement chartNode, ExcelDrawings drawings, XmlNode node, Uri uriChart, Packaging.ZipPackagePart part, XmlDocument chartXml, ExcelChart topChart)
		{
			switch (chartNode.LocalName)
			{
				case "area3DChart":
				case "areaChart":
				case "stockChart":
					if (topChart == null)
						return new ExcelChart(drawings, node, uriChart, part, chartXml, chartNode);
					else
						return new ExcelChart(topChart, chartNode);
				case "surface3DChart":
				case "surfaceChart":
					if (topChart == null)
						return new ExcelSurfaceChart(drawings, node, uriChart, part, chartXml, chartNode);
					else
						return new ExcelSurfaceChart(topChart, chartNode);
				case "radarChart":
					if (topChart == null)
						return new ExcelRadarChart(drawings, node, uriChart, part, chartXml, chartNode);
					else
						return new ExcelRadarChart(topChart, chartNode);
				case "bubbleChart":
					if (topChart == null)
						return new ExcelBubbleChart(drawings, node, uriChart, part, chartXml, chartNode);
					else
						return new ExcelBubbleChart(topChart, chartNode);
				case "barChart":
				case "bar3DChart":
					if (topChart == null)
						return new ExcelBarChart(drawings, node, uriChart, part, chartXml, chartNode);
					else
						return new ExcelBarChart(topChart, chartNode);
				case "doughnutChart":
					if (topChart == null)
						return new ExcelDoughnutChart(drawings, node, uriChart, part, chartXml, chartNode);
					else
						return new ExcelDoughnutChart(topChart, chartNode);
				case "pie3DChart":
				case "pieChart":
					if (topChart == null)
						return new ExcelPieChart(drawings, node, uriChart, part, chartXml, chartNode);
					else
						return new ExcelPieChart(topChart, chartNode);
				case "ofPieChart":
					if (topChart == null)
						return new ExcelOfPieChart(drawings, node, uriChart, part, chartXml, chartNode);
					else
						return new ExcelBarChart(topChart, chartNode);
				case "lineChart":
				case "line3DChart":
					if (topChart == null)
						return new ExcelLineChart(drawings, node, uriChart, part, chartXml, chartNode);
					else
						return new ExcelLineChart(topChart, chartNode);
				case "scatterChart":
					if (topChart == null)
						return new ExcelScatterChart(drawings, node, uriChart, part, chartXml, chartNode);
					else
						return new ExcelScatterChart(topChart, chartNode);
				default:
					return null;
			}
		}
		internal static ExcelChart GetNewChart(ExcelDrawings drawings, XmlNode drawNode, eChartType chartType, ExcelChart topChart, ExcelPivotTable PivotTableSource)
		{
			switch (chartType)
			{
				case eChartType.Pie:
				case eChartType.PieExploded:
				case eChartType.Pie3D:
				case eChartType.PieExploded3D:
					return new ExcelPieChart(drawings, drawNode, chartType, topChart, PivotTableSource);
				case eChartType.BarOfPie:
				case eChartType.PieOfPie:
					return new ExcelOfPieChart(drawings, drawNode, chartType, topChart, PivotTableSource);
				case eChartType.Doughnut:
				case eChartType.DoughnutExploded:
					return new ExcelDoughnutChart(drawings, drawNode, chartType, topChart, PivotTableSource);
				case eChartType.BarClustered:
				case eChartType.BarStacked:
				case eChartType.BarStacked100:
				case eChartType.BarClustered3D:
				case eChartType.BarStacked3D:
				case eChartType.BarStacked1003D:
				case eChartType.ConeBarClustered:
				case eChartType.ConeBarStacked:
				case eChartType.ConeBarStacked100:
				case eChartType.CylinderBarClustered:
				case eChartType.CylinderBarStacked:
				case eChartType.CylinderBarStacked100:
				case eChartType.PyramidBarClustered:
				case eChartType.PyramidBarStacked:
				case eChartType.PyramidBarStacked100:
				case eChartType.ColumnClustered:
				case eChartType.ColumnStacked:
				case eChartType.ColumnStacked100:
				case eChartType.Column3D:
				case eChartType.ColumnClustered3D:
				case eChartType.ColumnStacked3D:
				case eChartType.ColumnStacked1003D:
				case eChartType.ConeCol:
				case eChartType.ConeColClustered:
				case eChartType.ConeColStacked:
				case eChartType.ConeColStacked100:
				case eChartType.CylinderCol:
				case eChartType.CylinderColClustered:
				case eChartType.CylinderColStacked:
				case eChartType.CylinderColStacked100:
				case eChartType.PyramidCol:
				case eChartType.PyramidColClustered:
				case eChartType.PyramidColStacked:
				case eChartType.PyramidColStacked100:
					return new ExcelBarChart(drawings, drawNode, chartType, topChart, PivotTableSource);
				case eChartType.XYScatter:
				case eChartType.XYScatterLines:
				case eChartType.XYScatterLinesNoMarkers:
				case eChartType.XYScatterSmooth:
				case eChartType.XYScatterSmoothNoMarkers:
					return new ExcelScatterChart(drawings, drawNode, chartType, topChart, PivotTableSource);
				case eChartType.Line:
				case eChartType.Line3D:
				case eChartType.LineMarkers:
				case eChartType.LineMarkersStacked:
				case eChartType.LineMarkersStacked100:
				case eChartType.LineStacked:
				case eChartType.LineStacked100:
					return new ExcelLineChart(drawings, drawNode, chartType, topChart, PivotTableSource);
				case eChartType.Bubble:
				case eChartType.Bubble3DEffect:
					return new ExcelBubbleChart(drawings, drawNode, chartType, topChart, PivotTableSource);
				case eChartType.Radar:
				case eChartType.RadarFilled:
				case eChartType.RadarMarkers:
					return new ExcelRadarChart(drawings, drawNode, chartType, topChart, PivotTableSource);
				case eChartType.Surface:
				case eChartType.SurfaceTopView:
				case eChartType.SurfaceTopViewWireframe:
				case eChartType.SurfaceWireframe:
					return new ExcelSurfaceChart(drawings, drawNode, chartType, topChart, PivotTableSource);
				default:
					return new ExcelChart(drawings, drawNode, chartType, topChart, PivotTableSource);
			}
		}

		internal void SetPivotSource(ExcelPivotTable pivotTableSource)
		{
			this.PivotTableSource = pivotTableSource;
			XmlElement chart = this.ChartXml.SelectSingleNode(ExcelChart.ChartSpaceChartPath, this.NameSpaceManager) as XmlElement;

			var pivotSource = this.ChartXml.CreateElement("pivotSource", ExcelPackage.schemaChart);
			chart.ParentNode.InsertBefore(pivotSource, chart);
			pivotSource.InnerXml = string.Format("<c:name>[]{0}!{1}</c:name><c:fmtId val=\"0\"/>", this.PivotTableSource.Worksheet.Name, pivotTableSource.Name);

			var fmts = this.ChartXml.CreateElement("pivotFmts", ExcelPackage.schemaChart);
			chart.PrependChild(fmts);
			fmts.InnerXml = "<c:pivotFmt><c:idx val=\"0\"/><c:marker><c:symbol val=\"none\"/></c:marker></c:pivotFmt>";
			this.Series.AddPivotSerie(pivotTableSource);
		}

		internal override void DeleteMe()
		{
			try
			{
				this.Part.Package.DeletePart(this.UriChart);
			}
			catch (Exception ex)
			{
				throw (new InvalidDataException("EPPlus internal error when deleteing chart.", ex));
			}
			base.DeleteMe();
		}

		internal virtual eChartType GetChartType(string name)
		{
			switch (name)
			{
				case "area3DChart":
					if (this.Grouping == eGrouping.Stacked)
						return eChartType.AreaStacked3D;
					else if (this.Grouping == eGrouping.PercentStacked)
						return eChartType.AreaStacked1003D;
					else
						return eChartType.Area3D;
				case "areaChart":
					if (this.Grouping == eGrouping.Stacked)
						return eChartType.AreaStacked;
					else if (this.Grouping == eGrouping.PercentStacked)
						return eChartType.AreaStacked100;
					else
						return eChartType.Area;
				case "doughnutChart":
					return eChartType.Doughnut;
				case "pie3DChart":
					return eChartType.Pie3D;
				case "pieChart":
					return eChartType.Pie;
				case "radarChart":
					return eChartType.Radar;
				case "scatterChart":
					return eChartType.XYScatter;
				case "surface3DChart":
				case "surfaceChart":
					return eChartType.Surface;
				case "stockChart":
					return eChartType.StockHLC;
				default:
					return 0;
			}
		}
		#endregion

		#region Chart type functions
		internal static bool IsType3D(eChartType chartType)
		{
			return chartType == eChartType.Area3D ||
						chartType == eChartType.AreaStacked3D ||
						chartType == eChartType.AreaStacked1003D ||
						chartType == eChartType.BarClustered3D ||
						chartType == eChartType.BarStacked3D ||
						chartType == eChartType.BarStacked1003D ||
						chartType == eChartType.Column3D ||
						chartType == eChartType.ColumnClustered3D ||
						chartType == eChartType.ColumnStacked3D ||
						chartType == eChartType.ColumnStacked1003D ||
						chartType == eChartType.Line3D ||
						chartType == eChartType.Pie3D ||
						chartType == eChartType.PieExploded3D ||
						chartType == eChartType.ConeBarClustered ||
						chartType == eChartType.ConeBarStacked ||
						chartType == eChartType.ConeBarStacked100 ||
						chartType == eChartType.ConeCol ||
						chartType == eChartType.ConeColClustered ||
						chartType == eChartType.ConeColStacked ||
						chartType == eChartType.ConeColStacked100 ||
						chartType == eChartType.CylinderBarClustered ||
						chartType == eChartType.CylinderBarStacked ||
						chartType == eChartType.CylinderBarStacked100 ||
						chartType == eChartType.CylinderCol ||
						chartType == eChartType.CylinderColClustered ||
						chartType == eChartType.CylinderColStacked ||
						chartType == eChartType.CylinderColStacked100 ||
						chartType == eChartType.PyramidBarClustered ||
						chartType == eChartType.PyramidBarStacked ||
						chartType == eChartType.PyramidBarStacked100 ||
						chartType == eChartType.PyramidCol ||
						chartType == eChartType.PyramidColClustered ||
						chartType == eChartType.PyramidColStacked ||
						chartType == eChartType.PyramidColStacked100 ||
						chartType == eChartType.Surface ||
						chartType == eChartType.SurfaceTopView ||
						chartType == eChartType.SurfaceTopViewWireframe ||
						chartType == eChartType.SurfaceWireframe;
		}

		internal protected bool IsType3D()
		{
			return IsType3D(this.ChartType);
		}

		protected bool IsTypeLine()
		{
			return this.ChartType == eChartType.Line ||
					  this.ChartType == eChartType.LineMarkers ||
					  this.ChartType == eChartType.LineMarkersStacked100 ||
					  this.ChartType == eChartType.LineStacked ||
					  this.ChartType == eChartType.LineStacked100 ||
					  this.ChartType == eChartType.Line3D;
		}

		protected bool IsTypeScatterBubble()
		{
			return this.ChartType == eChartType.XYScatter ||
					  this.ChartType == eChartType.XYScatterLines ||
					  this.ChartType == eChartType.XYScatterLinesNoMarkers ||
					  this.ChartType == eChartType.XYScatterSmooth ||
					  this.ChartType == eChartType.XYScatterSmoothNoMarkers ||
					  this.ChartType == eChartType.Bubble ||
					  this.ChartType == eChartType.Bubble3DEffect;
		}

		protected bool IsTypeSurface()
		{
			return this.ChartType == eChartType.Surface ||
					 this.ChartType == eChartType.SurfaceTopView ||
					 this.ChartType == eChartType.SurfaceTopViewWireframe ||
					 this.ChartType == eChartType.SurfaceWireframe;
		}

		protected bool IsTypeShape()
		{
			return this.ChartType == eChartType.BarClustered3D ||
					  this.ChartType == eChartType.BarStacked3D ||
					  this.ChartType == eChartType.BarStacked1003D ||
					  this.ChartType == eChartType.BarClustered3D ||
					  this.ChartType == eChartType.BarStacked3D ||
					  this.ChartType == eChartType.BarStacked1003D ||
					  this.ChartType == eChartType.Column3D ||
					  this.ChartType == eChartType.ColumnClustered3D ||
					  this.ChartType == eChartType.ColumnStacked3D ||
					  this.ChartType == eChartType.ColumnStacked1003D ||
					  //this.ChartType == eChartType.3DPie ||
					  //this.ChartType == eChartType.3DPieExploded ||
					  //this.ChartType == eChartType.Bubble3DEffect ||
					  this.ChartType == eChartType.ConeBarClustered ||
					  this.ChartType == eChartType.ConeBarStacked ||
					  this.ChartType == eChartType.ConeBarStacked100 ||
					  this.ChartType == eChartType.ConeCol ||
					  this.ChartType == eChartType.ConeColClustered ||
					  this.ChartType == eChartType.ConeColStacked ||
					  this.ChartType == eChartType.ConeColStacked100 ||
					  this.ChartType == eChartType.CylinderBarClustered ||
					  this.ChartType == eChartType.CylinderBarStacked ||
					  this.ChartType == eChartType.CylinderBarStacked100 ||
					  this.ChartType == eChartType.CylinderCol ||
					  this.ChartType == eChartType.CylinderColClustered ||
					  this.ChartType == eChartType.CylinderColStacked ||
					  this.ChartType == eChartType.CylinderColStacked100 ||
					  this.ChartType == eChartType.PyramidBarClustered ||
					  this.ChartType == eChartType.PyramidBarStacked ||
					  this.ChartType == eChartType.PyramidBarStacked100 ||
					  this.ChartType == eChartType.PyramidCol ||
					  this.ChartType == eChartType.PyramidColClustered ||
					  this.ChartType == eChartType.PyramidColStacked ||
					  this.ChartType == eChartType.PyramidColStacked100; //||
						//this.ChartType == eChartType.Doughnut ||
						//this.ChartType == eChartType.DoughnutExploded;
		}

		protected internal bool IsTypePercentStacked()
		{
			return this.ChartType == eChartType.AreaStacked100 ||
						this.ChartType == eChartType.BarStacked100 ||
						this.ChartType == eChartType.BarStacked1003D ||
						this.ChartType == eChartType.ColumnStacked100 ||
						this.ChartType == eChartType.ColumnStacked1003D ||
						this.ChartType == eChartType.ConeBarStacked100 ||
						this.ChartType == eChartType.ConeColStacked100 ||
						this.ChartType == eChartType.CylinderBarStacked100 ||
						this.ChartType == eChartType.CylinderColStacked ||
						this.ChartType == eChartType.LineMarkersStacked100 ||
						this.ChartType == eChartType.LineStacked100 ||
						this.ChartType == eChartType.PyramidBarStacked100 ||
						this.ChartType == eChartType.PyramidColStacked100;
		}

		protected internal bool IsTypeStacked()
		{
			return this.ChartType == eChartType.AreaStacked ||
						this.ChartType == eChartType.AreaStacked3D ||
						this.ChartType == eChartType.BarStacked ||
						this.ChartType == eChartType.BarStacked3D ||
						this.ChartType == eChartType.ColumnStacked3D ||
						this.ChartType == eChartType.ColumnStacked ||
						this.ChartType == eChartType.ConeBarStacked ||
						this.ChartType == eChartType.ConeColStacked ||
						this.ChartType == eChartType.CylinderBarStacked ||
						this.ChartType == eChartType.CylinderColStacked ||
						this.ChartType == eChartType.LineMarkersStacked ||
						this.ChartType == eChartType.LineStacked ||
						this.ChartType == eChartType.PyramidBarStacked ||
						this.ChartType == eChartType.PyramidColStacked;
		}

		protected bool IsTypeClustered()
		{
			return this.ChartType == eChartType.BarClustered ||
						this.ChartType == eChartType.BarClustered3D ||
						this.ChartType == eChartType.ColumnClustered3D ||
						this.ChartType == eChartType.ColumnClustered ||
						this.ChartType == eChartType.ConeBarClustered ||
						this.ChartType == eChartType.ConeColClustered ||
						this.ChartType == eChartType.CylinderBarClustered ||
						this.ChartType == eChartType.CylinderColClustered ||
						this.ChartType == eChartType.PyramidBarClustered ||
						this.ChartType == eChartType.PyramidColClustered;
		}

		protected internal bool IsTypePieDoughnut()
		{
			return this.ChartType == eChartType.Pie ||
						this.ChartType == eChartType.PieExploded ||
						this.ChartType == eChartType.PieOfPie ||
						this.ChartType == eChartType.Pie3D ||
						this.ChartType == eChartType.PieExploded3D ||
						this.ChartType == eChartType.BarOfPie ||
						this.ChartType == eChartType.Doughnut ||
						this.ChartType == eChartType.DoughnutExploded;
		}
		#endregion
	}
}
