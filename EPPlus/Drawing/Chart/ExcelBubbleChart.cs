using System;
using System.Globalization;
using System.Xml;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Table.PivotTable;

namespace OfficeOpenXml.Drawing.Chart
{
	/// <summary>
	/// Provides access to bubble chart specific properties
	/// </summary>
	public sealed class ExcelBubbleChart : ExcelChart
	{
		#region Constants
		private const string BubbleScalePath = "c:bubbleScale/@val";
		private const string ShowNegativeBubblesPath = "c:showNegBubbles/@val";
		private const string Bubble3dPath = "c:bubble3D/@val";
		private const string SIZEREPRESENTS_PATH = "c:sizeRepresents/@val";
		#endregion

		#region Properties
		/// <summary>
		/// Specifies the scale factor for the bubble chart. Can range from 0 to 300, corresponding to a percentage of the default size.
		/// </summary>
		public int BubbleScale
		{
			get
			{
				return this.ChartXmlHelper.GetXmlNodeInt(ExcelBubbleChart.BubbleScalePath);
			}
			set
			{
				if (value < 0 && value > 300)
					throw (new ArgumentOutOfRangeException("Bubblescale out of range. 0-300 allowed"));
				this.ChartXmlHelper.SetXmlNodeString(ExcelBubbleChart.BubbleScalePath, value.ToString());
			}
		}

		/// <summary>
		/// Specifies the scale factor for the bubble chart. Can range from 0 to 300, corresponding to a percentage of the default size.
		/// </summary>
		public bool ShowNegativeBubbles
		{
			get
			{
				return this.ChartXmlHelper.GetXmlNodeBool(ExcelBubbleChart.ShowNegativeBubblesPath);
			}
			set
			{
				this.ChartXmlHelper.SetXmlNodeBool(ExcelBubbleChart.BubbleScalePath, value, true);
			}
		}

		/// <summary>
		/// Specifies if the bubblechart is three dimensional.
		/// </summary>
		public bool Bubble3D
		{
			get
			{
				return this.ChartXmlHelper?.GetXmlNodeBool(ExcelBubbleChart.Bubble3dPath) ?? false;
			}
			set
			{
				this.ChartXmlHelper.SetXmlNodeBool(ExcelBubbleChart.Bubble3dPath, value);
				this.ChartType = value ? eChartType.Bubble3DEffect : eChartType.Bubble;
			}
		}

		/// <summary>
		/// Specifies the scale factor for the bubble chart. Can range from 0 to 300, corresponding to a percentage of the default size.
		/// </summary>
		public eSizeRepresents SizeRepresents
		{
			get
			{
				var xmlNodeString = this.ChartXmlHelper.GetXmlNodeString(ExcelBubbleChart.SIZEREPRESENTS_PATH).ToLower(CultureInfo.InvariantCulture);
				if (xmlNodeString == "w")
					return eSizeRepresents.Width;
				else
					return eSizeRepresents.Area;
			}
			set
			{
				this.ChartXmlHelper.SetXmlNodeString(ExcelBubbleChart.SIZEREPRESENTS_PATH, value == eSizeRepresents.Width ? "w" : "area");
			}
		}

		/// <summary>
		/// The chart Series.
		/// </summary>
		public new ExcelBubbleChartSeries Series
		{
			get
			{
				return (ExcelBubbleChartSeries)this.ChartSeries;
			}
		}
		#endregion

		#region Constructors
		internal ExcelBubbleChart(ExcelDrawings drawings, XmlNode node, eChartType type, ExcelChart topChart, ExcelPivotTable PivotTableSource) :
			 base(drawings, node, type, topChart, PivotTableSource)
		{
			this.ShowNegativeBubbles = false;
			this.BubbleScale = 100;
			this.ChartSeries = new ExcelBubbleChartSeries(this, drawings.NameSpaceManager, this.ChartNode, PivotTableSource != null);
		}

		internal ExcelBubbleChart(ExcelDrawings drawings, XmlNode node, eChartType type, bool isPivot) :
			 base(drawings, node, type, isPivot)
		{
			this.ChartSeries = new ExcelBubbleChartSeries(this, drawings.NameSpaceManager, this.ChartNode, isPivot);
		}

		internal ExcelBubbleChart(ExcelDrawings drawings, XmlNode node, Uri uriChart, ZipPackagePart part, XmlDocument chartXml, XmlNode chartNode) :
			 base(drawings, node, uriChart, part, chartXml, chartNode)
		{
			this.ChartSeries = new ExcelBubbleChartSeries(this, this._drawings.NameSpaceManager, this.ChartNode, false);
		}

		internal ExcelBubbleChart(ExcelChart topChart, XmlNode chartNode) :
			 base(topChart, chartNode)
		{
			this.ChartSeries = new ExcelBubbleChartSeries(this, this._drawings.NameSpaceManager, this.ChartNode, false);
		}
		#endregion

		#region Internal Methods
		internal override eChartType GetChartType(string name)
		{
			if (this.Bubble3D)
				return eChartType.Bubble3DEffect;
			else
				return eChartType.Bubble;
		}
		#endregion
	}
}
