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
 * ******************************************************************************
 * Jan Källman		Initial Release		        2009-10-01
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using System;
using System.Globalization;
using System.Xml;
using OfficeOpenXml.Table.PivotTable;

namespace OfficeOpenXml.Drawing.Chart
{
	/// <summary>
	/// Provides access to doughnut chart specific properties
	/// </summary>
	public class ExcelDoughnutChart : ExcelPieChart
	{
		#region Constants
		private const string FirstSliceAnglePath = "c:firstSliceAng/@val";
		private const string HoleSizePath = "c:holeSize/@val";
		#endregion

		#region Properties
		/// <summary>
		/// Angle of the first slize
		/// </summary>
		public decimal FirstSliceAngle
		{
			get
			{
				return this.ChartXmlHelper.GetXmlNodeDecimal(FirstSliceAnglePath);
			}
			internal set
			{
				this.ChartXmlHelper.SetXmlNodeString(FirstSliceAnglePath, value.ToString(CultureInfo.InvariantCulture));
			}
		}
		
		/// <summary>
		/// Size of the doughnnut hole
		/// </summary>
		public decimal HoleSize
		{
			get
			{
				return this.ChartXmlHelper.GetXmlNodeDecimal(HoleSizePath);
			}
			internal set
			{
				this.ChartXmlHelper.SetXmlNodeString(HoleSizePath, value.ToString(CultureInfo.InvariantCulture));
			}
		}
		#endregion

		#region Constructors
		internal ExcelDoughnutChart(ExcelDrawings drawings, XmlNode node, eChartType type, bool isPivot) :
			 base(drawings, node, type, isPivot)
		{
		}

		internal ExcelDoughnutChart(ExcelDrawings drawings, XmlNode node, eChartType type, ExcelChart topChart, ExcelPivotTable PivotTableSource) :
			 base(drawings, node, type, topChart, PivotTableSource)
		{
		}

		internal ExcelDoughnutChart(ExcelDrawings drawings, XmlNode node, Uri uriChart, Packaging.ZipPackagePart part, XmlDocument chartXml, XmlNode chartNode) :
			base(drawings, node, uriChart, part, chartXml, chartNode)
		{
		}

		internal ExcelDoughnutChart(ExcelChart topChart, XmlNode chartNode) :
			 base(topChart, chartNode)
		{
		}
		#endregion

		#region Overrides
		internal override eChartType GetChartType(string name)
		{
			if (name == "doughnutChart")
			{
				if (this.IsExploded())
					return eChartType.DoughnutExploded;
				else
					return eChartType.Doughnut;
			}
			return base.GetChartType(name);
		}
		#endregion

		#region Private Methods
		private bool IsExploded()
		{
			if (this.Series == null)
				return false;
			for (int i = 0; i < this.Series.Count; i++)
			{
				if (((ExcelPieChartSerie)this.Series[i]).Explosion > 0)
					return true;
			}
			return false;
		}
		#endregion
	}
}
