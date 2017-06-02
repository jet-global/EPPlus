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
using System.Xml;
using OfficeOpenXml.Table.PivotTable;

namespace OfficeOpenXml.Drawing.Chart
{
	/// <summary>
	/// A Surface chart.
	/// </summary>
	public sealed class ExcelSurfaceChart : ExcelChart
	{
		#region Constants
		const string WireframePath = "c:wireframe/@val";
		#endregion

		#region Constructors
		internal ExcelSurfaceChart(ExcelDrawings drawings, XmlNode node, eChartType type, ExcelChart topChart, ExcelPivotTable PivotTableSource) :
			 base(drawings, node, type, topChart, PivotTableSource)
		{
			Init();
		}

		internal ExcelSurfaceChart(ExcelDrawings drawings, XmlNode node, Uri uriChart, Packaging.ZipPackagePart part, XmlDocument chartXml, XmlNode chartNode) :
			base(drawings, node, uriChart, part, chartXml, chartNode)
		{
			Init();
		}

		internal ExcelSurfaceChart(ExcelChart topChart, XmlNode chartNode) :
			 base(topChart, chartNode)
		{
			Init();
		}

		private void Init()
		{
			this.Floor = new ExcelChartSurface(this.NameSpaceManager, this.ChartXmlHelper.TopNode.SelectSingleNode("c:floor", this.NameSpaceManager));
			this.BackWall = new ExcelChartSurface(this.NameSpaceManager, this.ChartXmlHelper.TopNode.SelectSingleNode("c:sideWall", this.NameSpaceManager));
			this.SideWall = new ExcelChartSurface(this.NameSpaceManager, this.ChartXmlHelper.TopNode.SelectSingleNode("c:backWall", this.NameSpaceManager));
			this.SetTypeProperties();
		}
		#endregion

		#region Properties
		public ExcelChartSurface Floor { get; private set; }

		public ExcelChartSurface SideWall { get; private set; }

		public ExcelChartSurface BackWall { get; private set; }

		public bool Wireframe
		{
			get
			{
				return this.ChartXmlHelper.GetXmlNodeBool(ExcelSurfaceChart.WireframePath);
			}
			set
			{
				this.ChartXmlHelper.SetXmlNodeBool(ExcelSurfaceChart.WireframePath, value);
			}
		}
		#endregion

		#region Internal Methods
		internal void SetTypeProperties()
		{
			if (this.ChartType == eChartType.SurfaceWireframe || this.ChartType == eChartType.SurfaceTopViewWireframe)
				this.Wireframe = true;
			else
				this.Wireframe = false;
			if (this.ChartType == eChartType.SurfaceTopView || this.ChartType == eChartType.SurfaceTopViewWireframe)
			{
				this.View3D.RotY = 0;
				this.View3D.RotX = 90;
			}
			else
			{
				this.View3D.RotY = 20;
				this.View3D.RotX = 15;
			}
			this.View3D.RightAngleAxes = false;
			this.View3D.Perspective = 0;
			this.Axis[1].CrossBetween = eCrossBetween.MidCat;
		}

		internal override eChartType GetChartType(string name)
		{
			if (this.Wireframe)
			{
				if (name == "surfaceChart")
					return eChartType.SurfaceTopViewWireframe;
				else
					return eChartType.SurfaceWireframe;
			}
			else
			{
				if (name == "surfaceChart")
					return eChartType.SurfaceTopView;
				else
					return eChartType.Surface;
			}
		}
		#endregion
	}
}
