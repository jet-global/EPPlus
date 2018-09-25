/*******************************************************************************
* You may amend and distribute as you like, but don't remove this header!
*
* EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
* See http://www.codeplex.com/EPPlus for details.
*
* Copyright (C) 2011-2018 Jan Källman, Evan Schallerer, and others as noted in the source history.
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
* For code change notes, see the source control history.
*******************************************************************************/
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
	/// <summary>
	/// A pivot table field numeric grouping.
	/// </summary>
	public class ExcelPivotTableFieldNumericGroup : ExcelPivotTableFieldGroup
	{
		#region Constants
		private const string StartPath = "d:fieldGroup/d:rangePr/@startNum";
		private const string EndPath = "d:fieldGroup/d:rangePr/@endNum";
		private const string GroupIntervalPath = "d:fieldGroup/d:rangePr/@groupInterval";
		#endregion

		#region Properties
		/// <summary>
		/// Gets the start value.
		/// </summary>
		public double Start
		{
			get
			{
				return (double)base.GetXmlNodeDoubleNull(ExcelPivotTableFieldNumericGroup.StartPath);
			}
			private set
			{
				base.SetXmlNodeString(ExcelPivotTableFieldNumericGroup.StartPath, value.ToString(CultureInfo.InvariantCulture));
			}
		}

		/// <summary>
		/// Gets the end value.
		/// </summary>
		public double End
		{
			get
			{
				return (double)base.GetXmlNodeDoubleNull(ExcelPivotTableFieldNumericGroup.EndPath);
			}
			private set
			{
				base.SetXmlNodeString(ExcelPivotTableFieldNumericGroup.EndPath, value.ToString(CultureInfo.InvariantCulture));
			}
		}

		/// <summary>
		/// Gets the interval.
		/// </summary>
		public double Interval
		{
			get
			{
				return (double)base.GetXmlNodeDoubleNull(ExcelPivotTableFieldNumericGroup.GroupIntervalPath);
			}
			private set
			{
				base.SetXmlNodeString(ExcelPivotTableFieldNumericGroup.GroupIntervalPath, value.ToString(CultureInfo.InvariantCulture));
			}
		}
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="ExcelPivotTableFieldNumericGroup"/>.
		/// </summary>
		/// <param name="ns">The namespace of the worksheet.</param>
		/// <param name="topNode">The top node in the xml.</param>
		internal ExcelPivotTableFieldNumericGroup(XmlNamespaceManager ns, XmlNode topNode) :
			 base(ns, topNode)
		{

		}
		#endregion
	}
}