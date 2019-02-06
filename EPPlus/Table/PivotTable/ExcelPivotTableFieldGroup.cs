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
 * Jan Källman		Added		21-MAR-2011
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using System;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
	#region Enums
	/// <summary>
	/// The possible types of date groupings for a field.
	/// </summary>
	public enum PivotFieldDateGrouping
	{
		/// <summary>
		/// Group a pivot field items by seconds.
		/// </summary>
		Seconds,
		/// <summary>
		/// Group a pivot field items by minutes.
		/// </summary>
		Minutes,
		/// <summary>
		/// Group a pivot field items by hours.
		/// </summary>
		Hours,
		/// <summary>
		/// Group a pivot field items by days.
		/// </summary>
		Days,
		/// <summary>
		/// Group a pivot field items by months.
		/// </summary>
		Months,
		/// <summary>
		/// Group a pivot field items by quarters.
		/// </summary>
		Quarters,
		/// <summary>
		/// Group a pivot field items by years.
		/// </summary>
		Years,
		/// <summary>
		/// Not a date grouping field.
		/// </summary>
		None
	}
	#endregion

	/// <summary>
	/// Base class for pivot table field groups.
	/// </summary>
	public class ExcelPivotTableFieldGroup : XmlHelper
	{
		#region Properties
		/// <summary>
		/// Get the base field index for this group field.
		/// </summary>
		public int BaseField
		{
			get { return base.GetXmlNodeInt("@base"); }
		}
		
		/// <summary>
		/// Get the collection of group items.
		/// </summary>
		public SharedItemsCollection GroupItems { get; }

		/// <summary>
		/// Get the grouping type of how this field is grouped.
		/// </summary>
		public PivotFieldDateGrouping GroupBy { get; }

		/// <summary>
		/// Get the discrete grouping properties collection.
		/// </summary>
		public DiscreteGroupingPropertiesCollection DiscreteGroupingProperties { get; }

		private string RangeGroupingProperties
		{
			get { return base.GetXmlNodeString("d:rangePr/@groupBy"); }
		}
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="ExcelPivotTableFieldGroup"/>.
		/// </summary>
		/// <param name="ns">The namespace of the worksheet.</param>
		/// <param name="topNode">The top node in the xml.</param>
		internal ExcelPivotTableFieldGroup(XmlNamespaceManager ns, XmlNode topNode) :
			 base(ns, topNode)
		{
			if (ns == null)
				throw new ArgumentNullException(nameof(ns));
			if (topNode == null)
				throw new ArgumentNullException(nameof(topNode));
			var groupItemsNode = topNode.SelectSingleNode("d:groupItems", this.NameSpaceManager);
			if (groupItemsNode != null)
				this.GroupItems = new SharedItemsCollection(this.NameSpaceManager, groupItemsNode);
			var discretePrNode = topNode.SelectSingleNode("d:discretePr", this.NameSpaceManager);
			if (discretePrNode != null)
				this.DiscreteGroupingProperties = new DiscreteGroupingPropertiesCollection(this.NameSpaceManager, discretePrNode);
			if (string.IsNullOrEmpty(this.RangeGroupingProperties))
				this.GroupBy = PivotFieldDateGrouping.None;
			else
				this.GroupBy = (PivotFieldDateGrouping)Enum.Parse(typeof(PivotFieldDateGrouping), this.RangeGroupingProperties, true);
		}
		#endregion
	}
}