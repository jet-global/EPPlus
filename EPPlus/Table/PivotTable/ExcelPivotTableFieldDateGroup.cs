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
using System;
using System.Globalization;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
	/// <summary>
	/// A pivot table field date group.
	/// </summary>
	public class ExcelPivotTableFieldDateGroup : ExcelPivotTableFieldGroup
	{
		#region Constants
		private const string GroupByPath = "d:fieldGroup/d:rangePr/@groupBy";
		#endregion

		#region Properties
		/// <summary>
		/// Gets how to group the date field.
		/// </summary>
		public eDateGroupBy GroupBy
		{
			get
			{
				string v = base.GetXmlNodeString(ExcelPivotTableFieldDateGroup.GroupByPath);
				if (v != "")
					return (eDateGroupBy)Enum.Parse(typeof(eDateGroupBy), v, true);
				else
					throw (new Exception("Invalid date Groupby"));
			}
			private set
			{
				base.SetXmlNodeString(ExcelPivotTableFieldDateGroup.GroupByPath, value.ToString().ToLower(CultureInfo.InvariantCulture));
			}
		}

		/// <summary>
		/// Gets whether there exist an auto detect start date.
		/// </summary>
		public bool AutoStart
		{
			get
			{
				return base.GetXmlNodeBool("@autoStart", false);
			}
		}

		/// <summary>
		/// Gets whether there exist an auto detect end date.
		/// </summary>
		public bool AutoEnd
		{
			get
			{
				return base.GetXmlNodeBool("@autoStart", false);
			}
		}
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="ExcelPivotTableFieldDateGroup"/>.
		/// </summary>
		/// <param name="ns">The namespace of the worksheet.</param>
		/// <param name="topNode">The top node in the xml.</param>
		internal ExcelPivotTableFieldDateGroup(XmlNamespaceManager ns, XmlNode topNode) :
			 base(ns, topNode)
		{

		}
		#endregion
	}
}