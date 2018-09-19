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
using System.Collections.Generic;

namespace OfficeOpenXml.Table.PivotTable
{
	/// <summary>
	/// A collection of pivot table objects.
	/// </summary>
	public class ExcelPivotTableCollection : IEnumerable<ExcelPivotTable>
	{
		#region Class Variables
		/// <summary>
		/// A list of pivot tables.
		/// </summary>
		internal Dictionary<string, int> myPivotTableNames = new Dictionary<string, int>();
		private List<ExcelPivotTable> myPivotTables = new List<ExcelPivotTable>();
		private ExcelWorksheet myWorksheet;
		#endregion

		#region Properties
		/// <summary>
		/// Gets the count of existing pivot tables.
		/// </summary>
		public int Count
		{
			get
			{
				return myPivotTables.Count;
			}
		}
		
		/// <summary>
		/// Gets the pivot table Index starting at base 0.
		/// </summary>
		/// <param name="Index">The position of the pivot table.</param>
		/// <returns>The pivot table at the given index.</returns>
		public ExcelPivotTable this[int Index]
		{
			get
			{
				if (Index < 0 || Index >= myPivotTables.Count)
					throw (new ArgumentOutOfRangeException("PivotTable index out of range"));
				return myPivotTables[Index];
			}
		}
		
		/// <summary>
		/// Gets the pivot tables accesed by name.
		/// </summary>
		/// <param name="Name">The name of the pivot table.</param>
		/// <returns>The pivot table or null if there is no match found.</returns>
		public ExcelPivotTable this[string Name]
		{
			get
			{
				if (myPivotTableNames.ContainsKey(Name))
					return myPivotTables[myPivotTableNames[Name]];
				else
					return null;
			}
		}
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="ExcelPivotTableCollection"/>.
		/// </summary>
		/// <param name="ws">The worksheet of the pivot tables.</param>
		internal ExcelPivotTableCollection(ExcelWorksheet ws)
		{
			var pck = ws.Package.Package;
			myWorksheet = ws;
			foreach (var rel in ws.Part.GetRelationships())
			{
				if (rel.RelationshipType == ExcelPackage.schemaRelationships + "/pivotTable")
				{
					var tbl = new ExcelPivotTable(rel, ws);
					myPivotTableNames.Add(tbl.Name, myPivotTables.Count);
					myPivotTables.Add(tbl);
				}
			}
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Add a pivot table on the supplied range.
		/// </summary>
		/// <param name="range">The range address including header and total row</param>
		/// <param name="source">The Source data range address</param>
		/// <param name="name">The name of the table. Must be unique </param>
		/// <returns>The pivot table object</returns>
		public ExcelPivotTable Add(ExcelAddress range, ExcelRangeBase source, string name)
		{
			if (string.IsNullOrEmpty(name))
				name = this.GetNewTableName();
			if (range.WorkSheet != myWorksheet.Name)
				throw (new Exception("The Range must be in the current worksheet"));
			else if (myWorksheet.Workbook.ExistsTableName(name))
				throw (new ArgumentException("Tablename is not unique"));
			foreach (var t in myPivotTables)
			{
				if (t.Address.Collide(range) != ExcelAddress.eAddressCollition.No)
					throw (new ArgumentException(string.Format("Table range collides with table {0}", t.Name)));
			}
			return Add(new ExcelPivotTable(myWorksheet, range, source, name, myWorksheet.Workbook.NextPivotTableID++));
		}

		/// <summary>
		/// Gets the name of the new pivot table.
		/// </summary>
		/// <returns>The new name.</returns>
		internal string GetNewTableName()
		{
			string name = "PivotTable1";
			int i = 2;
			while (myWorksheet.Workbook.ExistsPivotTableName(name))
			{
				name = string.Format("PivotTable{0}", i++);
			}
			return name;
		}

		/// <summary>
		/// Get the enumerator.
		/// </summary>
		/// <returns>The pivot table enumerator.</returns>
		public IEnumerator<ExcelPivotTable> GetEnumerator()
		{
			return myPivotTables.GetEnumerator();
		}

		System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
		{
			return myPivotTables.GetEnumerator();
		}
		#endregion

		#region Private Methods
		private ExcelPivotTable Add(ExcelPivotTable tbl)
		{
			myPivotTables.Add(tbl);
			myPivotTableNames.Add(tbl.Name, myPivotTables.Count - 1);
			if (tbl.CacheID >= myWorksheet.Workbook.NextPivotTableID)
				myWorksheet.Workbook.NextPivotTableID = tbl.CacheID + 1;
			return tbl;
		}
		#endregion
	}
}