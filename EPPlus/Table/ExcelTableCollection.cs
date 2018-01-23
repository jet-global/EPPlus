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
 * Jan Källman		Added		30-AUG-2010
 * Jan Källman		License changed GPL-->LGPL 2011-12-16
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.Table
{
	/// <summary>
	/// A collection of table objects.
	/// </summary>
	public class ExcelTableCollection : IEnumerable<ExcelTable>
	{
		#region Properties
		/// <summary>
		/// Gets the number of tables in this collection.
		/// </summary>
		public int Count
		{
			get
			{
				return Tables.Count;
			}
		}

		internal Dictionary<string, int> TableNames { get; } = new Dictionary<string, int>(StringComparer.InvariantCultureIgnoreCase);

		private List<ExcelTable> Tables { get; } = new List<ExcelTable>();

		private ExcelWorksheet Worksheet { get; }
		#endregion

		#region Constructors
		/// <summary>
		/// Instantiates a new <see cref="ExcelTableCollection"/> with the tables from the specified <paramref name="worksheet"/>.
		/// </summary>
		/// <param name="worksheet">The worksheet whose tables should be represented in this collection.</param>
		internal ExcelTableCollection(ExcelWorksheet worksheet)
		{
			var pck = worksheet.Package.Package;
			this.Worksheet = worksheet;
			foreach (XmlElement node in worksheet.WorksheetXml.SelectNodes("//d:tableParts/d:tablePart", worksheet.NameSpaceManager))
			{
				var relationship = worksheet.Part.GetRelationship(node.GetAttribute("id", ExcelPackage.schemaRelationships));
				var table = new ExcelTable(relationship, worksheet);
				this.TableNames.Add(table.Name, Tables.Count);
				this.Tables.Add(table);
			}
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Create a table based on the specified range.
		/// </summary>
		/// <param name="range">The range address, including header and total row.</param>
		/// <param name="tableName">The name of the table. Must be unique. If none is provided, a default table name will be applied.</param>
		/// <returns>The <see cref="ExcelTable"/> object that represents the table.</returns>
		public ExcelTable Add(ExcelAddress range, string tableName = null)
		{
			if (range.WorkSheet != null && range.WorkSheet != this.Worksheet.Name)
			{
				throw new ArgumentException("Range does not belong to worksheet", "Range");
			}

			if (string.IsNullOrEmpty(tableName))
			{
				tableName = this.GetNewTableName();
			}
			else if (this.Worksheet.Workbook.ExistsTableName(tableName))
			{
				throw (new ArgumentException("Tablename is not unique"));
			}

			this.ValidateTableName(tableName);

			foreach (var t in Tables)
			{
				if (t.Address.Collide(range) != ExcelAddress.eAddressCollition.No)
				{
					throw (new ArgumentException(string.Format("Table range collides with table {0}", t.Name)));
				}
			}
			return this.Add(new ExcelTable(this.Worksheet, range, tableName, this.Worksheet.Workbook.NextTableID));
		}

		/// <summary>
		/// Delete the table at the specified <paramref name="index"/>.
		/// </summary>
		/// <param name="index">The index of the table to delete.</param>
		/// <param name="clearRange">Whether or not the contents of the table should be deleted from the worksheet.</param>
		public void Delete(int index, bool clearRange = false)
		{
			this.Delete(this[index], clearRange);
		}

		/// <summary>
		/// Deletes the table with the specified <paramref name="name"/>.
		/// </summary>
		/// <param name="name">The name of the <see cref="ExcelTable"/> to delete.</param>
		/// <param name="clearRange">Whether or not the contents of the table should be deleted from the worksheet.</param>
		public void Delete(string name, bool clearRange = false)
		{
			if (this[name] == null)
			{
				throw new ArgumentOutOfRangeException(string.Format("Cannot delete non-existant table {0} in sheet {1}.", name, this.Worksheet.Name));
			}
			this.Delete(this[name], clearRange);
		}

		/// <summary>
		/// Deletes the specified <paramref name="table"/>.
		/// </summary>
		/// <param name="table">The <see cref="ExcelTable"/> to delete.</param>
		/// <param name="clearRange">Whether or not the contents of the table should be removed from the worksheet.</param>
		public void Delete(ExcelTable table, bool clearRange = false)
		{
			if (!this.Tables.Contains(table))
			{
				throw new ArgumentOutOfRangeException("Table", String.Format("Table {0} does not exist in this collection", table.Name));
			}
			lock (this)
			{
				var range = Worksheet.Cells[table.Address.Address];
				this.TableNames.Remove(table.Name);
				this.Tables.Remove(table);
				if (table.TableUri != null && table.WorkSheet.Package.Package.PartExists(table.TableUri))
					table.WorkSheet.Package.Package.DeletePart(table.TableUri);
				var nodeToRemove = table.WorkSheet.WorksheetXml.SelectSingleNode($"//d:tableParts/d:tablePart[@r:id=\"{table.RelationshipID}\"]", table.WorkSheet.NameSpaceManager);
				if (nodeToRemove != null)
					nodeToRemove.ParentNode.RemoveChild(nodeToRemove);
				foreach (var sheet in table.WorkSheet.Workbook.Worksheets)
				{
					foreach (var nextTable in sheet.Tables)
					{
						if (nextTable.Id > table.Id) nextTable.Id--;
					}
					table.WorkSheet.Workbook.NextTableID--;
				}
				if (clearRange)
				{
					range.Clear();
				}
			}
		}

		/// <summary>
		/// Get the table object from a range.
		/// </summary>
		/// <param name="range">The range</param>
		/// <returns>The table. Null if no range matches</returns>
		public ExcelTable GetFromRange(ExcelRangeBase range)
		{
			foreach (var tbl in range.Worksheet.Tables)
			{
				if (tbl.Address._address == range._address)
				{
					return tbl;
				}
			}
			return null;
		}

		/// <summary>
		/// Gets the table at the specified 0-based index.
		/// Throws an <see cref="ArgumentOutOfRangeException"/> if there is no table at the specified index.
		/// </summary>
		/// <param name="index">The 0-based index of the table to retrieve.</param>
		/// <returns>The <see cref="ExcelTable"/> at the specified index.</returns>
		public ExcelTable this[int index]
		{
			get
			{
				if (index < 0 || index >= this.Tables.Count)
				{
					throw (new ArgumentOutOfRangeException("Table index out of range"));
				}
				return this.Tables[index];
			}
		}

		/// <summary>
		/// Gets a table from this collection by name.
		/// </summary>
		/// <param name="name">The name of the table</param>
		/// <returns>The table. Null if the table name is not found in the collection</returns>
		public ExcelTable this[string name]
		{
			get
			{
				if (this.TableNames.ContainsKey(name))
				{
					return this.Tables[this.TableNames[name]];
				}
				else
				{
					return null;
				}
			}
		}
		#endregion

		#region Internal Methods
		/// <summary>
		/// Gets the next valid table name.
		/// </summary>
		/// <returns>A valid table name, of the form 'Table[N]'.</returns>
		internal string GetNewTableName()
		{
			string name = "Table1";
			int i = 2;
			while (this.Worksheet.Workbook.ExistsTableName(name))
			{
				name = string.Format("Table{0}", i++);
			}
			return name;
		}
		#endregion

		#region Private Methods
		private ExcelTable Add(ExcelTable table)
		{
			this.Tables.Add(table);
			this.TableNames.Add(table.Name, this.Tables.Count - 1);
			if (table.Id >= this.Worksheet.Workbook.NextTableID)
			{
				this.Worksheet.Workbook.NextTableID = table.Id + 1;
			}
			return table;
		}

		private void ValidateTableName(string name)
		{
			if (string.IsNullOrEmpty(name))
			{
				throw new ArgumentException("Tablename is null or empty");
			}

			char firstLetterOfName = name[0];
			if (Char.IsLetter(firstLetterOfName) == false && firstLetterOfName != '_' && firstLetterOfName != '\\')
			{
				throw new ArgumentException("Tablename start with invalid character");
			}

			if (name.Contains(" "))
			{
				throw new ArgumentException("Tablename has spaces");
			}
		}
		#endregion

		#region IEnumerable Members
		public IEnumerator<ExcelTable> GetEnumerator()
		{
			return this.Tables.GetEnumerator();
		}

		System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
		{
			return this.Tables.GetEnumerator();
		}
		#endregion
	}
}
