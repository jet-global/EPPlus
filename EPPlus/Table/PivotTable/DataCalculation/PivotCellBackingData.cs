/*******************************************************************************
* You may amend and distribute as you like, but don't remove this header!
*
* EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
* See http://www.codeplex.com/EPPlus for details.
*
* Copyright (C) 2011-2019 Evan Schallerer and others as noted in the source history.
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
using System.Collections.Generic;

namespace OfficeOpenXml.Table.PivotTable.DataCalculation
{
	/// <summary>
	/// Contains data that backs a pivot table cell.
	/// </summary>
	internal class PivotCellBackingData
	{
		#region Properties
		/// <summary>
		/// Gets whether or not this contains backing data for a calculated cell.
		/// </summary>
		public bool IsCalculatedCell { get; }

		/// <summary>
		/// Gets or sets the formula if this data is backing a calculated field cell.
		/// </summary>
		public string Formula { get; set; }

		/// <summary>
		/// Gets or sets the result of evaluating the backing data.
		/// </summary>
		public object Result { get; set; }

		/// <summary>
		/// Gets or sets the row that this object corresponds to on the pivot table's worksheet.
		/// </summary>
		public int SheetRow { get; set; } = -1;

		/// <summary>
		/// Gets or sets the column that this object corresponds to on the pivot table's worksheet.
		/// </summary>
		public int SheetColumn { get; set; } = -1;

		/// <summary>
		/// Gets or sets the index into the data field collection that this object corresponds to.
		/// </summary>
		public int DataFieldCollectionIndex { get; set; } = -1;

		/// <summary>
		/// Gets or sets the index of this backing cell data in the pivot table.
		/// Corresponds to the index of the major axis header for this cell.
		/// </summary>
		public int MajorAxisIndex { get; set; }

		/// <summary>
		/// Gets or sets a value indicating whether or not the value of this 
		/// backing data should be shown. 
		/// </summary>
		public bool ShowValue { get; set; }

		private Dictionary<string, List<object>> CalculatedCellBackingData { get; set; }

		private List<object> BackingData { get; set; }
		#endregion

		#region Constructors
		/// <summary>
		/// Instantiates a new <see cref="PivotCellBackingData"/> for a non-calcuated cell.
		/// </summary>
		/// <param name="values">The backing values.</param>
		public PivotCellBackingData(List<object> values)
		{
			this.IsCalculatedCell = false;
			this.BackingData = values;
		}

		/// <summary>
		/// Instantiates a new <see cref="PivotCellBackingData"/> for a calcuated cell.
		/// </summary>
		/// <param name="cacheFieldNameToValues">The backing values for a calculated cell.</param>
		/// <param name="formula">The calculated cell's formula.</param>
		public PivotCellBackingData(Dictionary<string, List<object>> cacheFieldNameToValues, string formula)
		{
			this.IsCalculatedCell = true;
			this.CalculatedCellBackingData = cacheFieldNameToValues;
			this.Formula = formula;
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Gets the backing data values for a non-calculated cell. 
		/// Throws an exception if this is a calculated cell.
		/// </summary>
		/// <returns>The object list of backing data.</returns>
		public List<object> GetBackingValues()
		{
			if (this.IsCalculatedCell)
				throw new InvalidOperationException("Backing values can only be retrieved from non-calculated cells.");
			return this.BackingData;
		}

		/// <summary>
		/// Gets the backing data values mapped to cache field names for a calculated cell.
		/// Throws an exception if this is a non-calculated cell.
		/// </summary>
		/// <returns>The dictionary of field names to object lists of backing data.</returns>
		public Dictionary<string, List<object>> GetCalculatedCellBackingValues()
		{
			if (!this.IsCalculatedCell)
				throw new InvalidOperationException("Calculated cell backing values can only be retrieved from calculated cells.");
			return this.CalculatedCellBackingData;
		}

		/// <summary>
		/// Merges the data from the specified <paramref name="cellBackingData"/> into this <see cref="PivotCellBackingData"/>.
		/// Both must be either calculated or non-calculated backing datas.
		/// </summary>
		/// <param name="cellBackingData">The backing data to merge.</param>
		public void Merge(PivotCellBackingData cellBackingData)
		{
			if (this.IsCalculatedCell != cellBackingData.IsCalculatedCell)
				throw new InvalidOperationException($"Cannot merge {nameof(PivotCellBackingData)} of different types.");
			if (this.IsCalculatedCell)
			{
				if (cellBackingData.CalculatedCellBackingData != null)
				{
					if (this.CalculatedCellBackingData == null)
						this.CalculatedCellBackingData = new Dictionary<string, List<object>>();
					foreach (var fieldName in cellBackingData.CalculatedCellBackingData.Keys)
					{
						if (!this.CalculatedCellBackingData.ContainsKey(fieldName))
							this.CalculatedCellBackingData.Add(fieldName, cellBackingData.CalculatedCellBackingData[fieldName]);
						else
							this.CalculatedCellBackingData[fieldName].AddRange(cellBackingData.CalculatedCellBackingData[fieldName]);
					}
				}
			}
			else
			{
				if (cellBackingData.BackingData != null)
				{
					if (this.BackingData == null)
						this.BackingData = new List<object>();
					this.BackingData.AddRange(cellBackingData.BackingData);
				}
			}
		}

		/// <summary>
		/// Clones the current <see cref="PivotCellBackingData"/>.
		/// </summary>
		/// <returns>A copy of the current <see cref="PivotCellBackingData"/> object.</returns>
		public PivotCellBackingData Clone()
		{
			PivotCellBackingData backingData = null;
			if (this.IsCalculatedCell)
			{
				var valuesDictionary = new Dictionary<string, List<object>>();
				foreach (var keyValue in this.CalculatedCellBackingData)
				{
					valuesDictionary.Add(keyValue.Key, new List<object>(keyValue.Value));
				}
				backingData = new PivotCellBackingData(valuesDictionary, this.Formula);
			}
			else
			{
				backingData = new PivotCellBackingData(new List<object>(this.BackingData));
			}
			backingData.ShowValue = this.ShowValue;
			return backingData;
		}
		#endregion
	}
}
