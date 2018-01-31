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
 * Jan Källman		Added this class		        2010-01-28
 * Jan Källman		License changed GPL-->LGPL 2011-12-27
 *******************************************************************************/
using System;
using System.Collections;
using System.Collections.Generic;

namespace OfficeOpenXml
{
	/// <summary>
	/// Collection for named ranges
	/// </summary>
	public class ExcelNamedRangeCollection : IEnumerable<ExcelNamedRange>
	{
		#region Properties
		/// <summary>
		/// The current number of named ranges in the collection.
		/// </summary>
		public int Count
		{
			get
			{
				return this.NamedRanges.Count;
			}
		}

		/// <summary>
		/// Gets the <see cref="ExcelNamedRange"/> associated with the specified <paramref name="name"/>.
		/// </summary>
		/// <param name="name">The name of a named range to retrieve.</param>
		/// <returns>A reference to the named range found.</returns>
		/// <exception cref="KeyNotFoundException">Thrown if a named range with the specified <paramref name="name"/> was not in the collection.</exception>
		public ExcelNamedRange this[string name]
		{
			get
			{
				return this.NamedRanges[name];
			}
		}

		private ExcelWorkbook Workbook { get; }
		private ExcelWorksheet Worksheet { get; }
		private Dictionary<string, ExcelNamedRange> NamedRanges { get; } = new Dictionary<string, ExcelNamedRange>(StringComparer.InvariantCultureIgnoreCase);
		#endregion

		#region Constructors
		/// <summary>
		/// Instantiates a new <see cref="ExcelNamedRangeCollection"/> for the specified <paramref name="workbook"/>.
		/// </summary>
		/// <param name="workbook">The <see cref="ExcelWorkbook"/> that this named range collection is scoped to.</param>
		internal ExcelNamedRangeCollection(ExcelWorkbook workbook)
		{
			this.Workbook = workbook;
			this.Worksheet = null;
		}

		/// <summary>
		/// Instantiates a new <see cref="ExcelNamedRangeCollection"/> for the specified <paramref name="workbook"/>
		/// and <paramref name="worksheet"/>.
		/// </summary>
		/// <param name="workbook">The <see cref="ExcelWorkbook"/> that this named range collection belongs to.</param>
		/// <param name="worksheet">The <see cref="ExcelWorksheet"/> that this named range collection is scoped to.</param>
		internal ExcelNamedRangeCollection(ExcelWorkbook workbook, ExcelWorksheet worksheet)
		{
			this.Workbook = workbook;
			this.Worksheet = worksheet;
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Adds a new named range to the collection.
		/// </summary>
		/// <param name="name">The name of the named range to add.</param>
		/// <param name="range">The range that the named range references.</param>
		/// <returns>The named range that was added to the collection.</returns>
		public ExcelNamedRange Add(string name, ExcelRangeBase range)
		{
			if (range == null)
				throw new ArgumentNullException(nameof(range));
			ExcelNamedRange namedRange = new ExcelNamedRange(name, this.Workbook, this.Worksheet, range.FullAddress, this.NamedRanges.Count);
			this.NamedRanges.Add(name, namedRange);
			return namedRange;
		}

		/// <summary>
		/// Adds a new named range to the collection.
		/// </summary>
		/// <param name="name">The name of the named range.</param>
		/// <param name="formula">The formula of the named range.</param>
		/// <param name="isHidden">A value indicating whether or not the hidden attribute was set on the named range (optional).</param>
		/// <param name="comments">The comments set on the named range (optional).</param>
		/// <returns>The newly-added named named range.</returns>
		public ExcelNamedRange Add(string name, string formula, bool isHidden = false, string comments = null)
		{
			if (string.IsNullOrEmpty(name))
				throw new ArgumentNullException(nameof(name));
			if (string.IsNullOrEmpty(formula))
				throw new ArgumentNullException(nameof(formula));
			var namedRange = new ExcelNamedRange(name, this.Workbook, this.Worksheet, formula, this.NamedRanges.Count)
			{
				IsNameHidden = isHidden,
				NameComment = comments
			};
			this.NamedRanges.Add(name, namedRange);
			return namedRange;
		}

		/// <summary>
		/// Removes an <see cref="ExcelNamedRange"/> with the specified <paramref name="name"/> from the collection 
		/// if it exists.
		/// </summary>
		/// <param name="name">The name of the <see cref="ExcelNamedRange"/> to remove.</param>
		/// <returns>True if a named range with the specified <paramref name="name"/> exists, false otherwise.</returns>
		public bool Remove(string name)
		{
			return this.NamedRanges.Remove(name);
		}

		/// <summary>
		/// Checks collection for the presence of a named range with the specified <paramref name="name"/>.
		/// </summary>
		/// <param name="name">The name of the named range.</param>
		/// <returns>True if the named range is in the collection.</returns>
		public bool ContainsKey(string name)
		{
			return this.NamedRanges.ContainsKey(name);
		}
		#endregion

		#region Internal Methods
		/// <summary>
		/// Updates all named range formulas in this collection to reflect the specified
		/// row and/or column insertion on the specified <paramref name="worksheet"/>.
		/// </summary>
		/// <param name="rowFrom">The row at which the insert starts.</param>
		/// <param name="colFrom">The column at which the insert starts.</param>
		/// <param name="rows">The number of rows to insert.</param>
		/// <param name="cols">The number of columns to insert.</param>
		/// <param name="worksheet">The worksheet on which the insert is occurring.</param>
		internal void Insert(int rowFrom, int colFrom, int rows, int cols, ExcelWorksheet worksheet)
		{
			foreach (var namedRange in this)
			{
				namedRange.UpdateFormula(rowFrom, colFrom, rows, cols, worksheet);
			}
		}

		/// <summary>
		/// Updates all named range formulas in this collection to reflect the specified
		/// row and/or column deletion on the specified <paramref name="worksheet"/>.
		/// </summary>
		/// <param name="rowFrom">The row at which the delete starts.</param>
		/// <param name="colFrom">The column at which the delete starts.</param>
		/// <param name="rows">The number of rows to delete.</param>
		/// <param name="cols">The number of columns to delete.</param>
		/// <param name="worksheet">The worksheet on which the delete is occurring.</param>
		internal void Delete(int rowFrom, int colFrom, int rows, int cols, ExcelWorksheet worksheet)
		{
			foreach (var namedRange in this)
			{
				namedRange.UpdateFormula(rowFrom, colFrom, -rows, -cols, worksheet);
			}
		}
		#endregion

		#region IEnumerable Overrides
		#region IEnumerable<ExcelNamedRange> Members
		/// <summary>
		/// Implement interface method IEnumerator&lt;ExcelNamedRange&gt; GetEnumerator().
		/// </summary>
		/// <returns>An enumerator of <see cref="ExcelNamedRange"/>s for the current collection.</returns>
		public IEnumerator<ExcelNamedRange> GetEnumerator()
		{
			return this.NamedRanges.Values.GetEnumerator();
		}
		#endregion

		#region IEnumerable Members
		/// <summary>
		/// Implement interface method IEnumeratable GetEnumerator().
		/// </summary>
		/// <returns>An enumerator for the current collection.</returns>
		IEnumerator IEnumerable.GetEnumerator()
		{
			return this.NamedRanges.Values.GetEnumerator();
		}
		#endregion
		#endregion
	}
}
