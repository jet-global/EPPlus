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
	/// Abstract base collection class for pivot table fields.
	/// </summary>
	/// <typeparam name="T">An instance of {T}.</typeparam>
	public abstract class ExcelPivotTableFieldCollectionBase<T> : IEnumerable<T>
	{
		#region Class Variables
		/// <summary>
		/// The pivot table.
		/// </summary>
		protected ExcelPivotTable myTable;

		/// <summary>
		/// A list of fields.
		/// </summary>
		internal List<T> myList = new List<T>();
		#endregion

		#region Properties
		/// <summary>
		/// Gets the generic enumerator of the list.
		/// </summary>
		/// <returns>The enumerator.</returns>
		public IEnumerator<T> GetEnumerator()
		{
			return myList.GetEnumerator();
		}

		/// <summary>
		/// Gets the specified type enumerator of the list.
		/// </summary>
		/// <returns>The enumerator.</returns>
		System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
		{
			return myList.GetEnumerator();
		}

		/// <summary>
		/// Gets the count of fields in the pivot table.
		/// </summary>
		public int Count
		{
			get
			{
				return myList.Count;
			}
		}

		/// <summary>
		/// Gets the field at the given index.
		/// </summary>
		/// <param name="Index">The position of the field in the list.</param>
		/// <returns>The pivot table field.</returns>
		public T this[int Index]
		{
			get
			{
				if (Index < 0 || Index >= myList.Count)
					throw (new ArgumentOutOfRangeException("Index out of range"));
				return myList[Index];
			}
		}
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="ExcelPivotTableFieldCollection"/>.
		/// </summary>
		/// <param name="table">The existing pivot table.</param>
		internal ExcelPivotTableFieldCollectionBase(ExcelPivotTable table)
		{
			myTable = table;
		}
		#endregion

		#region Methods
		/// <summary>
		/// Adds a field to the collection.
		/// </summary>
		/// <param name="field"></param>
		internal void AddInternal(T field)
		{
			myList.Add(field);
		}

		/// <summary>
		/// Clears the field collection.
		/// </summary>
		internal void Clear()
		{
			myList.Clear();
		}
		#endregion
	}
}