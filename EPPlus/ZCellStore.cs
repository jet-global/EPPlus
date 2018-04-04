/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * EPPlus provides server-side generation of Excel 2007/2010 spreadsheets.
 * See http://www.codeplex.com/EPPlus for details.
 *
 * Copyright (C) 2017  Zachary Faltersack
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
 * Author							   Change																				Date
 * ******************************************************************************
 * Zachary Faltersack    Added       																	2017-12-22
 * Zachary Faltersack    Complete work, add comments       		        2018-02-07
 *******************************************************************************
 * 
 * NOTES:
 * There is more optimization that can be done but for now this appears to be stable.
 * I think in particular that enumerating can be sped up. 
 * I also think that removing empty pages when clearing/shifting in PagedStructure might be beneficial.
 * 
 *******************************************************************************/

using System;
using System.Collections;
using System.Collections.Generic;

namespace OfficeOpenXml
{
	/// <summary>
	/// Understanding the ZCellStore
	/// Author: Zachary Faltersack
	/// 
	/// The original cellstore was very complicated and contained some bad paging issues.
	/// This is my implementation. The behavior is my best understanding of how the original cellstore was intended to work.
	/// The following notes describe some invariants I've come to believe about the usage of the cellstore throughout EPPlus.
	/// 
	/// Excel uses 1-based indexing for cells, so A1 = [1,1].
	/// In the cellstore:
	/// The 0 index for row is used to store Column Metadata.
	/// The 0 index for column is used to store Row Metadata.
	/// [0,0] is NOT considered valid as far as I can tell.
	/// 
	/// 
	/// 		Assume the current invariant (where X is at [0,0]):
	/// 		X : Invalid coordinate.
	/// 		C : Column metadata.
	/// 		R : Row metadata.
	/// 		* : Cell content.
	/// 		
	/// 		         XCCCCCCCC
	/// 		         R********
	/// 		         R********
	/// 		         R********
	/// 		         R********
	/// 		         R********
	/// 		         R********
	/// 		         R********
	/// 		         R********
	/// 
	/// ZCellStore uses 3 instances of a PagedStructure to store information (see below for more information on PagedStructure)
	/// 
	/// ColumnData [<![CDATA[PagedStructure<T>]]>] : This contains the metadata for the columns in the ZCellStore
	/// RowData [<![CDATA[PagedStructure<T>]]>] : This contains the metadata for the rows in the ZCellStore
	/// CellData [<![CDATA[PagedStructure<PagedStructure<T>>]]>] : This contains all of the cell data. The outer PagedStructure maps to columns
	/// and the inner PagedStructure maps to rows.
	/// 
	/// Given that PagedStructure is 0-index, and the 0 indices have a special meaning for the ZCellStore, we need to mindfully
	/// handle accessing data when retrieving/editing/enumerating values in the ZCellStore.
	/// 
	/// If the ZCellStore receives a request to access data in row 0, then we know we need to look at the ColumnData structure.
	/// If the ZCellStore receives a request to access data in column 0, then we know we need to look at the RowData structure.
	/// ALL other coordinates must be reduced by 1 (to accomodate the 0-indexing) and then we need to look in the CellData structure.
	/// 
	/// This is mostly straightfoward, but provides some interesting complexity when enumerating the ZCellStore.
	/// 
	/// Before attempting to understand the <see cref="ZCellStore{T}"/> it is highly recommended to review the notes for
	/// the <see cref="PagedStructure{S}"/>.
	/// </summary>
	/// <typeparam name="T">The type this ZCellStore is going to contain.</typeparam>
	internal class ZCellStore<T> : ICellStore<T>
	{
		#region Constants
		/*
		 * Default Row Page Bits results in an index range of: 1..1048576
		 * Default Column Page Bits results in an index range of: 1..16384
		 * 
		 * These values represent the number of bits to use that provide row and column ranges that match Excel.
		 * When using them as a right-shift then you will get the page index that contains the coordinate that was requested.
		 * When using them as a bit-mask you will get the index into the page that contains the coordinate that was requested.
		 * 
		 * In this case:
		 * DefaultRowPageBits    : 10 -- yields :: 2^10 = 1024 ; 1024^2 = 1048576 [the max row for Excel]
		 * DefaultColumnPageBits :  7 -- yields :: 2^7  =  128 ;  128^2 = 16384   [the max column for Excel]
		 * 
		 * See PagedStructure for how these numbers are actually used.
		 * */
		private const int DefaultRowPageBits = 10;
		private const int DefaultColumnPageBits = 7;
		#endregion

		#region Properties
		private PagedStructure<PagedStructure<T>> CellData { get; } 
		private PagedStructure<T> ColumnData { get; }
		private PagedStructure<T> RowData { get; }
		private int RowPageBits { get; }
		private int ColumnPageBits { get; }
		#endregion

		#region Constructors
		/// <summary>
		/// Creates a default <see cref="ZCellStore{T}"/>.
		/// </summary>
		public ZCellStore() : this(ZCellStore<T>.DefaultRowPageBits, ZCellStore<T>.DefaultColumnPageBits) { }

		/// <summary>
		/// Creates a <see cref="ZCellStore{T}"/> with custom dimensions.
		/// </summary>
		/// <param name="rowPageBits">Half the number of bits to use for representing row coordinates.</param>
		/// <param name="columnPageBits">Half the number of bits to use for representing column coordinates.</param>
		public ZCellStore(int rowPageBits, int columnPageBits)
		{
			this.RowPageBits = rowPageBits;
			this.ColumnPageBits = columnPageBits;
			this.CellData = new PagedStructure<PagedStructure<T>>(this.ColumnPageBits);
			this.ColumnData = new PagedStructure<T>(this.ColumnPageBits);
			this.RowData = new PagedStructure<T>(this.RowPageBits);
			// We need to accomodate 1-Indexing for Excel
			this.MaximumRow = this.RowData.MaximumIndex + 1;
			this.MaximumColumn = this.ColumnData.MaximumIndex + 1;
		}
		#endregion

		#region ICellStore<T> Members
		/// <summary>
		/// Gets the maximum row of the cellstore.
		/// </summary>
		public int MaximumRow { get; }

		/// <summary>
		/// Gets the maximum column of the cellstore.
		/// </summary>
		public int MaximumColumn { get; }

		/// <summary>
		/// Get the value at a particular location.
		/// </summary>
		/// <param name="row">The row to read a value from.</param>
		/// <param name="column">The column to read a value from.</param>
		/// <returns>The value in the cell coordinate.</returns>
		public T GetValue(int row, int column)
		{
			// Row 0 means accessing column metadata
			if (row == 0)
			{
				var item = this.ColumnData.GetItem(column - 1);
				if (item.HasValue)
					return item.Value;
			}
			// Column 0 means accessing row metadata
			else if (column == 0)
			{
				var item = this.RowData.GetItem(row - 1);
				if (item.HasValue)
					return item.Value;
			}
			// Since we're in the main sheet, the row and column are reduced by 1
			else if (this.TryUpdateIndices(ref row, ref column))
			{
				var columnItem = this.CellData.GetItem(column)?.Value;
				if (columnItem != null)
				{
					var rowItem = columnItem.GetItem(row);
					if (rowItem.HasValue)
						return rowItem.Value.Value;
				}
			}
			return default(T);
		}

		/// <summary>
		/// Set the value at a particular location.
		/// </summary>
		/// <param name="row">The row of the location to set a value at.</param>
		/// <param name="column">The column of the location to set a value at.</param>
		/// <param name="value">The value to store at the location.</param>
		public void SetValue(int row, int column, T value)
		{
			// Row 0 means accessing column metadata
			if (row == 0)
				this.ColumnData.SetItem(column - 1, value);
			// Column 0 means accessing row metadata
			else if (column == 0)
				this.RowData.SetItem(row - 1, value);
			// Since we're in the main sheet, the row and column are reduced by 1
			else if (this.TryUpdateIndices(ref row, ref column))
			{
				var columnStructure = this.CellData.GetItem(column);
				if (columnStructure == null)
				{
					columnStructure = new PagedStructure<T>(this.RowPageBits);
					this.CellData.SetItem(column, columnStructure);
				}
				columnStructure?.Value.SetItem(row, value);
			}
		}

		/// <summary>
		/// Determine if a value exists at the given location.
		/// </summary>
		/// <param name="row">The row to look for a value at.</param>
		/// <param name="column">The column to look for a value at.</param>
		/// <returns>True if a value was found at the location; false otherwise.</returns>
		public bool Exists(int row, int column)
		{
			return this.Exists(row, column, out _);
		}

		/// <summary>
		/// Determine if a value exists at the given location, and return it as an out parameter if found.
		/// </summary>
		/// <param name="row">The row to look for a value at.</param>
		/// <param name="column">The column to look for a value at.</param>
		/// <param name="value">The value found, if one exists.</param>
		/// <returns>True if a value was found at the location; false otherwise.</returns>
		public bool Exists(int row, int column, out T value)
		{
			value = default(T);
			// Row 0 means accessing column metadata
			if (row == 0)
			{
				var item = this.ColumnData.GetItem(column - 1);
				if (item.HasValue)
				{
					value = item.Value.Value;
					return true;
				}
			}
			// Column 0 means accessing row metadata
			else if (column == 0)
			{
				var item = this.RowData.GetItem(row - 1);
				if (item.HasValue)
				{
					value = item.Value.Value;
					return true;
				}
			}
			// Since we're in the main sheet, the row and column are reduced by 1
			else if (this.TryUpdateIndices(ref row, ref column))
			{
				var item = this.CellData.GetItem(column)?.Value.GetItem(row);
				if (item.HasValue)
				{
					value = item.Value.Value;
					return true;
				}
			}
			return false;
		}

		/// <summary>
		/// Get the location of the next cell after the given row and column, if one exists.
		/// Relies on the assumption that cells are sorted by column and then sorted within each column by row.
		/// </summary>
		/// <param name="row">The row to start searching from, and also the new location's row, if one exists.</param>
		/// <param name="column">The column to start searching from, and also the new location's column, if one exists.</param>
		/// <returns>True if a next cell has been found and the row and column parameters have been updated; false otherwise.</returns>
		public bool NextCell(ref int row, ref int column)
		{
			return this.NextCellBound(ref row, ref column, 0, this.MaximumRow, 0, this.MaximumColumn);
		}

		/// <summary>
		/// Get the location of the first cell before the given row and column, if one exists.
		/// Relies on the assumption that cells are sorted by column and then sorted within each column by row.
		/// </summary>
		/// <param name="row">The row to start searching from, and also the new location's row, if one exists.</param>
		/// <param name="column">The column to start searching from, and also the new location's column, if one exists.</param>
		/// <returns>True if a previous cell has been found and the row and column parameters have been updated; false otherwise.</returns>
		public bool PrevCell(ref int row, ref int column)
		{
			int columnSearch, rowSearch;
			// tests for this
			if (row > this.MaximumRow)
				row = this.MaximumRow;
			if (column > this.MaximumColumn + 1)
				column = this.MaximumColumn + 1;
			if (column <= 0)
			{
				column = this.MaximumColumn + 1;
				row--;
			}
			if (row == 0)
			{
				columnSearch = column - 1; // 1-indexing
				if (this.ColumnData.PreviousItem(ref columnSearch))
				{
					column = columnSearch + 1;
					row = 0;
					return true;
				}
				return false;
			}
			columnSearch = column - 1;
			// STEP 1:: Finish searching current row
			while (this.CellData.PreviousItem(ref columnSearch))
			{
				var page = this.CellData.GetItem(columnSearch);
				var item = page?.Value.GetItem(row - 1);
				if (item.HasValue)
				{
					column = columnSearch + 1;
					return true;
				}
			}
			// STEP 2:: Check row metadata
			if (this.RowData.GetItem(row - 1).HasValue)
			{
				column = 0;
				return true;
			}
			column = this.MaximumColumn + 1;
			// STEP 3:: Search the rest of the sheet
			int currentRow = int.MinValue;
			int currentColumn = int.MinValue;
			columnSearch = column;
			while (this.CellData.PreviousItem(ref columnSearch))
			{
				var page = this.CellData.GetItem(columnSearch);
				rowSearch = row - 1; // no offset is applied because we're already off by one (too high due to 1-indexing)
				if (page?.Value.PreviousItem(ref rowSearch) == true)
				{
					if (rowSearch + 1 > currentRow)
					{
						currentRow = rowSearch + 1;
						currentColumn = columnSearch + 1;
					}
				}
			}
			rowSearch = row - 1; // no offset is applied because we're already off by one (too high due to 1-indexing)
			if (this.RowData.PreviousItem(ref rowSearch))
			{
				if (rowSearch + 1 > currentRow)
				{
					currentRow = rowSearch + 1;
					currentColumn = 0;
				}
			}
			row = currentRow;
			column = currentColumn;
			if (row == int.MinValue && column == int.MinValue)
			{
				columnSearch = column - 1; // 1-indexing
				if (this.ColumnData.PreviousItem(ref columnSearch))
				{
					column = columnSearch + 1;
					row = 0;
					return true;
				}
				return false;
			}
			return true;
		}

		/// <summary>
		/// Deletes rows and/or columns from the workbook. This deletes all existing nodes in the specified range, and updates the keys of all subsequent nodes to reflect their new positions after being shifted to fill the newly-vacated space.
		/// </summary>
		/// <param name="fromRow">The first row to delete.</param>
		/// <param name="fromCol">The first column to delete.</param>
		/// <param name="rows">The number of rows to delete.</param>
		/// <param name="columns">The number of columns to delete.</param>
		public void Delete(int fromRow, int fromCol, int rows, int columns)
		{
			if (fromRow != 0 && fromCol != 0)
				throw new InvalidOperationException("Only delete rows or columns in a single operation.");
			// Invariant: The actual row or column to be deleting from will not be 0.
			//						Only valid coordinates are passed in. A 0 is used to disambiguate between deleting full rows 
			//						or full columns.
			fromRow--;
			fromCol--;
			if (fromCol >= 0)
			{
				this.CellData.ShiftItems(fromCol, -columns);
				this.ColumnData.ShiftItems(fromCol, -columns);
			}
			else
			{
				int column = -1;
				while (this.CellData.NextItem(ref column))
				{
					this.CellData.GetItem(column)?.Value.ShiftItems(fromRow, -rows);
				}
				this.RowData.ShiftItems(fromRow, -rows);
			}
		}

		/// <summary>
		/// Removes the values in the specified range without updating cells below or to the right of the specified range.
		/// </summary>
		/// <param name="fromRow">The first row whose cells should be cleared.</param>
		/// <param name="fromCol">The first column whose cells should be cleared.</param>
		/// <param name="rows">The number of rows to clear.</param>
		/// <param name="columns">The number of columns to clear.</param>
		public void Clear(int fromRow, int fromCol, int rows, int columns)
		{
			if (fromRow == 0)
			{
				this.ColumnData.ClearItems(fromCol - 1, columns);
				fromRow++;
				rows--;
			}
			if (fromCol == 0)
			{
				this.RowData.ClearItems(fromRow - 1, rows);
				fromCol++;
				columns--;
			}
			fromRow--;
			fromCol--;
			// Clearing full columns
			if (fromRow <= 1 && fromRow + rows >= this.MaximumRow)
			{
				this.CellData.ClearItems(fromCol, columns);
				this.ColumnData.ClearItems(fromCol, columns);
			}
			// Clearing full rows
			else if (fromCol <= 1 && fromCol + columns >= this.MaximumColumn)
			{
				int column = -1;
				while (this.CellData.NextItem(ref column))
				{
					this.CellData.GetItem(column)?.Value.ClearItems(fromRow, rows);
				}
			}
			// Clearing a nested block
			else
			{
				int column = fromCol - 1;
				while (this.CellData.NextItem(ref column) && column < (fromCol + columns))
				{
					this.CellData.GetItem(column)?.Value.ClearItems(fromRow, rows);
				}
			}
		}

		/// <summary>
		/// Get the range of cells contained in this collection.
		/// </summary>
		/// <param name="fromRow">The first row contained in this collection.</param>
		/// <param name="fromCol">The first column contained in this collection.</param>
		/// <param name="toRow">The last row contained in this collection.</param>
		/// <param name="toCol">The last column contained in this collection.</param>
		/// <returns>True if the collection contains at least one cell (and therefore has a dimension); false if the collection contains no cells and thus has no dimension.</returns>
		public bool GetDimension(out int fromRow, out int fromCol, out int toRow, out int toCol)
		{
			fromCol = this.CellData.MinimumUsedIndex + 1;
			toCol = this.CellData.MaximumUsedIndex + 1;
			fromRow = ExcelPackage.MaxRows + 1;
			toRow = 0;
			int searchIndex = -1;
			while (this.CellData.NextItem(ref searchIndex))
			{
				var column = this.CellData.GetItem(searchIndex);
				var localMinRow = column?.Value.MinimumUsedIndex + 1;
				if (localMinRow < fromRow)
					fromRow = localMinRow ?? 0;
				var localMaxRow = column?.Value.MaximumUsedIndex + 1;
				if (localMaxRow > toRow)
					toRow = localMaxRow ?? 0;
			}
			return fromRow != ExcelPackage.MaxRows + 1 && toRow != 0 && fromCol != 0 && toCol != 0;
		}

		/// <summary>
		/// "Insert space" into the cellStore by updating all keys beyond the specified row or column by the specified number of rows or columns.
		/// </summary>
		/// <param name="fromRow">The row to start updating keys from.</param>
		/// <param name="fromCol">The columnn to start updating keys from.</param>
		/// <param name="rows">The number of rows being inserted.</param>
		/// <param name="columns">The number of columns being inserted.</param>
		public void Insert(int fromRow, int fromCol, int rows, int columns)
		{
			if (fromRow != 0 && fromCol != 0)
				throw new InvalidOperationException("Only insert rows or columns in a single operation.");
			// Invariant: The actual row or column to be inserting from will not be 0.
			//						Only valid coordinates are passed in. A 0 is used to disambiguate between inserting full rows 
			//						or full columns.
			fromRow--;
			fromCol--;
			if (fromCol >= 0)
			{
				this.CellData.ShiftItems(fromCol, columns);
				this.ColumnData.ShiftItems(fromCol, columns);
			}
			else
			{
				if (fromRow < 0)
				{
					fromRow = 0;
					rows--;
				}
				int column = -1;
				while (this.CellData.NextItem(ref column))
				{
					this.CellData.GetItem(column)?.Value.ShiftItems(fromRow, rows);
				}
				this.RowData.ShiftItems(fromRow, rows);
			}
		}

		public void Dispose()
		{
			// This will be deleted when the old cellstore is finally removed.
		}

		/// <summary>
		/// Gets a default enumerator for the cellstore.
		/// </summary>
		/// <returns>The default enumerator for the cellstore.</returns>
		public ICellStoreEnumerator<T> GetEnumerator()
		{
			return new ZCellStoreEnumerator(this);
		}

		/// <summary>
		/// Gets a contained enumerator for the cellstore.
		/// </summary>
		/// <param name="startRow">The minimum row to enumerate.</param>
		/// <param name="startColumn">The minimum column to enumerate.</param>
		/// <param name="endRow">The maximum row to enumerate.</param>
		/// <param name="endColumn">The maximum column to enumerate.</param>
		/// <returns>The constrained enumerator for the cellstore.</returns>
		public ICellStoreEnumerator<T> GetEnumerator(int startRow, int startColumn, int endRow, int endColumn)
		{
			return new ZCellStoreEnumerator(this, startRow, startColumn, endRow, endColumn);
		}
		#endregion

		#region Private Methods
		private bool TryUpdateIndices(ref int row, ref int column)
		{
			if (row < 1 || row > this.MaximumRow || column < 1 || column > this.MaximumColumn)
				return false;
			row--;
			column--;
			return true;
		}

		private bool NextCellBound(ref int row, ref int column, int minRow, int maxRow, int minColumn, int maxColumn)
		{
			// This method finds the next cell in ZCellStore that is within the provided bound indices.
			// The algorithm is as follows [Assume that at all points there is also validation on any found item and row resets are within the bound indices]:
			// 1. Determine if currently iterating row 0 which checks the column metadata
			//		If so: search the column metadata
			// 2. Determine if currently in the middle of a row
			//		If so: finish checking exactly that row and return true if a value is found.
			//				If no value is found, reset to the next row.
			//		If not: go to step 3
			// 3. Determine if the current column to search includes the row metadata
			// 3.1  If so: check that specific cell and return true if a value is found. 
			//					If no value is found, move to sheet content.
			//			If not: go to step 3.2
			// 3.2	Check the worksheet across all columns for the closest cell to the target row
			//
			// There is some room for improvement. It makes the algorithm much messier so I'm saving it for future work.
			// We can theoretically combine steps 2 and 3 so that the search across all columns only happens once. 
			// As it stands, there are columns that will be revisited because Step 2 only checks the specific target row
			// as opposed to just the next cell in that column.

			// tests for this
			if (row < minRow)
				row = minRow;
			if (column < minColumn - 1)
				column = minColumn - 1;

			int columnSearch, rowSearch;
			// STEP 1:: Search column metadata if targeting row 0
			if (row == 0)
			{
				columnSearch = column - 1; // 1-indexing
				if (this.ColumnData.NextItem(ref columnSearch) && columnSearch < maxColumn)
				{
					column = columnSearch + 1;
					return true;
				}
				row++;
				if (row > maxRow)
					return false;
				column = minColumn - 1;
			}
			// STEP 2:: Finish searching current target row
			if (column != -1)
			{
				columnSearch = column - 1;
				while (this.CellData.NextItem(ref columnSearch) && columnSearch < maxColumn)
				{
					var page = this.CellData.GetItem(columnSearch);
					var item = page?.Value.GetItem(row - 1);
					if (item.HasValue)
					{
						column = columnSearch + 1;
						return true;
					}
				}
				row++;
				if (row > maxRow)
					return false;
				column = minColumn - 1;
			}
			// STEP 3:: Search a full row from start
			// STEP 3.1:: Check the row metadata cell if necessary
			int currentRow = int.MaxValue;
			int currentColumn = int.MaxValue;
			if (minColumn == 0)
			{
				rowSearch = row - 1 - 1; // 1-indexing - negative search offset
				if (this.RowData.NextItem(ref rowSearch) && rowSearch < maxRow)
				{
					currentRow = rowSearch + 1;
					currentColumn = 0;
					if (currentRow == row)
					{
						column = currentColumn;
						return true;
					}
				}
			}
			// STEP 3.2:: Check the sheet content for the target row
			columnSearch = column - 1;
			while (this.CellData.NextItem(ref columnSearch) && columnSearch < maxColumn)
			{
				var page = this.CellData.GetItem(columnSearch);
				rowSearch = row - 1 - 1; // 1-indexing - negative search offset
				if (page?.Value.NextItem(ref rowSearch) == true && rowSearch < maxRow)
				{
					if (rowSearch + 1 < currentRow)
					{
						currentRow = rowSearch + 1;
						currentColumn = columnSearch + 1;
						if (currentRow == row)
						{
							column = currentColumn;
							return true;
						}
					}
				}
			}
			row = currentRow;
			column = currentColumn;
			return row != int.MaxValue && column != int.MaxValue;
		}
		#endregion

		#region Nested Classes
		/// <summary>
		///  Understanding the PagedStructure
		///  Author: Zachary Faltersack
		/// 
		///  The PagedStructure is a 0-index collection that stores information across pages.
		///  When the PagedStructure is instantiated, it is provided a bit count that indicates how large it is.
		///  This bit count represents the number of bits used to represent the indices in a single page.
		///  That same number is the number of pages to create in the page list.
		///  Effectively this means data is stored in a square.
		/// 
		///  Example:
		///  Given a PagedStructure with 2 bits, we end up with 2^2 == 4 indices per page, with 4 pages and a total of 16 data points.
		///  Here we see that there are 4 pages, with 4 indices each and the actual index is the value in each slot:
		///  P : Page
		///  I : Index
		/// 
		///  				I0	I1	I2  I3
		///  	{
		///  P0		{  0,  1,  2,  3 },
		///  P1		{  4,  5,  6,  7 },
		///  P2		{  8,  9, 10, 11 },
		///  P3		{ 12, 13, 14, 15 }
		///  	}
		/// 
		///  This allows a rapid lookup by deconstructing an external index using bitshifting and bitmasking 
		///  into page and sub-indices (eg 13 = [P3,I1]).
		///  It also improves memory usage by breaking the structure up into smaller units that can be distributed in the heap
		///  and only instantiates pages that are required to store data.
		/// 
		///  Since the PagedStructure can contain any type of data, the actual values are wrapped in a struct
		///  called ValueHolder. This allows us to defined values as Nullable and improves performance by ensuring null
		///  is written into the ZCellStore (specifically this helps improve the enumeration performance).
		/// 
		/// </summary>
		/// <typeparam name="S">The type this structure is going to contain.</typeparam>
		internal class PagedStructure<S>
		{
			#region Properties
			/// <summary>
			/// Gets the number of bits used to define a <see cref="Page"/>'s indices.
			/// Use to define a bit shift to retrive the primary index (the <see cref="Page"/> index).
			/// </summary>
			public int PageBits { get; } 

			/// <summary>
			/// Gets the maximum index for any single <see cref="Page"/>.
			/// </summary>
			public int PageSize { get; } 

			/// <summary>
			/// Gets the bitmask to use for identifying the secondary index (the index into a <see cref="Page"/>).
			/// </summary>
			public int PageMask { get; }

			/// <summary>
			/// Gets the maximum 0-based index for this <see cref="PagedStructure{S}"/>.
			/// </summary>
			public int MaximumIndex { get; }

			/// <summary>
			/// Gets the minimum index that contains a value in this <see cref="PagedStructure{S}"/>.
			/// </summary>
			public int MinimumUsedIndex { get; private set; }

			/// <summary>
			/// Gets the maximum index that contains a value in this <see cref="PagedStructure{S}"/>.
			/// </summary>
			public int MaximumUsedIndex { get; private set; }

			/// <summary>
			/// Gets a value indicating whether or not this <see cref="PagedStructure{S}"/> contains any data.
			/// </summary>
			public bool IsEmpty
			{
				get { return this.MinimumUsedIndex == this.MaximumIndex + 1 && this.MaximumUsedIndex == -1; }
			}

			private Page[] Pages { get; } 
			#endregion

			#region Constructors
			/// <summary>
			/// Creates an instance of a <see cref="PagedStructure{S}"/>.
			/// </summary>
			/// <param name="pageBits">The number of bits to use for defining the index range for this <see cref="PagedStructure{S}"/>.</param>
			public PagedStructure(int pageBits)
			{
				this.PageBits = pageBits;
				this.PageSize = 1 << pageBits;
				this.PageMask = this.PageSize - 1;
				this.MaximumIndex = (this.PageSize << pageBits) - 1;
				this.Pages = new Page[this.PageSize];
				this.MinimumUsedIndex = this.MaximumIndex + 1;
				this.MaximumUsedIndex = -1;
				// TODO ZPF Future work is to ensure that when a page becomes empty (by removing/clearing/shifting values) we null it out and remove it from out page list.
			}
			#endregion

			#region Public Methods
			/// <summary>
			/// Gets the item at the provided index if it exists.
			/// </summary>
			/// <param name="index">The index of the desired item.</param>
			/// <returns>The value if it exists; otherwise null.</returns>
			public ValueHolder? GetItem(int index)
			{
				if (index < 0 || index > this.MaximumIndex)
					return null;
				this.DeConstructIndex(index, out int page, out int innerIndex);
				return this.Pages[page]?[innerIndex];
			}

			/// <summary>
			/// Sets the provided <paramref name="item"/> and the given index.
			/// </summary>
			/// <param name="index">The index to put the item in.</param>
			/// <param name="item">The item to insert.</param>
			public void SetItem(int index, ValueHolder? item)
			{
				if (index < 0 || index > this.MaximumIndex)
					return;
				this.DeConstructIndex(index, out int page, out int innerIndex);
				var pageArray = this.Pages[page];
				if (null == pageArray)
					this.Pages[page] = pageArray = new Page(this.PageSize);
				pageArray[innerIndex] = item;
				this.UpdateBounds();
			}

			/// <summary>
			/// Shifts items in this collection.
			/// If <paramref name="amount"/> is positive, then all items after <paramref name="index"/> are shifted forward
			/// and null is inserted in the generated space.
			/// If <paramref name="amount"/> is negative, then all items after <paramref name="index"/> are shifted backwards,
			/// overwriting the <paramref name="amount"/> of items immediately after <paramref name="index"/> and null is inserted
			/// at the end.
			/// </summary>
			/// <param name="index">The index after which all items are shifted.</param>
			/// <param name="amount">The amount to shift items forward or backwards.</param>
			public void ShiftItems(int index, int amount)
			{
				if (index < 0 || index > this.MaximumIndex)
					return;
				// Shift forward
				if (amount > 0)
				{
					this.DeConstructIndex(this.MaximumUsedIndex, out int sourcePageIndex, out int sourceInnerIndex);
					this.DeConstructIndex(amount, out int numberOfPagesToShift, out int numberOfInnerIndexToShift);

					for (int i = this.MaximumUsedIndex; i >= Math.Max(index, this.MinimumUsedIndex); --i)
					{
						var targetPageIndex = sourcePageIndex + numberOfPagesToShift;
						var targetPageInnerIndex = sourceInnerIndex + numberOfInnerIndexToShift;
						if (targetPageInnerIndex >= this.PageSize)
						{
							targetPageIndex++;
							targetPageInnerIndex &= this.PageMask;
						}
						if (targetPageIndex >= this.Pages.Length)
							throw new ArgumentOutOfRangeException();
						var targetPage = this.Pages[targetPageIndex] ?? (this.Pages[targetPageIndex] = new Page(this.PageSize));
						var sourcePage = this.Pages[sourcePageIndex];
						if (sourcePage == null)
							targetPage[targetPageInnerIndex] = null;
						else
						{
							targetPage[targetPageInnerIndex] = sourcePage[sourceInnerIndex];
							sourcePage[sourceInnerIndex] = null;
						}
						if (sourceInnerIndex > 0)
							sourceInnerIndex--;
						else
						{
							sourceInnerIndex = this.PageSize - 1;
							sourcePageIndex--;
						}
					}
				}
				// Shift backward
				else if (amount < 0)
				{
					// This represents a delete operation so start at the given index and copy back
					// the desired data.
					amount = -amount;

					this.DeConstructIndex(index, out int targetPageIndex, out int targetInnerIndex);
					this.DeConstructIndex(amount, out int numberOfPagesToShift, out int numberOfInnerIndexToShift);

					for (int i = index; i <= this.MaximumUsedIndex; ++i)
					{
						var sourcePageIndex = targetPageIndex + numberOfPagesToShift;
						var sourcePageInnerIndex = targetInnerIndex + numberOfInnerIndexToShift;
						if (sourcePageInnerIndex >= this.PageSize)
						{
							sourcePageIndex++;
							sourcePageInnerIndex &= this.PageMask;
						}
						var targetPage = this.Pages[targetPageIndex] ?? (this.Pages[targetPageIndex] = new Page(this.PageSize));
						if (sourcePageIndex < this.Pages.Length)
						{
							var sourcePage = this.Pages[sourcePageIndex];
							if (sourcePage == null)
								targetPage[targetInnerIndex] = null;
							else
							{
								targetPage[targetInnerIndex] = sourcePage[sourcePageInnerIndex];
								sourcePage[sourcePageInnerIndex] = null;
							}
						}
						else
							targetPage[targetInnerIndex] = null;
						if (targetInnerIndex < this.PageSize - 1)
							targetInnerIndex++;
						else
						{
							targetInnerIndex = 0;
							targetPageIndex++;
						}
					}
				}
				this.UpdateBounds();
			}

			/// <summary>
			/// Clears <paramref name="amount"/> items after the provided <paramref name="index"/>.
			/// </summary>
			/// <param name="index">The index to start clearing from.</param>
			/// <param name="amount">The number of items to clear.</param>
			public void ClearItems(int index, int amount)
			{
				if (index < 0 || index > this.MaximumIndex)
					return;

				this.DeConstructIndex(index, out int pageIndex, out int pageInnerIndex);

				var firstPage = this.Pages[pageIndex];
				if (firstPage != null)
				{
					for (int i = pageInnerIndex; i < this.PageSize && amount > 0; ++i)
					{
						firstPage[i] = null;
						amount--;
					}
				}
				pageIndex++;
				if (amount > 0 && pageIndex < this.Pages.Length)
				{
					this.DeConstructIndex(amount, out int pagesToClear, out int pagesInnerIndexToClear);
					for (int i = 0; i < pagesToClear; ++i)
					{
						if (pageIndex + i >= this.Pages.Length)
							break;
						this.Pages[pageIndex + i] = null;
					}
					if (pageIndex + pagesToClear < this.Pages.Length)
					{
						var lastPage = this.Pages[pageIndex + pagesToClear];
						if (lastPage != null)
						{
							for (int i = 0; i < pagesInnerIndexToClear; ++i)
							{
								lastPage[i] = null;
							}
						}
					}
				}
				this.UpdateBounds();
			}

			/// <summary>
			/// Updates <paramref name="index"/> to the next index that has a non-null value in the collection
			/// and returns true if an item is found. Otherwise false is returned.
			/// </summary>
			/// <param name="index">The index after which the next item should be found.</param>
			/// <returns>true if another item exists after index; otherwise false.</returns>
			public bool NextItem(ref int index)
			{
				if (this.IsEmpty || index > this.MaximumUsedIndex)
					return false;
				if (index < this.MinimumUsedIndex)
				{
					index = this.MinimumUsedIndex;
					return true;
				}

				this.DeConstructIndex(index, out int minimumPage, out int minimumInnerIndex);
				this.DeConstructIndex(this.MaximumUsedIndex, out int maximumPage, out int maximumInnerIndex);

				int nextIndex = minimumInnerIndex;
				if (this.Pages[minimumPage]?.TryGetNextIndex(minimumInnerIndex, out nextIndex) == true)
				{
					index = this.ReConstructIndex(minimumPage, nextIndex);
					return true;
				}
				else
				{
					for (int page = minimumPage + 1; page <= maximumPage; ++page)
					{
						var currentPage = this.Pages[page];
						if (currentPage?.IsEmpty == false)
						{
							index = this.ReConstructIndex(page, currentPage.MinimumUsedIndex);
							return true;
						}
					}
				}
				return false;
			}

			/// <summary>
			/// Updates <paramref name="index"/> to the previous index that has a non-null value in the collection
			/// and returns true if an item is found. Otherwise false is returned.
			/// </summary>
			/// <param name="index">The index before which the next item should be found.</param>
			/// <returns>true if another item exists before index; otherwise false.</returns>
			public bool PreviousItem(ref int index)
			{
				if (this.IsEmpty || index < this.MinimumUsedIndex)
					return false;
				if (index > this.MaximumUsedIndex)
				{
					index = this.MaximumUsedIndex;
					return true;
				}

				this.DeConstructIndex(this.MinimumUsedIndex, out int minimumPage, out int minimumInnerIndex);
				this.DeConstructIndex(index, out int maximumPage, out int maximumInnerIndex);

				int previousIndex = maximumInnerIndex;
				if (this.Pages[maximumPage]?.TryGetPreviousIndex(maximumInnerIndex, out previousIndex) == true)
				{
					index = this.ReConstructIndex(maximumPage, previousIndex);
					return true;
				}
				else
				{
					for (int page = maximumPage - 1; page >= minimumPage; --page)
					{
						var currentPage = this.Pages[page];
						if (currentPage?.IsEmpty == false)
						{
							index = this.ReConstructIndex(page, currentPage.MaximumUsedIndex);
							return true;
						}
					}
				}
				return false;
			}
			#endregion

			#region Private Methods
			private void UpdateBounds()
			{
				this.MinimumUsedIndex = this.MaximumIndex + 1;
				this.MaximumUsedIndex = -1;
				for (int i = 0; i < this.PageSize; i++)
				{
					var page = this.Pages[i];
					if (page != null)
					{
						if (page.IsEmpty)
							this.Pages[i] = null;
						else
						{
							this.MinimumUsedIndex = this.ReConstructIndex(i, this.Pages[i].MinimumUsedIndex);
							break;
						}
					}
				}
				// Find the max index starting at the end in order to avoid enumerating
				// pages that are likely to have data.
				for (int i = this.PageSize - 1; i >= 0; --i)
				{
					var page = this.Pages[i];
					if (page != null)
					{
						if (page.IsEmpty)
							this.Pages[i] = null;
						else
						{
							this.MaximumUsedIndex = this.ReConstructIndex(i, this.Pages[i].MaximumUsedIndex);
							break;
						}
					}
				}
			}

			private void DeConstructIndex(int index, out int page, out int innerIndex)
			{
				page = index >> this.PageBits;
				innerIndex = index & this.PageMask;
			}

			private int ReConstructIndex(int page, int innerIndex)
			{
				var index = page << this.PageBits;
				index = index | innerIndex;
				return index;
			}
			#endregion

			#region Nested Structs
			/// <summary>
			/// This is a Nullable placeholder for the values to contain within the <see cref="PagedStructure{S}"/>.
			/// Generics can't be guaranteed to be nullable so this is allows nullable values to be used elsewhere.
			/// </summary>
			internal struct ValueHolder
			{
				#region Properties
				/// <summary>
				/// Gets or sets the value for this <see cref="ValueHolder"/>.
				/// </summary>
				public S Value { get; set; }
				#endregion

				#region Implicit Operators
				/// <summary>
				/// Converts the <paramref name="value"/> to a <see cref="ValueHolder"/>.
				/// </summary>
				/// <param name="value">The value to wrap.</param>
				static public implicit operator ValueHolder(S value)
				{
					return new ValueHolder { Value = value };
				}

				/// <summary>
				/// Returns the value that this <see cref="ValueHolder"/> contains.
				/// </summary>
				/// <param name="valueHolder">The <see cref="ValueHolder"/> to extract a value from.</param>
				static public implicit operator S(ValueHolder valueHolder)
				{
					return valueHolder.Value;
				}
				#endregion
			}
			#endregion

			#region Nested Classes
			/// <summary>
			/// This is a helper class that represents a single page within the <see cref="PagedStructure{S}"/> and
			/// contains functionality to abstract basic operations.
			/// </summary>
			public class Page
			{
				#region Properties
				/// <summary>
				/// Gets the minimum used index on this page.
				/// </summary>
				public int MinimumUsedIndex { get; private set; }

				/// <summary>
				/// Gets the maximum used index on this page.
				/// </summary>
				public int MaximumUsedIndex { get; private set; }
				
				/// <summary>
				/// Gets or sets the value at the given <paramref name="index"/>.
				/// </summary>
				/// <param name="index">The index to get a value for.</param>
				/// <returns>The value at that index.</returns>
				public ValueHolder? this[int index]
				{
					get { return this.Values[index]; }
					set
					{
						this.Values[index] = value;
						if (null == value)
							this.UpdateNulledIndex(index);
						else
							this.UpdateIndex(index);
					}
				}

				/// <summary>
				/// Gets a value indicating whether or not this page is empty.
				/// </summary>
				public bool IsEmpty
				{
					get { return this.MinimumUsedIndex == this.Values.Length && this.MaximumUsedIndex == -1; }
				}

				private ValueHolder?[] Values { get; }
				#endregion

				#region Constructors
				/// <summary>
				/// Creates an instance of a <see cref="Page"/>.
				/// </summary>
				/// <param name="size">The number of items on this page.</param>
				public Page(int size)
				{
					this.Values = new ValueHolder?[size];
					this.SetEmptyIndices();
				}
				#endregion

				#region Public Methods
				/// <summary>
				/// Finds the next index on this page with an item in it.
				/// </summary>
				/// <param name="fromIndex">The index to start searching from.</param>
				/// <param name="foundIndex">The index that was found.</param>
				/// <returns>true if an item was found; otherwise false.</returns>
				public bool TryGetNextIndex(int fromIndex, out int foundIndex)
				{
					for (foundIndex = fromIndex + 1; foundIndex < this.Values.Length; ++foundIndex)
					{
						if (this.Values[foundIndex].HasValue)
							return true;
					}
					return false;
				}

				/// <summary>
				/// Finds the previous item on this page with an item in it.
				/// </summary>
				/// <param name="fromIndex">The index to start searching from.</param>
				/// <param name="foundIndex">The index that was found.</param>
				/// <returns>true if an item was found; otherwise false.</returns>
				public bool TryGetPreviousIndex(int fromIndex, out int foundIndex)
				{
					for (foundIndex = fromIndex - 1; foundIndex >= 0; --foundIndex)
					{
						if (this.Values[foundIndex].HasValue)
							return true;
					}
					return false;
				}
				#endregion

				#region Private Methods
				private void UpdateNulledIndex(int index)
				{
					var equalsMinimum = index == this.MinimumUsedIndex;
					var equalsMaximum = index == this.MaximumUsedIndex;
					if (equalsMinimum)
					{
						if (equalsMaximum)
							this.SetEmptyIndices();
						else
						{
							this.TryGetNextIndex(index, out index);
							this.MinimumUsedIndex = index;
						}
					}
					else if (equalsMaximum)
					{
						this.TryGetPreviousIndex(index, out index);
						this.MaximumUsedIndex = index;
					}
				}

				private void UpdateIndex(int index)
				{
					if (index < this.MinimumUsedIndex)
						this.MinimumUsedIndex = index;
					if (index > this.MaximumUsedIndex)
						this.MaximumUsedIndex = index;
				}

				private void SetEmptyIndices()
				{
					this.MinimumUsedIndex = this.Values.Length;
					this.MaximumUsedIndex = -1;
				}
				#endregion
			}
			#endregion

			#region Test Helpers
			/// <summary>
			/// Loads data into the <see cref="PagedStructure{S}"/>.
			/// This is used for unit tests.
			/// </summary>
			/// <param name="pageData">The data to load into the cell store.</param>
			internal void LoadPages(ValueHolder?[,] pageData)
			{
				for (int row = 0; row <= pageData.GetUpperBound(0); ++row)
				{
					for (int column = 0; column <= pageData.GetUpperBound(1); ++column)
					{
						var pageArray = this.Pages[row];
						if (null == pageArray)
							this.Pages[row] = pageArray = new Page(this.PageSize);
						pageArray[column] = pageData[row, column];
					}
				}
				this.UpdateBounds();
			}

			/// <summary>
			/// Validates the content in the <see cref="PagedStructure{S}"/>.
			/// This is used for unit tests.
			/// </summary>
			/// <param name="pageData">The data against which to validate.</param>
			/// <param name="invalidIndex">An action to invoke with a message and the coordinates for invalid data.</param>
			internal void ValidatePages(ValueHolder?[,] pageData, Action<int, int, string> invalidIndex)
			{
				for (int row = 0; row <= pageData.GetUpperBound(0); ++row)
				{
					for (int column = 0; column <= pageData.GetUpperBound(1); ++column)
					{
						var item = this.Pages[row]?[column];
						var data = pageData[row, column];
						if (data.HasValue != item.HasValue || (data.HasValue && !data.Value.Value.Equals(item.Value.Value)))
							invalidIndex(row, column, $"Expected: {(data?.Value.ToString() ?? "null")}, Actual: {(item?.Value.ToString() ?? "null")}");
					}
				}
			}
			#endregion
		}

		/// <summary>
		/// Enumerates a cell store within a bounded set of coordinates.
		/// </summary>
		private class ZCellStoreEnumerator : ICellStoreEnumerator<T>
		{
			#region Class Variables
			private int myColumn;
			private int myRow;
			#endregion
			
			#region Properties
			private ZCellStore<T> CellStore { get; }
			private int StartRow { get; }
			private int StartColumn { get; }
			private int EndRow { get; }
			private int Endcolumn { get; }
			#endregion

			#region Constructors
			/// <summary>
			/// Creates a default instance of a <see cref="ZCellStoreEnumerator"/> that enumerates the entire cell store.
			/// </summary>
			/// <param name="zCellStore">The cell store to enumerate.</param>
			public ZCellStoreEnumerator(ZCellStore<T> zCellStore) :
				this(zCellStore, 0, 0, zCellStore.MaximumRow, zCellStore.MaximumColumn) { }

			/// <summary>
			/// Creates an instance of a <see cref="ZCellStoreEnumerator"/> that enumerates the cell store in a bounded context.
			/// </summary>
			/// <param name="zCellStore">The cell store to enumerate.</param>
			/// <param name="startRow">The starting row to enumerate.</param>
			/// <param name="startColumn">The starting column to enumerate.</param>
			/// <param name="endRow">The ending row to enumerate.</param>
			/// <param name="endColumn">The ending column to enumerate.</param>
			public ZCellStoreEnumerator(ZCellStore<T> zCellStore, int startRow, int startColumn, int endRow, int endColumn)
			{
				this.CellStore = zCellStore;
				
				this.StartRow = startRow;
				this.StartColumn = startColumn;
				this.EndRow = endRow;
				this.Endcolumn = endColumn;
				this.Reset();
			}
			#endregion

			#region ICellStorEnumerator Members
			/// <summary>
			/// Gets the current cell address in string format.
			/// </summary>
			public string CellAddress => ExcelAddress.GetAddress(this.Row, this.Column);

			/// <summary>
			/// Gets the current column.
			/// </summary>
			public int Column => myColumn;

			/// <summary>
			/// Gets the current row.
			/// </summary>
			public int Row => myRow;

			/// <summary>
			/// Gets or sets the value at the current row and column.
			/// </summary>
			public T Value
			{
				get { return this.CellStore.GetValue(this.Row, this.Column); }
				set { this.CellStore.SetValue(this.Row, this.Column, value); }
			}

			/// <summary>
			/// Gets the current value.
			/// </summary>
			public T Current => this.Value;

			/// <summary>
			/// Gets the current value.
			/// </summary>
			object IEnumerator.Current
			{
				get
				{
					this.Reset();
					return this;
				}
			}

			/// <summary>
			/// Disposes the <see cref="ZCellStoreEnumerator"/>.
			/// </summary>
			public void Dispose()
			{
				// Nothing to dispose since we just index into a cellstore.
			}

			/// <summary>
			/// Resets this <see cref="ZCellStoreEnumerator"/> and returns itself.
			/// </summary>
			/// <returns>This instance.</returns>
			public IEnumerator<T> GetEnumerator()
			{
				this.Reset();
				return this;
			}

			/// <summary>
			/// Moves to the next item in the cell store.
			/// </summary>
			/// <returns>true if another item exists; otherwise false.</returns>
			public bool MoveNext()
			{
				return this.CellStore.NextCellBound(ref myRow, ref myColumn, this.StartRow, this.EndRow, this.StartColumn, this.Endcolumn);
			}

			/// <summary>
			/// Resets this instance to the start.
			/// </summary>
			public void Reset()
			{
				myRow = this.StartRow;
				myColumn = this.StartColumn - 1;
			}

			/// <summary>
			/// Resets this <see cref="ZCellStoreEnumerator"/> and returns itself.
			/// </summary>
			/// <returns>This instance.</returns>
			IEnumerator IEnumerable.GetEnumerator()
			{
				this.Reset();
				return this;
			}
			#endregion
		}
		#endregion
	}

	/// <summary>
	/// At some point this data structure can go away. It's left for legacy purposes as future cleanup.
	/// For the most part, it looks like the only CellFlags that is using it is RichText. Parsing XML
	/// is likely also setting values for the other CellFlags but they aren't ever modified after that
	/// based on a cursory search.
	/// </summary>
	internal class ZFlagStore : ZCellStore<byte>, IFlagStore
	{
		#region IFlagStore Members
		/// <summary>
		/// Adds or removes the given <paramref name="cellFlags"/> value based on <paramref name="value"/>.
		/// </summary>
		/// <param name="Row">The cell row to set a flag for.</param>
		/// <param name="Col">The cell column to set a flag for.</param>
		/// <param name="value">A boolean value indicating whether to add or remove the flag value.</param>
		/// <param name="cellFlags">The flags to set for the cell.</param>
		public void SetFlagValue(int Row, int Col, bool value, CellFlags cellFlags)
		{
			CellFlags currentValue = (CellFlags)base.GetValue(Row, Col);
			if (value)
				base.SetValue(Row, Col, (byte)(currentValue | cellFlags)); // add the CellFlag bit
			else
				base.SetValue(Row, Col, (byte)(currentValue & ~cellFlags)); // remove the CellFlag bit
		}

		/// <summary>
		/// Gets the flag values from a given cell.
		/// </summary>
		/// <param name="Row">The cell row to get a flag for.</param>
		/// <param name="Col">The cell column to get a flag for.</param>
		/// <param name="cellFlags">The flags to query for.</param>
		/// <returns>True if the flags are set; otherwise false.</returns>
		public bool GetFlagValue(int Row, int Col, CellFlags cellFlags)
		{
			return !(((byte)cellFlags & base.GetValue(Row, Col)) == 0);
		}
		#endregion
	}
}
