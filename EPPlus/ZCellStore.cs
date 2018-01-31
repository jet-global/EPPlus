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
 * Author							   Change									Date
 * ******************************************************************************
 * Zachary Faltersack    Added       		        2017-12-22
 *******************************************************************************/

using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

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
	/// </summary>
	/// <typeparam name="S">The type this ZCellStore is going to contain.</typeparam>
	internal class ZCellStore<T> : ICellStore<T>
	{
		#region Constants
		/*
		 * Default Row Page Bits results in an index range of: 1..1048576
		 * Default Column Page Bits results in an index range of: 1..16384
		 * 
		 * */
		private const int DefaultRowPageBits = 10; // 2 ^ 10 = 1024 so right-shifting by 10 bits divides by 1024
		private const int DefaultColumnPageBits = 7; // 2 ^ 7 = 128 so right-shifting by 7 bits divides by 128
		#endregion

		#region Properties
		private PagedStructure<PagedStructure<T>> CellData { get; } 
		private PagedStructure<T> ColumnData { get; }
		private PagedStructure<T> RowData { get; }
		private int RowPageBits { get; }
		private int ColumnPageBits { get; }
		#endregion

		#region Constructors
		public ZCellStore() : this(ZCellStore<T>.DefaultRowPageBits, ZCellStore<T>.DefaultColumnPageBits) { }

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

		public bool Exists(int row, int column)
		{
			return this.Exists(row, column, out _);
		}

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

		public bool NextCell(ref int row, ref int column)
		{
			return this.NextCellBound(ref row, ref column, 0, this.MaximumRow, 0, this.MaximumColumn);
		}

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

		public void Delete(int fromRow, int fromCol, int rows, int columns)
		{
			if (fromRow != 0 && fromCol != 0)
				throw new InvalidOperationException("Only delete rows or columns in a single operation.");

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

		public void Insert(int fromRow, int fromCol, int rows, int columns)
		{
			if (fromRow != 0 && fromCol != 0)
				throw new InvalidOperationException("Only insert rows or columns in a single operation.");
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
			// TODO ZPF do we really need this?...
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
			// This method finds the next cell in ZCellStore that is within the provided bound indices
			// There is some performance tuning that can be achieved here, particularly in STEP 3.

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
			int currentRow = int.MaxValue;
			int currentColumn = int.MaxValue;
			if (minColumn == 0)
			{
				rowSearch = row - 1 - 1; // 1-indexing - negative search offset
				if (this.RowData.NextItem(ref rowSearch) && rowSearch < maxRow)
				{
					currentRow = rowSearch + 1;
					currentColumn = 0;
				}
			}
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
			public ValueHolder? GetItem(int index)
			{
				if (index < 0 || index > this.MaximumIndex)
					return null;
				this.DeConstructIndex(index, out int page, out int innerIndex);
				return this.Pages[page]?[innerIndex];
			}

			public void SetItem(int index, ValueHolder? item)
			{
				this.SetItem(index, item, true);
			}

			public void ShiftItems(int index, int amount)
			{
				// TODO ZPF This can be highly optimized.
				// What we want to do is breakdown the index and directly work with the pages
				// Even though deconstructing the indices is relatively quick, we don't need to do it 
				// for every item that moves.
				if (index < 0 || index > this.MaximumIndex)
					return;
				// Shift forward
				if (amount > 0)
				{
					// Start at the end and shift back from there so as not to overwrite data.
					for (int i = this.MaximumUsedIndex; i >= index; --i)
					{
						var target = i + amount;
						if (target <= this.MaximumIndex)
							this.SetItem(target, this.GetItem(i), false);
						else
							throw new ArgumentOutOfRangeException();
						this.SetItem(i, null, false);
					}
				}
				// Shift backward
				else if (amount < 0)
				{
					// This represents a delete operation so start at the given index and copy back
					// the desired data.
					amount = -amount;
					for (int i = index; i <= this.MaximumUsedIndex; ++i)
					{
						var source = i + amount;
						if (source <= this.MaximumIndex)
							this.SetItem(i, this.GetItem(source), false);
						else
							this.SetItem(i, null, false);
					}
				}
				this.UpdateBounds();
			}

			public void ClearItems(int index, int amount)
			{
				// TODO ZPF This can be optimized by working directly with the pages instead of deconstructing 
				// each index for every call.
				if (index < 0 || index > this.MaximumIndex)
					return;
				var target = Math.Min(index + amount - 1, this.MaximumUsedIndex);
				for (; index <= target; ++index)
				{
					this.SetItem(index, null, false);
				}
				this.UpdateBounds();
			}

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
					if (this.Pages[i]?.IsEmpty == false)
					{
						this.MinimumUsedIndex = this.ReConstructIndex(i, this.Pages[i].MinimumUsedIndex);
						break;
					}
				}
				for (int i = this.PageSize - 1; i >= 0; i--)
				{
					if (this.Pages[i]?.IsEmpty == false)
					{
						this.MaximumUsedIndex = this.ReConstructIndex(i, this.Pages[i].MaximumUsedIndex);
						break;
					}
				}
			}

			private void SetItem(int index, ValueHolder? item, bool doBoundsUpdate)
			{
				if (index < 0 || index > this.MaximumIndex)
					return;
				this.DeConstructIndex(index, out int page, out int innerIndex);
				var pageArray = this.Pages[page];
				if (null == pageArray)
					this.Pages[page] = pageArray = new Page(this.PageSize);
				pageArray[innerIndex] = item;
				if (doBoundsUpdate)
					this.UpdateBounds();
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
			internal struct ValueHolder
			{
				public S Value { get; set; }
				static public implicit operator ValueHolder(S value)
				{
					return new ValueHolder { Value = value };
				}
				static public implicit operator S(ValueHolder valueHolder)
				{
					return valueHolder.Value;
				}
			}
			#endregion

			#region Nested Classes
			public class Page
			{
				#region Properties
				public int MinimumUsedIndex { get; set; }

				public int MaximumUsedIndex { get; set; }
				
				public ValueHolder? this[int index]
				{
					get { return this.Values[index]; }
					set
					{
						this.Values[index] = value;
						if (null == value)
							this.UpdatedNulledIndex(index);
						else
							this.UpdateIndex(index);
					}
				}

				public bool IsEmpty
				{
					get { return this.MinimumUsedIndex == this.Values.Length && this.MaximumUsedIndex == -1; }
				}

				private ValueHolder?[] Values { get; }
				#endregion

				#region Constructors
				public Page(int size)
				{
					this.Values = new ValueHolder?[size];
					this.SetEmptyIndices();
				}
				#endregion

				#region Public Methods
				// TODO ZPF cleanup? Only used for test setup/validation.
				public ValueHolder?[] GetValues()
				{
					return this.Values;
				}

				public bool TryGetNextIndex(int fromIndex, out int foundIndex)
				{
					for (foundIndex = fromIndex + 1; foundIndex < this.Values.Length; ++foundIndex)
					{
						if (this.Values[foundIndex].HasValue)
							return true;
					}
					return false;
				}

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
				private void UpdatedNulledIndex(int index)
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
			/// This is included for testing purposes: DO NOT USE
			/// </summary>
			/// <returns></returns>
			public ValueHolder?[][] GetPages()
			{
				return this.Pages.Select(p => p.GetValues()).ToArray();
			}

			public void LoadPages(ValueHolder?[,] pageData)
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

			public void ValidatePages(ValueHolder?[,] pageData, Action<int, int, string> invalidIndex)
			{
				for (int row = 0; row <= pageData.GetUpperBound(0); ++row)
				{
					for (int column = 0; column <= pageData.GetUpperBound(1); ++column)
					{
						var item = this.Pages[row]?[column];
						var data = pageData[row, column];
						if (data.HasValue != item.HasValue || (data.HasValue && !data.Value.Value.Equals(item.Value.Value)))
							invalidIndex(row, column, $"Expected: {(data?.ToString() ?? "null")}");
					}
				}
			}
			#endregion
		}

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
			public ZCellStoreEnumerator(ZCellStore<T> zCellStore) :
				this(zCellStore, 0, 0, ExcelPackage.MaxRows, ExcelPackage.MaxColumns)
			{
			}

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
			public string CellAddress => ExcelAddressBase.GetAddress(this.Row, this.Column);

			public int Column => myColumn;

			public int Row => myRow;

			public T Value
			{
				get { return this.CellStore.GetValue(this.Row, this.Column); }
				set { this.CellStore.SetValue(this.Row, this.Column, value); }
			}

			public T Current => this.Value;

			object IEnumerator.Current
			{
				get
				{
					this.Reset();
					return this;
				}
			}

			public void Dispose()
			{
				// TODO ZPF Can we just take this off of the interface?...
			}

			public IEnumerator<T> GetEnumerator()
			{
				this.Reset();
				return this;
			}

			public bool MoveNext()
			{
				return this.CellStore.NextCellBound(ref myRow, ref myColumn, this.StartRow, this.EndRow, this.StartColumn, this.Endcolumn);
			}

			public void Reset()
			{
				myRow = this.StartRow;
				myColumn = this.StartColumn - 1;
			}

			IEnumerator IEnumerable.GetEnumerator()
			{
				this.Reset();
				return this;
			}
			#endregion
		}
		#endregion
	}
}
