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
 * Jan Källman		    Added       		        2012-11-25
 *******************************************************************************/
// This preprocessor constant turns on invariant validation for all modification methods on the CellStore (useful for tracking down weird CellStore bugs). 
// #define DEBUGGING

using System;
using System.Collections;
using System.Collections.Generic;
using OfficeOpenXml;

/// <summary>
/// This is the store for all Rows, Columns and Cells.
/// It is a Dictionary implementation that allows you to change the Key (the RowID, ColumnID or CellID )
/// </summary>
internal class CellStore<T> : ICellStore<T>, IDisposable// : IEnumerable<ulong>, IEnumerator<ulong>
{
	#region Nested classes
	private class IndexBase : IComparable<IndexBase>
	{
		internal short Index;
		public int CompareTo(IndexBase other)
		{
			//return Index < other.Index ? -1 : Index > other.Index ? 1 : 0;
			return Index - other.Index;
		}
	}
	// to compress memory size, use struct
	private struct IndexItem : IComparable<IndexItem>
	{
		internal int IndexPointer
		{
			get;
			set;
		}
		internal short Index;
		public int CompareTo(IndexItem other)
		{
			return Index - other.Index;
		}
	}
	private class ColumnIndex : IndexBase, IDisposable
	{
		internal IndexBase _searchIx = new IndexBase();
		public ColumnIndex()
		{
			_pages = new PageIndex[CellStore<int>.PagesPerColumnMin];
			PageCount = 0;
		}
		~ColumnIndex()
		{
			_pages = null;
		}
		internal int GetPosition(int Row)
		{
			var page = (short)(Row >> CellStore<int>.pageBits);
			_searchIx.Index = page;
			var res = Array.BinarySearch(_pages, 0, PageCount, _searchIx);
			if (res >= 0)
			{
				GetPage(Row, ref res);
				return res;
			}
			else
			{
				var p = ~res;

				if (GetPage(Row, ref p))
				{
					return p;
				}
				else
				{
					return res;
				}
			}
		}

		private bool GetPage(int Row, ref int res)
		{
			if (res < PageCount && _pages[res].MinIndex <= Row && _pages[res].MaxIndex >= Row)
			{
				return true;
			}
			else
			{
				if (res + 1 < PageCount && (_pages[res + 1].MinIndex <= Row))
				{
					do
					{
						res++;
					}
					while (res + 1 < PageCount && _pages[res + 1].MinIndex <= Row);
					//if (res + 1 < PageCount && _pages[res + 1].MaxIndex >= Row)
					//{
					return true;
					//}
					//else
					//{
					//    return false;
					//}
				}
				else if (res - 1 >= 0 && _pages[res - 1].MaxIndex >= Row)
				{
					do
					{
						res--;
					}
					while (res - 1 > 0 && _pages[res - 1].MaxIndex >= Row);
					//if (res > 0)
					//{
					return true;
					//}
					//else
					//{
					//    return false;
					//}
				}
				return false;
			}
		}
		internal int GetNextRow(int row)
		{
			var p = GetPosition(row);
			if (p < 0)
			{
				p = ~p;
				if (p >= PageCount)
				{
					return -1;
				}
				else
				{

					if (_pages[p].IndexOffset + _pages[p].Rows[0].Index < row)
					{
						if (p + 1 >= PageCount)
						{
							return -1;
						}
						else
						{
							return _pages[p + 1].IndexOffset + _pages[p].Rows[0].Index;
						}
					}
					else
					{
						return _pages[p].IndexOffset + _pages[p].Rows[0].Index;
					}
				}
			}
			else
			{
				if (p < PageCount)
				{
					var r = _pages[p].GetNextRow(row);
					if (r >= 0)
					{
						return _pages[p].IndexOffset + _pages[p].Rows[r].Index;
					}
					else
					{
						if (++p < PageCount)
						{
							return _pages[p].IndexOffset + _pages[p].Rows[0].Index;
						}
						else
						{
							return -1;
						}
					}
				}
				else
				{
					return -1;
				}
			}
		}
		internal int FindNext(int Page)
		{
			var p = GetPosition(Page);
			if (p < 0)
			{
				return ~p;
			}
			return p;
		}
		internal PageIndex[] _pages;
		internal int PageCount;

		public void Dispose()
		{
			for (int p = 0; p < PageCount; p++)
			{
				((IDisposable)_pages[p]).Dispose();
			}
			_pages = null;
		}

	}
	private class PageIndex : IndexBase, IDisposable
	{
		internal IndexItem _searchIx = new IndexItem();
		public PageIndex()
		{
			Rows = new IndexItem[CellStore<int>.PageSizeMin];
			RowCount = 0;
		}
		public PageIndex(IndexItem[] rows, int count)
		{
			Rows = rows;
			RowCount = count;
		}
		public PageIndex(PageIndex pageItem, int start, int size)
			: this(pageItem, start, size, pageItem.Index, pageItem.Offset)
		{

		}
		public PageIndex(PageIndex pageItem, int start, int size, short index, int offset)
		{
			Rows = new IndexItem[CellStore<int>.GetSize(size)];
			Array.Copy(pageItem.Rows, start, Rows, 0, size);
			RowCount = size;
			Index = index;
			Offset = offset;
		}
		~PageIndex()
		{
			Rows = null;
		}
		internal int Offset = 0;
		internal int IndexOffset
		{
			get
			{
				return IndexExpanded + (int)Offset;
			}
		}
		internal int IndexExpanded
		{
			get
			{
				return (Index << CellStore<int>.pageBits);
			}
		}
		internal IndexItem[] Rows { get; set; }
		internal int RowCount;

		internal int GetPosition(int offset)
		{
			_searchIx.Index = (short)offset;
			return Array.BinarySearch(Rows, 0, RowCount, _searchIx);
		}
		internal int GetNextRow(int row)
		{
			int offset = row - IndexOffset;
			var o = GetPosition(offset);
			if (o < 0)
			{
				o = ~o;
				if (o < RowCount)
				{
					return o;
				}
				else
				{
					return -1;
				}
			}
			return o;
		}

		public int MinIndex
		{
			get
			{
				if (Rows.Length > 0)
				{
					return IndexOffset + Rows[0].Index;
				}
				else
				{
					return -1;
				}
			}
		}
		public int MaxIndex
		{
			get
			{
				if (RowCount > 0)
				{
					return IndexOffset + Rows[RowCount - 1].Index;
				}
				else
				{
					return -1;
				}
			}
		}
		public int GetIndex(int pos)
		{
			return IndexOffset + Rows[pos].Index;
		}
		public void Dispose()
		{
			Rows = null;
		}
	}

	/// <summary>
	/// This enumerator allows for partial enumeration of an <see cref="ICellStore{T}"/>.
	/// </summary>
	/// <typeparam name="S">The type of the <see cref="ICellStore{T}"/> being enumerated.</typeparam>
	private class CellsStoreEnumerator<S> : ICellStoreEnumerator<S>
	{
		#region Class Variables
		private CellStore<S> _cellStore;
		private int row, colPos;
		private int[] pagePos, cellPos;
		private int _startRow, _startCol, _endRow, _endCol;
		private int minRow, minColPos, maxRow, maxColPos;
		#endregion

		#region Properties
		/// <summary>
		/// Gets the current cell address as a string.
		/// </summary>
		public string CellAddress
		{
			get
			{
				return ExcelAddress.GetAddress(this.Row, this.Column);
			}
		}

		/// <summary>
		///  Gets the current value.
		/// </summary>
		public S Current
		{
			get
			{
				return this.Value;
			}
		}

		/// <summary>
		/// Get the current Enumerator objct.
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
		/// Get the current row.
		/// </summary>
		public int Row
		{
			get
			{
				return this.row;
			}
		}

		/// <summary>
		/// Get the current column.
		/// </summary>
		public int Column
		{
			get
			{
				if (this.colPos == -1)
					this.MoveNext();
				if (this.colPos == -1)
					return 0;
				return this._cellStore.GetColumnPosIndex(this.colPos);
			}
		}

		/// <summary>
		/// Gets or sets the current value.
		/// </summary>
		public S Value
		{
			get
			{
				lock (this._cellStore)
				{
					return this._cellStore.GetValue(this.row, this.Column);
				}
			}
			set
			{
				lock (_cellStore)
				{
					this._cellStore.SetValue(this.row, this.Column, value);
				}
			}
		}
		#endregion

		#region Constructors
		/// <summary>
		/// Initialize a new <see cref="CellsStoreEnumerator{T}"/> for all the cells in the <see cref="ICellStore{T}"/>.
		/// </summary>
		/// <param name="cellStore">The <see cref="ICellStore{T}"/> to fully enumerate.</param>
		public CellsStoreEnumerator(CellStore<S> cellStore) :
			this(cellStore, 0, 0, ExcelPackage.MaxRows, ExcelPackage.MaxColumns)
		{
		}

		/// <summary>
		/// Initialize a new <see cref="CellsStoreEnumerator{T}"/> for the specified subset of cells in the <see cref="ICellStore{T}"/>.
		/// </summary>
		/// <param name="cellStore">The CellStore to partially enumerate.</param>
		/// <param name="StartRow">The first row to include cells from.</param>
		/// <param name="StartCol">The first column to include cells from.</param>
		/// <param name="EndRow">The last row to include cells from.</param>
		/// <param name="EndCol">The last column to include cells from.</param>
		public CellsStoreEnumerator(CellStore<S> cellStore, int StartRow, int StartCol, int EndRow, int EndCol)
		{
			var specificCellStore = cellStore as CellStore<S>;
			if (specificCellStore == null)
				throw new ArgumentException("Unexpected Cell Store Type for the CellsStoreEnumerator.");
			this._cellStore = specificCellStore;

			this._startRow = StartRow;
			this._startCol = StartCol;
			this._endRow = EndRow;
			this._endCol = EndCol;

			this.Init();

		}
		#endregion

		#region Public Methods
		public IEnumerator<S> GetEnumerator()
		{
			this.Reset();
			return this;
		}

		IEnumerator IEnumerable.GetEnumerator()
		{
			this.Reset();
			return this;
		}

		public void Dispose()
		{
			//_cellStore=null;
		}

		public bool MoveNext()
		{
			return this._cellStore.GetNextCell(ref row, ref colPos, minColPos, maxRow, maxColPos);
		}

		public void Reset()
		{
			this.Init();
		}
		#endregion

		#region Private Methods
		private void Init()
		{
			this.minRow = this._startRow;
			this.maxRow = this._endRow;

			this.minColPos = this._cellStore.GetPosition(this._startCol);
			if (minColPos < 0)
				this.minColPos = ~this.minColPos;
			this.maxColPos = this._cellStore.GetPosition(this._endCol);
			if (this.maxColPos < 0)
				this.maxColPos = ~this.maxColPos - 1;
			this.row = this.minRow;
			this.colPos = this.minColPos - 1;

			var cols = this.maxColPos - this.minColPos + 1;
			this.pagePos = new int[cols];
			this.cellPos = new int[cols];
			for (int i = 0; i < cols; i++)
			{
				this.pagePos[i] = -1;
				this.cellPos[i] = -1;
			}
		}
		#endregion
	}
	#endregion

	#region Constants
	/**** Size constants ****/
	internal const int pageBits = 10;   //13bits=8192  Note: Maximum is 13 bits since short is used (PageMax=16K)
	internal const int PageSize = 1 << pageBits;
	internal const int PageSizeMin = 1 << 10;
	internal const int PageSizeMax = PageSize << 1; //Double page size
	internal const int ColSizeMin = 32;
	internal const int PagesPerColumnMin = 32;
	#endregion

	#region Class Variables
	private List<T> _values = new List<T>();
	private ColumnIndex[] _columnIndex;
	private IndexBase _searchIx = new IndexBase();
	private IndexItem _searchItem = new IndexItem();
	private int ColumnCount;
	#endregion

	#region Constructors
	/// <summary>
	/// Initialize a new CellStore.
	/// </summary>
	public CellStore()
	{
		_columnIndex = new ColumnIndex[ColSizeMin];
	}
	#endregion

	#region Destructors
	~CellStore()
	{
		if (_values != null)
		{
			_values.Clear();
			_values = null;
		}
		_columnIndex = null;
	}
	#endregion

	#region Properties
	/// <summary>
	/// Gets the maximum row of the cellstore.
	/// </summary>
	public int MaximumRow => ExcelPackage.MaxRows;

	/// <summary>
	/// Gets the maximum column of the cellstore.
	/// </summary>
	public int MaximumColumn => ExcelPackage.MaxColumns;
	#endregion

	#region Public Methods
	/// <summary>
	/// Dispose of the objects allocated by this CellStore.
	/// </summary>
	public void Dispose()
	{
		if (this._values != null)
			this._values.Clear();
		for (var c = 0; c < this.ColumnCount; c++)
		{
			if (this._columnIndex[c] != null)
			{
				((IDisposable)this._columnIndex[c]).Dispose();
			}
		}
		this._values = null;
		this._columnIndex = null;
	}

	/// <summary>
	/// Set the specified <paramref name="value"/> at the specified <paramref name="row"/> and <paramref name="column"/>.
	/// </summary>
	/// <param name="row">The row to insert at.</param>
	/// <param name="column">The column to insert at.</param>
	/// <param name="value">The value to be inserted.</param>
	public void SetValue(int row, int column, T value)
	{
		lock (this._columnIndex)
		{
			var col = Array.BinarySearch(this._columnIndex, 0, this.ColumnCount, new IndexBase() { Index = (short)(column) });
			var page = (short)(row >> pageBits);
			if (col >= 0)
			{
				var pos = this._columnIndex[col].GetPosition(row);

				if (pos < 0)
				{
					pos = ~pos;
					if (pos - 1 < 0 || this._columnIndex[col]._pages[pos - 1].IndexOffset + CellStore<T>.PageSize - 1 < row)
					{
						this.AddPage(this._columnIndex[col], pos, page);
					}
					else
					{
						pos--;
					}
				}
				if (pos >= this._columnIndex[col].PageCount)
				{
					this.AddPage(this._columnIndex[col], pos, page);
				}
				var pageItem = this._columnIndex[col]._pages[pos];
				if (Math.Min(pageItem.IndexOffset, pageItem.MinIndex) > row)
				{
					pos--;
					page--;
					if (pos < 0)
					{
						throw (new Exception("Unexpected error when setting value"));
					}
					pageItem = this._columnIndex[col]._pages[pos];
				}

				short ix = (short)(row - ((pageItem.Index << CellStore<T>.pageBits) + pageItem.Offset));

				this._searchItem.Index = ix;
				var cellPos = Array.BinarySearch(pageItem.Rows, 0, pageItem.RowCount, this._searchItem);
				if (cellPos < 0)
				{
					cellPos = ~cellPos;
					this.AddCell(this._columnIndex[col], pos, cellPos, ix, value);
				}
				else
				{
					this._values[pageItem.Rows[cellPos].IndexPointer] = value;
				}
			}
			else //Column does not exist
			{
				col = ~col;
				this.AddColumn(col, column);
				this.AddPage(this._columnIndex[col], 0, page);
				short ix = (short)(row - (page << CellStore<T>.pageBits));
				this.AddCell(this._columnIndex[col], 0, 0, ix, value);
			}
		}
#if DEBUGGING
		this.AssertInvariants();
#endif
	}

	/// <summary>
	/// Get the value for a specified cell location.
	/// </summary>
	/// <param name="Row">The row to read data from.</param>
	/// <param name="Column">The column to read data from.</param>
	/// <returns>The value stored at the specified row and column.</returns>
	public T GetValue(int Row, int Column)
	{
		int i = this.GetPointer(Row, Column);
		if (i >= 0)
		{
			return this._values[i];
		}
		else
		{
			return default(T);
		}
	}

	/// <summary>
	/// Gets the location of the first cell before the specified position.
	/// </summary>
	/// <param name="row">The row to start at / the new row, if a cell is found.</param>
	/// <param name="col">The column to start at / the new column, if a cell is found.</param>
	/// <returns>True if a previous cell exists and the values have been updated; false otherwise.</returns>
	public bool PrevCell(ref int row, ref int col)
	{
		return this.PrevCell(ref row, ref col, 0, 0, ExcelPackage.MaxRows, ExcelPackage.MaxColumns);
	}

	/// <summary>
	/// Determine if a value exists at the specified <paramref name="row"/> and <paramref name="column"/>.
	/// </summary>
	/// <param name="row">The row to check for a value.</param>
	/// <param name="column">The column to check for a value.</param>
	/// <returns></returns>
	public bool Exists(int row, int column)
	{
		return this.GetPointer(row, column) >= 0;
	}

	/// <summary>
	/// Determine if a value exists at the specified <paramref name="row"/> and <paramref name="column"/>, and return it if available.
	/// </summary>
	/// <param name="row">The row to check for a value.</param>
	/// <param name="column">The column to check for a value.</param>
	/// <param name="value">The resulting value, if it exists.</param>
	/// <returns>True if a value exists at the specified location; false otherwise.</returns>
	public bool Exists(int row, int column, out T value)
	{
		var p = this.GetPointer(row, column);
		if (p >= 0)
		{
			value = this._values[p];
			return true;
		}
		else
		{
			value = default(T);
			return false;
		}
	}

	/// <summary>
	/// Get the next cell after the specified position.
	/// </summary>
	/// <param name="row">The row to start looking at.</param>
	/// <param name="col">The column to start looking at.</param>
	/// <returns>True if a next cell has been found / the row and column have been updated to reflect the new position.</returns>
	public bool NextCell(ref int row, ref int col)
	{

		return this.NextCell(ref row, ref col, 0, 0, ExcelPackage.MaxRows, ExcelPackage.MaxColumns);
	}

	/// <summary>
	/// Delete the values between the specified row and columns, shifting cells up or left as necessary. 
	/// </summary>
	/// <param name="fromRow">The row to start deleting from.</param>
	/// <param name="fromCol">The column to start deleting at.</param>
	/// <param name="rows">The number of rows to be deleted.</param>
	/// <param name="columns">The number of columns to be deleted.</param>
	public void Delete(int fromRow, int fromCol, int rows, int columns)
	{
		this.Delete(fromRow, fromCol, rows, columns, true);
	}

	/// <summary>
	/// Clear the values in a specific range.
	/// </summary>
	/// <param name="fromRow">The first row to clear from.</param>
	/// <param name="fromCol">The first column to clear from.</param>
	/// <param name="rows">The number of rows to clear.</param>
	/// <param name="columns">The number of columns to clear.</param>
	public void Clear(int fromRow, int fromCol, int rows, int columns)
	{
		this.Delete(fromRow, fromCol, rows, columns, false);
	}

	/// <summary>
	/// Get the dimension (range of used cells) that this CellStore contains. 
	/// </summary>
	/// <param name="fromRow">The top bound of the CellStore's represented area.</param>
	/// <param name="fromCol">The left bound of the CellStore's represented area.</param>
	/// <param name="toRow">The bottom bound of the CellStore's represented area.</param>
	/// <param name="toCol">The right bound of the CellStore's represented area.</param>
	/// <returns>True if the CellStore contains any cells and thus has a dimension; false otherwise.</returns>
	public bool GetDimension(out int fromRow, out int fromCol, out int toRow, out int toCol)
	{
		if (this.ColumnCount == 0)
		{
			fromRow = fromCol = toRow = toCol = 0;
			return false;
		}
		else
		{
			fromCol = this._columnIndex[0].Index;
			var fromIndex = 0;
			if (fromCol <= 0 && this.ColumnCount > 1)
			{
				fromCol = this._columnIndex[1].Index;
				fromIndex = 1;
			}
			else if (this.ColumnCount == 1 && fromCol <= 0)
			{
				fromRow = fromCol = toRow = toCol = 0;
				return false;
			}
			var col = this.ColumnCount - 1;
			while (col > 0)
			{
				if (this._columnIndex[col].PageCount == 0 || this._columnIndex[col]._pages[0].RowCount > 1 || this._columnIndex[col]._pages[0].Rows[0].Index > 0)
				{
					break;
				}
				col--;
			}
			toCol = this._columnIndex[col].Index;
			if (toCol == 0)
			{
				fromRow = fromCol = toRow = toCol = 0;
				return false;
			}
			fromRow = toRow = 0;

			for (int c = fromIndex; c < ColumnCount; c++)
			{
				int first, last;
				if (this._columnIndex[c].PageCount == 0)
					continue;
				if (this._columnIndex[c]._pages[0].RowCount > 0 && this._columnIndex[c]._pages[0].Rows[0].Index > 0)
				{
					first = this._columnIndex[c]._pages[0].IndexOffset + this._columnIndex[c]._pages[0].Rows[0].Index;
				}
				else
				{
					if (this._columnIndex[c]._pages[0].RowCount > 1)
					{
						first = this._columnIndex[c]._pages[0].IndexOffset + this._columnIndex[c]._pages[0].Rows[1].Index;
					}
					else if (this._columnIndex[c].PageCount > 1)
					{
						first = this._columnIndex[c]._pages[0].IndexOffset + this._columnIndex[c]._pages[1].Rows[0].Index;
					}
					else
					{
						first = 0;
					}
				}
				var lp = this._columnIndex[c].PageCount - 1;
				while (this._columnIndex[c]._pages[lp].RowCount == 0 && lp != 0)
				{
					lp--;
				}
				var p = this._columnIndex[c]._pages[lp];
				if (p.RowCount > 0)
				{
					last = p.IndexOffset + p.Rows[p.RowCount - 1].Index;
				}
				else
				{
					last = first;
				}
				if (first > 0 && (first < fromRow || fromRow == 0))
				{
					fromRow = first;
				}
				if (first > 0 && (last > toRow || toRow == 0))
				{
					toRow = last;
				}
			}
			if (fromRow <= 0 || toRow <= 0)
			{
				fromRow = fromCol = toRow = toCol = 0;
				return false;
			}
			else
			{
				return true;
			}
		}
	}

	/// <summary>
	/// Shift the cells in this CellStore as a result of adding empty space.
	/// </summary>
	/// <param name="fromRow">The row to start inserting from.</param>
	/// <param name="fromCol">The column to start inserting from.</param>
	/// <param name="rows">The number of rows to be inserted.</param>
	/// <param name="columns">The number of columns to be inserted.</param>
	public void Insert(int fromRow, int fromCol, int rows, int columns)
	{
		lock (this._columnIndex)
		{

			if (columns > 0)
			{
				var col = this.GetPosition(fromCol);
				if (col < 0)
				{
					col = ~col;
				}
				for (var c = col; c < ColumnCount; c++)
				{
					this._columnIndex[c].Index += (short)columns;
				}
			}
			else
			{
				var page = fromRow >> pageBits;
				for (int c = 0; c < ColumnCount; c++)
				{
					var column = this._columnIndex[c];
					var pagePos = column.GetPosition(fromRow);
					if (pagePos >= 0)
					{
						if (fromRow >= column._pages[pagePos].MinIndex && fromRow <= column._pages[pagePos].MaxIndex) //The row is inside the page
						{
							int offset = fromRow - column._pages[pagePos].IndexOffset;
							var rowPos = column._pages[pagePos].GetPosition(offset);
							if (rowPos < 0)
							{
								rowPos = ~rowPos;
							}
							this.UpdateIndexOffset(column, pagePos, rowPos, fromRow, rows);
						}
						else if (column._pages[pagePos].MinIndex > fromRow - 1 && pagePos > 0) //The row is on the page before.
						{
							int offset = fromRow - ((page - 1) << pageBits);
							var rowPos = column._pages[pagePos - 1].GetPosition(offset);
							if (rowPos > 0 && pagePos > 0)
							{
								this.UpdateIndexOffset(column, pagePos - 1, rowPos, fromRow, rows);
							}
						}
						else if (column.PageCount >= pagePos + 1)
						{
							int offset = fromRow - column._pages[pagePos].IndexOffset;
							var rowPos = column._pages[pagePos].GetPosition(offset);
							if (rowPos < 0)
							{
								rowPos = ~rowPos;
							}
							if (column._pages[pagePos].RowCount > rowPos)
							{
								this.UpdateIndexOffset(column, pagePos, rowPos, fromRow, rows);
							}
							else
							{
								this.UpdateIndexOffset(column, pagePos + 1, 0, fromRow, rows);
							}
						}
					}
					else
					{
						this.UpdateIndexOffset(column, ~pagePos, 0, fromRow, rows);
					}
				}
			}
		}
#if DEBUGGING
		this.AssertInvariants();
#endif
	}

	/// <summary>
	/// Gets a default enumerator for the cellstore.
	/// </summary>
	/// <returns>The default enumerator for the cellstore.</returns>
	public ICellStoreEnumerator<T> GetEnumerator()
	{
		return new CellStore<T>.CellsStoreEnumerator<T>(this);
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
		return new CellStore<T>.CellsStoreEnumerator<T>(this, startRow, startColumn, endRow, endColumn);
	}
	#endregion

	#region Private Methods
	private void Delete(int fromRow, int fromCol, int rows, int columns, bool shift)
	{
		lock (this._columnIndex)
		{
			if (columns > 0 && fromRow == 0 && rows >= ExcelPackage.MaxRows)
			{
				this.DeleteColumns(fromCol, columns, shift);
			}
			else
			{
				var toCol = fromCol + columns - 1;
				var pageFromRow = fromRow >> pageBits;
				for (int c = 0; c < this.ColumnCount; c++)
				{
					int rowsToDelete = rows;
					var column = this._columnIndex[c];
					if (column.Index >= fromCol)
					{
						if (column.Index > toCol)
							break;
						var pagePos = column.GetPosition(fromRow);
						if (pagePos < 0) pagePos = ~pagePos;
						if (pagePos < column.PageCount)
						{
							var page = column._pages[pagePos];
							if (shift && page.RowCount > 0 && page.MinIndex > fromRow && page.MaxIndex >= fromRow + rowsToDelete)
							{
								// The entire page is being shifted.
								var o = page.MinIndex - fromRow;
								if (o < rowsToDelete)
								{
									rowsToDelete -= o;
									page.Offset -= o;
									this.UpdatePageOffset(column, pagePos, o);
								}
								else
								{
									page.Offset -= rowsToDelete;
									this.UpdatePageOffset(column, pagePos, rowsToDelete);
									continue;
								}
							}
							if (page.RowCount > 0 && page.MinIndex <= fromRow + rowsToDelete - 1 && page.MaxIndex >= fromRow)
							{
								// The range starts within a page: shift rows within the page. 
								var endRow = fromRow + rowsToDelete;
								var delEndRow = this.DeleteCells(column._pages[pagePos], fromRow, endRow, shift);
								if (shift && delEndRow != fromRow)
									this.UpdatePageOffset(column, pagePos, delEndRow - fromRow);
								if (endRow > delEndRow && pagePos < column.PageCount && column._pages[pagePos].MinIndex < endRow)
								{
									pagePos = (delEndRow == fromRow ? pagePos : pagePos + 1);
									var rowsLeft = this.DeletePage(shift ? fromRow : delEndRow, endRow - delEndRow, column, pagePos, shift);
									//if (shift) UpdatePageOffset(column, pagePos, endRow - fromRow - rowsLeft);
									if (rowsLeft > 0)
									{
										var fr = shift ? fromRow : endRow - rowsLeft;
										pagePos = column.GetPosition(fr);
										// This should never be the same page we deleted from earlier in this method.
										// Rather, it should be the next [remaining] page.
										// The only valid same-index case is when the entire original page was deleted. 
										if (page == column._pages[pagePos] && column._pages[pagePos + 1] != null)
											pagePos++;
										delEndRow = this.DeleteCells(column._pages[pagePos], fr, shift ? fr + rowsLeft : endRow, shift);
										if (shift)
											this.UpdatePageOffset(column, pagePos, rowsLeft);
									}
								}
							}
							else if (pagePos > 0 && column._pages[pagePos].IndexOffset > fromRow) //The row is on the page before.
							{
								int offset = fromRow + rowsToDelete - 1 - ((pageFromRow - 1) << pageBits);
								var rowPos = column._pages[pagePos - 1].GetPosition(offset);
								if (rowPos > 0 && pagePos > 0)
								{
									if (shift)
										this.UpdateIndexOffset(column, pagePos - 1, rowPos, fromRow + rowsToDelete - 1, -rowsToDelete);
								}
							}
							else
							{
								if (shift && pagePos + 1 < column.PageCount)
									this.UpdateIndexOffset(column, pagePos + 1, 0, column._pages[pagePos + 1].MinIndex, -rowsToDelete);
							}
						}
					}
				}
			}
		}
#if DEBUGGING
		this.AssertInvariants();
#endif
	}

	private int GetPointer(int Row, int Column)
	{
		var col = GetPosition(Column);
		if (col >= 0)
		{
			var pos = _columnIndex[col].GetPosition(Row);
			if (pos >= 0 && pos < _columnIndex[col].PageCount)
			{
				var pageItem = _columnIndex[col]._pages[pos];
				if (pageItem.MinIndex > Row)
				{
					pos--;
					if (pos < 0)
					{
						return -1;
					}
					else
					{
						pageItem = _columnIndex[col]._pages[pos];
					}
				}
				short ix = (short)(Row - pageItem.IndexOffset);
				_searchItem.Index = ix;
				var cellPos = Array.BinarySearch(pageItem.Rows, 0, pageItem.RowCount, _searchItem);
				if (cellPos >= 0)
				{
					return pageItem.Rows[cellPos].IndexPointer;
				}
				else //Cell does not exist
				{
					return -1;
				}
			}
			else //Page does not exist
			{
				return -1;
			}
		}
		else //Column does not exist
		{
			return -1;
		}
	}

	private bool NextCell(ref int row, ref int col, int minRow, int minColPos, int maxRow, int maxColPos)
	{
		if (minColPos >= ColumnCount)
		{
			return false;
		}
		if (maxColPos >= ColumnCount)
		{
			maxColPos = ColumnCount - 1;
		}
		var c = this.GetPosition(col);
		if (c >= 0)
		{
			if (c > maxColPos)
			{
				if (col <= minColPos)
				{
					return false;
				}
				col = minColPos;
				return this.NextCell(ref row, ref col);
			}
			else
			{
				var r = this.GetNextCell(ref row, ref c, minColPos, maxRow, maxColPos);
				col = this._columnIndex[c].Index;
				return r;
			}
		}
		else
		{
			c = ~c;
			if (c >= ColumnCount)
				c = this.ColumnCount - 1;
			if (col > this._columnIndex[c].Index)
			{
				if (col <= minColPos)
				{
					return false;
				}
				col = minColPos;
				return this.NextCell(ref row, ref col, minRow, minColPos, maxRow, maxColPos);
			}
			else
			{
				var r = this.GetNextCell(ref row, ref c, minColPos, maxRow, maxColPos);
				col = this._columnIndex[c].Index;
				return r;
			}
		}
	}

	private bool PrevCell(ref int row, ref int col, int minRow, int minColPos, int maxRow, int maxColPos)
	{
		if (minColPos >= this.ColumnCount)
		{
			return false;
		}
		if (maxColPos >= this.ColumnCount)
		{
			maxColPos = this.ColumnCount - 1;
		}
		var c = this.GetPosition(col);
		if (c >= 0)
		{
			if (c == 0)
			{
				if (col >= maxColPos)
				{
					return false;
				}
				if (row == minRow)
				{
					return false;
				}
				row--;
				col = maxColPos;
				return this.PrevCell(ref row, ref col, minRow, minColPos, maxRow, maxColPos);
			}
			else
			{
				var ret = this.GetPrevCell(ref row, ref c, minRow, minColPos, maxColPos);
				if (ret)
				{
					col = this._columnIndex[c].Index;
				}
				return ret;
			}
		}
		else
		{
			c = ~c;
			if (c == 0)
			{
				if (col >= maxColPos || row <= 0)
				{
					return false;
				}
				col = maxColPos;
				row--;
				return this.PrevCell(ref row, ref col, minRow, minColPos, maxRow, maxColPos);
			}
			else
			{
				var ret = this.GetPrevCell(ref row, ref c, minRow, minColPos, maxColPos);
				if (ret)
				{
					col = this._columnIndex[c].Index;
				}
				return ret;
			}
		}
	}

	private static int GetSize(int size)
	{
		var newSize = 256;
		while (newSize < size)
		{
			newSize <<= 1;
		}
		return newSize;
	}

	private void UpdatePageOffset(ColumnIndex column, int pagePos, int rows)
	{
		if (++pagePos < column.PageCount)
		{
			for (int p = pagePos; p < column.PageCount; p++)
			{
				if (column._pages[p].Offset - rows <= -PageSize)
				{
					column._pages[p].Index--;
					column._pages[p].Offset -= rows - PageSize;
				}
				else
				{
					column._pages[p].Offset -= rows;
				}
			}

			if (Math.Abs(column._pages[pagePos].Offset) > PageSize ||
				Math.Abs(column._pages[pagePos].Rows[column._pages[pagePos].RowCount - 1].Index) > PageSizeMax) //Split or Merge???
			{
				rows = ResetPageOffset(column, pagePos, rows);
				return;
			}
		}
	}

	private int ResetPageOffset(ColumnIndex column, int pagePos, int rows)
	{
		PageIndex fromPage = column._pages[pagePos];
		PageIndex toPage;
		short pageAdd = 0;
		if (fromPage.Offset < -PageSize)
		{
			toPage = column._pages[pagePos - 1];
			pageAdd = -1;
			if (fromPage.Index - 1 == toPage.Index)
			{
				if (fromPage.IndexOffset + fromPage.Rows[fromPage.RowCount - 1].Index -
					toPage.IndexOffset + toPage.Rows[0].Index <= PageSizeMax)
				{
					this.MergePage(column, pagePos - 1);
				}
			}
			else //No page after 
			{
				fromPage.Index -= pageAdd;
				fromPage.Offset += PageSize;
			}
		}
		else if (fromPage.Offset > PageSize)
		{
			toPage = column._pages[pagePos + 1];
			pageAdd = 1;
			if (fromPage.Index + 1 != toPage.Index)
			{
				fromPage.Index += pageAdd;
				fromPage.Offset += PageSize;
			}
		}
		return rows;
	}

	private int DeletePage(int fromRow, int rows, ColumnIndex column, int pagePos, bool shift)
	{
		PageIndex page = column._pages[pagePos];
		var startRows = rows;
		while (page != null && page.MinIndex >= fromRow && ((shift && page.MaxIndex < fromRow + rows) || (!shift && page.MaxIndex < fromRow + startRows)))
		{
			//Delete entire page.
			var delSize = page.MaxIndex - page.MinIndex + 1;
			rows -= delSize;
			var prevOffset = page.Offset;
			Array.Copy(column._pages, pagePos + 1, column._pages, pagePos, column.PageCount - pagePos + 1);
			column.PageCount--;
			if (column.PageCount == 0)
			{
				return 0;
			}
			if (shift)
			{
				for (int i = pagePos; i < column.PageCount; i++)
				{
					column._pages[i].Offset -= delSize;
					if (column._pages[i].Offset <= -PageSize)
					{
						column._pages[i].Index--;
						column._pages[i].Offset += PageSize;
					}
				}
			}
			if (column.PageCount > pagePos)
			{
				page = column._pages[pagePos];
			}
			else
			{
				//No more pages, return 0
				return 0;
			}
		}
		return rows;
	}

	private int DeleteCells(PageIndex page, int fromRow, int toRow, bool shift)
	{
		if (page.RowCount == 0)
			return fromRow;
		var fromPos = page.GetPosition(fromRow - (page.IndexOffset));
		if (fromPos < 0)
		{
			fromPos = ~fromPos;
		}
		var maxRow = page.MaxIndex;
		var offset = toRow - page.IndexOffset;
		if (offset > PageSizeMax) offset = PageSizeMax;
		var toPos = page.GetPosition(offset);
		if (toPos < 0)
		{
			toPos = ~toPos;
		}

		if (fromPos <= toPos && fromPos < page.RowCount && page.GetIndex(fromPos) < toRow)
		{
			if (toRow > page.MaxIndex)
			{
				if (fromRow == page.MinIndex) //Delete entire page, late in the page delete method
				{
					return fromRow;
				}
				var r = page.MaxIndex;
				var deletedRow = page.RowCount - fromPos;
				page.RowCount -= deletedRow;
				return r + 1;
			}
			else
			{
				var rows = toRow - fromRow;
				if (shift) UpdateRowIndex(page, toPos, rows);
				Array.Copy(page.Rows, toPos, page.Rows, fromPos, page.RowCount - toPos);
				page.RowCount -= toPos - fromPos;

				return toRow;
			}
		}
		else if (shift)
		{
			CellStore<T>.UpdateRowIndex(page, toPos, toRow - fromRow);
		}
		return toRow < maxRow ? toRow : maxRow;
	}

	private static void UpdateRowIndex(PageIndex page, int toPos, int rows)
	{
		for (int r = toPos; r < page.RowCount; r++)
		{
			page.Rows[r].Index -= (short)rows;
		}
	}

	private void DeleteColumns(int fromCol, int columns, bool shift)
	{
		var fPos = this.GetPosition(fromCol);
		if (fPos < 0)
		{
			fPos = ~fPos;
		}
		int tPos = fPos;
		for (var c = fPos; c <= ColumnCount; c++)
		{
			tPos = c;
			if (tPos == this.ColumnCount || this._columnIndex[c].Index >= fromCol + columns)
			{
				break;
			}
		}

		if (this.ColumnCount <= fPos)
		{
			return;
		}

		if (_columnIndex[fPos].Index >= fromCol && this._columnIndex[fPos].Index <= fromCol + columns)
		{
			if (tPos < this.ColumnCount)
			{
				Array.Copy(this._columnIndex, tPos, this._columnIndex, fPos, ColumnCount - tPos);
			}
			this.ColumnCount -= (tPos - fPos);
		}
		if (shift)
		{
			for (var c = fPos; c < this.ColumnCount; c++)
			{
				this._columnIndex[c].Index -= (short)columns;
			}
		}
	}

	private void UpdateIndexOffset(ColumnIndex column, int pagePos, int rowPos, int row, int rows)
	{
		if (pagePos >= column.PageCount) return;    //A page after last cell.
		var page = column._pages[pagePos];
		if (rows > PageSize)
		{
			short addPages = (short)(rows >> pageBits);
			int offset = +(int)(rows - (PageSize * addPages));
			for (int p = pagePos + 1; p < column.PageCount; p++)
			{
				if (column._pages[p].Offset + offset > PageSize)
				{
					column._pages[p].Index += (short)(addPages + 1);
					column._pages[p].Offset += offset - PageSize;
				}
				else
				{
					column._pages[p].Index += addPages;
					column._pages[p].Offset += offset;
				}

			}

			var size = page.RowCount - rowPos;
			if (page.RowCount > rowPos)
			{
				if (column.PageCount - 1 == pagePos) //No page after, create a new one.
				{
					//Copy rows to next page.
					var newPage = CopyNew(page, rowPos, size);
					newPage.Index = (short)((row + rows) >> pageBits);
					newPage.Offset = row + rows - (newPage.Index * PageSize) - newPage.Rows[0].Index;
					if (newPage.Offset > PageSize)
					{
						newPage.Index++;
						newPage.Offset -= PageSize;
					}
					this.AddPage(column, pagePos + 1, newPage);
					page.RowCount = rowPos;
				}
				else
				{
					if (column._pages[pagePos + 1].RowCount + size > PageSizeMax) //Split Page
					{
						this.SplitPageInsert(column, pagePos, rowPos, rows, size, addPages);
					}
					else //Copy Page.
					{
						this.CopyMergePage(page, rowPos, rows, size, column._pages[pagePos + 1]);
					}
				}
			}
		}
		else
		{
			//Add to Pages.
			for (int r = rowPos; r < page.RowCount; r++)
			{
				page.Rows[r].Index += (short)rows;
			}
			if (page.Offset + page.Rows[page.RowCount - 1].Index >= PageSizeMax)   //Can not be larger than the max size of the page.
			{
				this.AdjustIndex(column, pagePos);
				if (page.Offset + page.Rows[page.RowCount - 1].Index >= PageSizeMax)
				{
					pagePos = SplitPage(column, pagePos);
				}
			}

			for (int p = pagePos + 1; p < column.PageCount; p++)
			{
				if (column._pages[p].Offset + rows < PageSize)
				{
					column._pages[p].Offset += rows;
				}
				else
				{
					column._pages[p].Index++;
					column._pages[p].Offset = (column._pages[p].Offset + rows) % PageSize;
				}
			}
		}
	}

	private void SplitPageInsert(ColumnIndex column, int pagePos, int rowPos, int rows, int size, int addPages)
	{
		var newRows = new IndexItem[GetSize(size)];
		var page = column._pages[pagePos];

		var rStart = -1;
		for (int r = rowPos; r < page.RowCount; r++)
		{
			if (page.IndexExpanded - (page.Rows[r].Index + rows) > PageSize)
			{
				rStart = r;
				break;
			}
			else
			{
				page.Rows[r].Index += (short)rows;
			}
		}
		var rc = page.RowCount - rStart;
		page.RowCount = rStart;
		if (rc > 0)
		{
			//Copy to a new page
			var row = page.IndexOffset;
			var newPage = CopyNew(page, rStart, rc);
			var ix = (short)(page.Index + addPages);
			var offset = page.IndexOffset + rows - (ix * PageSize);
			if (offset > PageSize)
			{
				ix += (short)(offset / PageSize);
				offset %= PageSize;
			}
			newPage.Index = ix;
			newPage.Offset = offset;
			this.AddPage(column, pagePos + 1, newPage);
		}
	}

	private void CopyMergePage(PageIndex page, int rowPos, int rows, int size, PageIndex toPage)
	{
		var startRow = page.IndexOffset + page.Rows[rowPos].Index + rows;
		var newRows = new IndexItem[GetSize(toPage.RowCount + size)];
		page.RowCount -= size;
		Array.Copy(page.Rows, rowPos, newRows, 0, size);
		for (int r = 0; r < size; r++)
		{
			newRows[r].Index += (short)(page.IndexOffset + rows - toPage.IndexOffset);
		}

		Array.Copy(toPage.Rows, 0, newRows, size, toPage.RowCount);
		toPage.Rows = newRows;
		toPage.RowCount += size;
	}

	private void MergePage(ColumnIndex column, int pagePos)
	{
		PageIndex page1 = column._pages[pagePos];
		PageIndex page2 = column._pages[pagePos + 1];

		var newPage = new PageIndex(page1, 0, page1.RowCount + page2.RowCount);
		newPage.RowCount = page1.RowCount + page2.RowCount;
		Array.Copy(page1.Rows, 0, newPage.Rows, 0, page1.RowCount);
		Array.Copy(page2.Rows, 0, newPage.Rows, page1.RowCount, page2.RowCount);
		for (int r = page1.RowCount; r < newPage.RowCount; r++)
		{
			newPage.Rows[r].Index += (short)(page2.IndexOffset - page1.IndexOffset);
		}

		column._pages[pagePos] = newPage;
		column.PageCount--;

		if (column.PageCount > (pagePos + 1))
		{
			Array.Copy(column._pages, pagePos + 2, column._pages, pagePos + 1, column.PageCount - (pagePos + 1));
			for (int p = pagePos + 1; p < column.PageCount; p++)
			{
				column._pages[p].Index--;
				column._pages[p].Offset += PageSize;
			}
		}
	}

	private PageIndex CopyNew(PageIndex pageFrom, int rowPos, int size)
	{
		IndexItem[] newRows = new IndexItem[GetSize(size)];
		Array.Copy(pageFrom.Rows, rowPos, newRows, 0, size);
		return new PageIndex(newRows, size);
	}

	private void AddCell(ColumnIndex columnIndex, int pagePos, int pos, short ix, T value)
	{
		PageIndex pageItem = columnIndex._pages[pagePos];
		if (pageItem.RowCount == pageItem.Rows.Length)
		{
			if (pageItem.RowCount == CellStore<T>.PageSizeMax) //Max size-->Split
			{

				pagePos = this.SplitPage(columnIndex, pagePos);
				// Should the new value be stored on the previous page?
				if (columnIndex._pages[pagePos - 1].RowCount > pos)
				{
					pagePos--;
				}
				else
				{
					pos -= columnIndex._pages[pagePos - 1].RowCount;
				}
				pageItem = columnIndex._pages[pagePos];
				ix = (short)(ix - pageItem.IndexOffset);
			}
			else //Expand to double size.
			{
				var rowsTmp = new IndexItem[pageItem.Rows.Length << 1];
				Array.Copy(pageItem.Rows, 0, rowsTmp, 0, pageItem.RowCount);
				pageItem.Rows = rowsTmp;
			}
		}
		if (pos < pageItem.RowCount)
		{
			Array.Copy(pageItem.Rows, pos, pageItem.Rows, pos + 1, pageItem.RowCount - pos);
		}
		pageItem.Rows[pos] = new IndexItem() { Index = ix, IndexPointer = _values.Count };
		this._values.Add(value);
		pageItem.RowCount++;
	}

	private int SplitPage(ColumnIndex columnIndex, int pagePos)
	{
		var page = columnIndex._pages[pagePos];
		if (page.Offset != 0)
		{
			var offset = page.Offset;
			page.Offset = 0;
			for (int r = 0; r < page.RowCount; r++)
			{
				page.Rows[r].Index += (short)offset;
			}
		}
		//Find Split pos
		int splitPos = 0;
		for (int r = 0; r < page.RowCount; r++)
		{
			if (page.Rows[r].Index > PageSize)
			{
				splitPos = r;
				break;
			}
		}
		var newPage = new PageIndex(page, 0, splitPos);
		var nextPage = new PageIndex(page, splitPos, page.RowCount - splitPos, (short)(page.Index + 1), page.Offset);

		for (int r = 0; r < nextPage.RowCount; r++)
		{
			nextPage.Rows[r].Index = (short)(nextPage.Rows[r].Index - PageSize);
		}

		columnIndex._pages[pagePos] = newPage;
		if (columnIndex.PageCount + 1 > columnIndex._pages.Length)
		{
			var pageTmp = new PageIndex[columnIndex._pages.Length << 1];
			Array.Copy(columnIndex._pages, 0, pageTmp, 0, columnIndex.PageCount);
			columnIndex._pages = pageTmp;
		}
		Array.Copy(columnIndex._pages, pagePos + 1, columnIndex._pages, pagePos + 2, columnIndex.PageCount - pagePos - 1);
		columnIndex._pages[pagePos + 1] = nextPage;
		page = nextPage;
		columnIndex.PageCount++;
		return pagePos + 1;
	}

	private bool AdjustIndex(ColumnIndex columnIndex, int pagePos)
	{
		PageIndex page = columnIndex._pages[pagePos];
		//First Adjust indexes
		if (page.Offset + page.Rows[0].Index >= PageSize ||
			page.Offset >= PageSize ||
			page.Rows[0].Index >= PageSize)
		{
			page.Index++;
			if (page.Offset > 0)
				page.Offset -= PageSize;
			else
				for (int i = 0; i < page.RowCount; i++)
				{
					page.Rows[i].Index -= PageSize;
				}
			return true;
		}
		else if (page.Offset + page.Rows[0].Index <= -PageSize ||
				 page.Offset <= -PageSize ||
				 page.Rows[0].Index <= -PageSize)
		{
			page.Index--;
			if (page.Offset < 1)
				page.Offset += PageSize;
			else
				for (int i = 0; i < page.RowCount; i++)
				{
					page.Rows[i].Index += PageSize;
				}
			return true;
		}
		return false;
	}

	private void AddPage(ColumnIndex column, int pos, short index)
	{
		AddPage(column, pos);
		column._pages[pos] = new PageIndex() { Index = index };
		if (pos > 0)
		{
			var pp = column._pages[pos - 1];
			if (pp.RowCount > 0 && pp.Rows[pp.RowCount - 1].Index > PageSize)
			{
				column._pages[pos].Offset = pp.Rows[pp.RowCount - 1].Index - PageSize;
			}
		}
	}

	/// <summary>
	/// Add a new page to the collection
	/// </summary>
	/// <param name="column">The column</param>
	/// <param name="pos">Position</param>
	/// <param name="page">The new page object to add</param>
	private void AddPage(ColumnIndex column, int pos, PageIndex page)
	{
		this.AddPage(column, pos);
		column._pages[pos] = page;
	}

	/// <summary>
	/// Add a new page to the collection
	/// </summary>
	/// <param name="column">The column</param>
	/// <param name="pos">Position</param>
	private void AddPage(ColumnIndex column, int pos)
	{
		if (column.PageCount == column._pages.Length)
		{
			var pageTmp = new PageIndex[column._pages.Length * 2];
			Array.Copy(column._pages, 0, pageTmp, 0, column.PageCount);
			column._pages = pageTmp;
		}
		if (pos < column.PageCount)
		{
			Array.Copy(column._pages, pos, column._pages, pos + 1, column.PageCount - pos);
		}
		column.PageCount++;
	}

	/// <summary>
	/// Get the index into the columns array that corresponds to the specified <paramref name="column"/>.
	/// </summary>
	/// <param name="column">The column to find the location of.</param>
	/// <returns></returns>
	private int GetPosition(int column)
	{
		this._searchIx.Index = (short)column;
		return Array.BinarySearch(this._columnIndex, 0, this.ColumnCount, this._searchIx);
	}

	/// <summary>
	/// Update the row and column position to reflect the first cell before the current position.
	/// </summary>
	/// <param name="row">The current row / the resulting row.</param>
	/// <param name="colPos">The current column index / the resulting column index.</param>
	/// <param name="startRow">The row to start looking in.</param>
	/// <param name="startColPos">The column index to start looking in.</param>
	/// <param name="endColPos">The last column index to look in.</param>
	/// <returns>True if a previous value was found; false otherwise.</returns>
	private bool GetPrevCell(ref int row, ref int colPos, int startRow, int startColPos, int endColPos)
	{
		if (this.ColumnCount == 0)
		{
			return false;
		}
		else
		{
			if (--colPos >= startColPos)
			{
				var r = _columnIndex[colPos].GetNextRow(row);
				if (r == row) //Exists next Row
				{
					return true;
				}
				else
				{
					int minRow, minCol;
					if (r > row && r >= startRow)
					{
						minRow = r;
						minCol = colPos;
					}
					else
					{
						minRow = int.MaxValue;
						minCol = 0;
					}

					var c = colPos - 1;
					if (c >= startColPos)
					{
						while (c >= startColPos)
						{
							r = this._columnIndex[c].GetNextRow(row);
							if (r == row) //Exists next Row
							{
								colPos = c;
								return true;
							}
							if (r > row && r < minRow && r >= startRow)
							{
								minRow = r;
								minCol = c;
							}
							c--;
						}
					}
					if (row > startRow)
					{
						c = endColPos;
						row--;
						while (c > colPos)
						{
							r = this._columnIndex[c].GetNextRow(row);
							if (r == row) //Exists next Row
							{
								colPos = c;
								return true;
							}
							if (r > row && r < minRow && r >= startRow)
							{
								minRow = r;
								minCol = c;
							}
							c--;
						}
					}
					if (minRow == int.MaxValue || startRow < minRow)
					{
						return false;
					}
					else
					{
						row = minRow;
						colPos = minCol;
						return true;
					}
				}
			}
			else
			{
				colPos = this.ColumnCount;
				row--;
				if (row < startRow)
				{
					return false;
				}
				else
				{
					return this.GetPrevCell(ref colPos, ref row, startRow, startColPos, endColPos);
				}
			}
		}
	}

	/// <summary>
	/// Get the next cell after the current position.
	/// </summary>
	/// <param name="row">The row to start searching in.</param>
	/// <param name="colPos">The resulting column position.</param>
	/// <param name="startColPos">The column position to start searching in.</param>
	/// <param name="endRow">The last row to search.</param>
	/// <param name="endColPos">The last column index to search.</param>
	/// <returns>True if the cell could be found; false otherwise.</returns>
	private bool GetNextCell(ref int row, ref int colPos, int startColPos, int endRow, int endColPos)
	{
		if (this.ColumnCount == 0)
		{
			return false;
		}
		else
		{
			if (++colPos < this.ColumnCount && colPos <= endColPos)
			{
				var r = this._columnIndex[colPos].GetNextRow(row);
				if (r == row) //Exists next Row
				{
					return true;
				}
				else
				{
					int minRow, minCol;
					if (r > row)
					{
						minRow = r;
						minCol = colPos;
					}
					else
					{
						minRow = int.MaxValue;
						minCol = 0;
					}

					var c = colPos + 1;
					while (c < this.ColumnCount && c <= endColPos)
					{
						r = this._columnIndex[c].GetNextRow(row);
						if (r == row) //Exists next Row
						{
							colPos = c;
							return true;
						}
						if (r > row && r < minRow)
						{
							minRow = r;
							minCol = c;
						}
						c++;
					}
					c = startColPos;
					if (row < endRow)
					{
						row++;
						while (c < colPos)
						{
							r = this._columnIndex[c].GetNextRow(row);
							if (r == row) //Exists next Row
							{
								colPos = c;
								return true;
							}
							if (r > row && (r < minRow || (r == minRow && c < minCol)) && r <= endRow)
							{
								minRow = r;
								minCol = c;
							}
							c++;
						}
					}

					if (minRow == int.MaxValue || minRow > endRow)
					{
						return false;
					}
					else
					{
						row = minRow;
						colPos = minCol;
						return true;
					}
				}
			}
			else
			{
				if (colPos <= startColPos || row >= endRow)
				{
					return false;
				}
				colPos = startColPos - 1;
				row++;
				return this.GetNextCell(ref row, ref colPos, startColPos, endRow, endColPos);
			}
		}
	}

	/// <summary>
	/// Gets the index into the cell array that corresponds to the particular column position.
	/// </summary>
	/// <param name="colPos">The column position whose index should be returned.</param>
	/// <returns>The index into the data array that corresponds to the given column position.</returns>
	private int GetColumnPosIndex(int colPos)
	{
		return this._columnIndex[colPos].Index;
	}

	private void AddColumn(int pos, int Column)
	{
		if (this.ColumnCount == this._columnIndex.Length)
		{
			var colTmp = new ColumnIndex[this._columnIndex.Length * 2];
			Array.Copy(_columnIndex, 0, colTmp, 0, ColumnCount);
			this._columnIndex = colTmp;
		}
		if (pos < ColumnCount)
		{
			Array.Copy(this._columnIndex, pos, _columnIndex, pos + 1, ColumnCount - pos);
		}
		this._columnIndex[pos] = new ColumnIndex() { Index = (short)(Column) };
		this.ColumnCount++;
	}

#if DEBUGGING
	private void AssertInvariants()
	{
		int index = -1;
		for (int i = 0; i < this.ColumnCount; i++)
		{
			var column = this._columnIndex[i];
			if (column.Index <= index)
				throw new InvalidOperationException($"The column {column.Index} violates the invariant that the list of columns is always sorted.");
			index = column.Index;
			this.AssertPageInvariants(column);
		}
	}

	private void AssertPageInvariants(ColumnIndex column)
	{
		int lastIndex = -1;
		for (int i = 0; i < column.PageCount; i++)
		{
			var page = column._pages[i];
			if (Math.Abs(page.Offset) > CellStore<T>.PageSize)
				throw new InvalidOperationException($"The offset {page.Offset} is too extreme, must be in the range (-{CellStore<T>.PageSize} to {CellStore<T>.PageSize}).");
			if (page.RowCount != 0 && page.MaxIndex < page.MinIndex)
				throw new InvalidOperationException($"Page {i} is not sorted: begins with {page.MinIndex} and ends with {page.MaxIndex}.");
			else if (page.MinIndex <= lastIndex)
				throw new InvalidOperationException($"The page {i} on column.Index {column.Index} is out of order.");
			lastIndex = page.MinIndex;
		}
	}
#endif
	#endregion

}

internal class FlagCellStore : CellStore<byte>, IFlagStore
{
	#region Public Methods
	public void SetFlagValue(int Row, int Col, bool value, CellFlags cellFlags)
	{
		CellFlags currentValue = (CellFlags)GetValue(Row, Col);
		if (value)
		{
			SetValue(Row, Col, (byte)(currentValue | cellFlags)); // add the CellFlag bit
		}
		else
		{
			SetValue(Row, Col, (byte)(currentValue & ~cellFlags)); // remove the CellFlag bit
		}
	}
	public bool GetFlagValue(int Row, int Col, CellFlags cellFlags)
	{
		return !(((byte)cellFlags & GetValue(Row, Col)) == 0);
	}
	#endregion
}
