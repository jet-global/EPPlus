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
using System.Text;
using System.Threading.Tasks;

namespace OfficeOpenXml
{
	internal class ZCellStore<T> : ICellStore<T>
	{
		#region Constants
		private const int RowPageSize = 1024; // The number of rows in a row page
		private const int ColumnPageSize = 128; // The number of columns in a column page
		private const int RowMask = 1023; // 111111111 ( 9 1's )
		private const int ColumnMask = 127; // 111111 ( 6 1's)
		private const int RowPageBits = 10; // 2 ^ 10 = 1024 so right-shifting by 10 bits divides by 1024
		private const int ColumnPageBits = 7; // 2 ^ 7 = 128 so right-shifting by 7 bits divides by 128
		#endregion

		#region Properties
		private PagedStructure<PagedStructure<T>> Data { get; } = new PagedStructure<PagedStructure<T>>(ZCellStore<T>.ColumnPageBits);
		#endregion

		#region ICellStore<T> Members
		public T GetValue(int row, int column)
		{
			if (this.TryUpdateIndices(ref row, ref column, true))
			{
				var columnItem = this.Data.GetItem(column)?.Value;
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
			if (this.TryUpdateIndices(ref row, ref column, true))
			{
				var columnStructure = this.Data.GetItem(column);
				if (columnStructure == null)
				{
					columnStructure = new PagedStructure<T>(ZCellStore<T>.RowPageBits);
					this.Data.SetItem(column, columnStructure);
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
			if (this.TryUpdateIndices(ref row, ref column, true))
			{
				var item = this.Data.GetItem(column)?.Value.GetItem(row);
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
			int betaRow = ExcelPackage.MaxRows + 1, betaColumn = -1;
			bool found = this.TryFindNextCell(column, ExcelPackage.MaxColumns, row, ref betaRow, ref betaColumn) ||
									 this.TryFindNextCell(0, column, row + 1, ref betaRow, ref betaColumn);
			row = betaRow + 1;
			column = betaColumn + 1;
			return found || betaRow != ExcelPackage.MaxRows + 1;
		}

		public bool PrevCell(ref int row, ref int column)
		{
			int betaRow = -1, betaColumn = -1;
			bool found = this.TryFindPreviousCell(column, 1, row, ref betaRow, ref betaColumn) ||
									 this.TryFindPreviousCell(ExcelPackage.MaxColumns + 1, column, row - 1, ref betaRow, ref betaColumn);
			row = betaRow + 1;
			column = betaColumn + 1;
			return found || betaRow != -1;
		}

		public void Delete(int fromRow, int fromCol, int rows, int columns)
		{
			if (this.TryUpdateIndices(ref fromRow, ref fromCol, false))
			{
				if (fromCol >= 0)
					this.Data.ShiftItems(fromCol, -columns);
				if (fromRow >= 0)
				{
					int column = fromCol - 1; // Only from the column forward // TODO ZPF test this
					while (this.Data.NextItem(ref column))
					{
						this.Data.GetItem(column)?.Value.ShiftItems(fromRow, -rows);
					}
				}
			}
		}

		public void Delete(int fromRow, int fromCol, int rows, int columns, bool shift)
		{
			if (shift)
				this.Delete(fromRow, fromCol, rows, columns);
			else
				this.Clear(fromRow, fromCol, rows, columns);
		}

		public void Clear(int fromRow, int fromCol, int rows, int columns)
		{
			if (this.TryUpdateIndices(ref fromRow, ref fromCol, false))
			{
				if (fromCol >= 0)
					this.Data.ClearItems(fromCol, columns);
				if (fromRow >= 0)
				{
					int column = fromCol - 1; // Only from the column forward // TODO ZPF test this
					while (this.Data.NextItem(ref column))
					{
						this.Data.GetItem(column)?.Value.ClearItems(fromRow, rows);
					}
				}
			}
		}

		public bool GetDimension(out int fromRow, out int fromCol, out int toRow, out int toCol)
		{
			fromCol = this.Data.MinimumUsedIndex + 1;
			toCol = this.Data.MaximumUsedIndex + 1;
			fromRow = ExcelPackage.MaxRows + 1;
			toRow = 0;
			int searchIndex = 0;
			while (this.Data.NextItem(ref searchIndex))
			{
				var column = this.Data.GetItem(searchIndex);
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
			if (this.TryUpdateIndices(ref fromRow, ref fromCol, false))
			{
				if (fromCol >= 0)
					this.Data.ShiftItems(fromCol, columns);
				if (fromRow >= 0)
				{
					int column = fromCol - 1; // Only from the column forward // TODO ZPF test this
					while (this.Data.NextItem(ref column))
					{
						this.Data.GetItem(column)?.Value.ShiftItems(fromRow, rows);
					}
				}
			}
		}

		public void Dispose()
		{
			// TODO ZPF do we really need this?...
		}

		public ICellStoreEnumerator<T> GetEnumerator()
		{
			return new ZCellStoreEnumerator(this);
		}

		public ICellStoreEnumerator<T> GetEnumerator(int startRow, int startColumn, int endRow, int endColumn)
		{
			return new ZCellStoreEnumerator(this, startRow, startColumn, endRow, endColumn);
		}
		#endregion

		#region Private Methods
		private bool TryFindNextCell(int startColumn, int endColumn, int targetRow, ref int betaRow, ref int betaColumn)
		{
			startColumn--; // 1-indexing offset
			endColumn--; // 1-indexing offset
			targetRow--; // 1-indexing offset
			while (this.Data.NextItem(ref startColumn) && startColumn <= endColumn)
			{
				var page = this.Data.GetItem(startColumn);
				int rowSearch = targetRow - 1;
				if (page?.Value.NextItem(ref rowSearch) == true)
				{
					if (rowSearch == targetRow || targetRow < 0)
					{
						betaRow = rowSearch;
						betaColumn = startColumn;
						return true;
					}
					else if (rowSearch < betaRow)
					{
						betaRow = rowSearch;
						betaColumn = startColumn;
					}
				}
			}
			return false;
		}

		private bool TryFindPreviousCell(int startColumn, int endColumn, int targetRow, ref int betaRow, ref int betaColumn)
		{
			startColumn--; // 1-indexing offset
			endColumn--; // 1-indexing offset
			targetRow--; // 1-indexing offset
			while (this.Data.PreviousItem(ref startColumn) && startColumn >= endColumn)
			{
				var page = this.Data.GetItem(startColumn);
				int rowSearch = targetRow + 1;
				if (page?.Value.PreviousItem(ref rowSearch) == true)
				{
					if (rowSearch == targetRow || targetRow >= ExcelPackage.MaxRows)
					{
						betaRow = rowSearch;
						betaColumn = startColumn;
						return true;
					}
					else if (rowSearch > betaRow)
					{
						betaRow = rowSearch;
						betaColumn = startColumn;
					}
				}
			}
			return false;
		}

		private bool TryUpdateIndices(ref int row, ref int column, bool validate)
		{
			if (validate && (row < 1 || row > ExcelPackage.MaxRows || column < 1 || column > ExcelPackage.MaxColumns))
				return false;
			row--;
			column--;
			return true;
		}
		#endregion

		#region Nested Classes
		// This structure is 0-indexed
		internal class PagedStructure<S>
		{
			#region Properties
			public int PageBits { get; } // Defines a bit shift to retrive the primary index
			public int PageSize { get; } // Defines the max index of a single page
			public int PageMask { get; } // Defines a bit mask to retrieve the secondary index
			public int MaximumIndex { get; }
			public int MinimumUsedIndex { get; private set; }
			public int MaximumUsedIndex { get; private set; }

			public bool IsEmpty
			{
				get { return this.MinimumUsedIndex == this.MaximumIndex + 1 && this.MaximumUsedIndex == -1; }
			}

			private Page[] Pages { get; } 
			#endregion

			#region Constructors
			public PagedStructure(int pageBits)
			{
				this.PageBits = pageBits;
				this.PageSize = 1 << pageBits;
				this.PageMask = this.PageSize - 1;
				this.MaximumIndex = (this.PageSize << pageBits) - 1;
				this.Pages = new Page[this.PageSize];
				this.MinimumUsedIndex = this.MaximumIndex + 1;
				this.MaximumUsedIndex = -1;
			}
			#endregion

			#region Public Methods
			public ValueHolder? GetItem(int index)
			{
				if (index < 0 || index > this.MaximumIndex)
					return null;
				this.DeConstructIndex(index, out int page, out int innerIndex);
				var pageArray = this.Pages[page];
				if (null == pageArray)
					return null;
				return pageArray[innerIndex];
			}

			public void SetItem(int index, ValueHolder? item)
			{
				this.SetItem(index, item, true);
			}

			public void ShiftItems(int index, int amount)
			{
				if (index < 0 || index > this.MaximumIndex)
					return;
				if (amount > 0)
				{
					for (int i = this.MaximumUsedIndex; i >= index; --i)
					{
						var target = i + amount;
						if (target <= this.MaximumIndex)
							this.SetItem(target, this.GetItem(i), false);
						this.SetItem(i, null, false);
					}
				}
				else if (amount < 0)
				{
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
				// TODO cleanup? Only used for test setup/validation.
				public ValueHolder?[] GetValues()
				{
					return this.Values; ;
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
				this(zCellStore, 1, 1, ExcelPackage.MaxRows, ExcelPackage.MaxColumns)
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
				// TODO ZPF Can we just take this off of the interfact?...
			}

			public IEnumerator<T> GetEnumerator()
			{
				this.Reset();
				return this;
			}

			public bool MoveNext()
			{ 
				int betaRow = ExcelPackage.MaxRows + 1, betaColumn = -1;
				bool found = this.CellStore.TryFindNextCell(myColumn, this.Endcolumn, myRow, ref betaRow, ref betaColumn) ||
										 this.CellStore.TryFindNextCell(this.StartColumn - 1, myColumn, myRow + 1, ref betaRow, ref betaColumn);
				myRow = betaRow + 1;
				myColumn = betaColumn + 1;
				// TODO ZPF This validation can likely be optimized
				return myRow >= this.StartRow && myRow <= this.EndRow && myColumn >= this.StartColumn && myColumn <= this.Endcolumn; 
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
