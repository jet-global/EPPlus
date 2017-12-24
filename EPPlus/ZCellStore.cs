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
			if (row < 1 || row > ExcelPackage.MaxRows || column < 1 || column > ExcelPackage.MaxColumns)
				return default(T);
			var columnItem = this.Data.GetItem(column - 1)?.Value;
			if (columnItem != null)
			{
				var rowItem = columnItem.GetItem(row - 1);
				if (rowItem.HasValue)
					return rowItem.Value.Value;
			}
			return default(T);
		}

		public void SetValue(int row, int column, T value)
		{
			if (row < 1 || row > ExcelPackage.MaxRows || column < 1 || column > ExcelPackage.MaxColumns)
				return;
			column--;
			var columnStructure = this.Data.GetItem(column);
			if (columnStructure == null)
			{
				columnStructure = new PagedStructure<T>(ZCellStore<T>.RowPageBits);
				this.Data.SetItem(column, columnStructure);
			}
			columnStructure?.Value.SetItem(row - 1, value);
		}

		public bool Exists(int row, int column)
		{
			return this.Exists(row, column, out _);
		}

		public bool Exists(int row, int column, out T value)
		{
			value = default(T);
			if (row < 1 || row > ExcelPackage.MaxRows || column < 1 || column > ExcelPackage.MaxColumns)
				return false;
			var item = this.Data.GetItem(column - 1)?.Value.GetItem(row - 1);
			if (item.HasValue)
			{
				value = item.Value.Value;
				return true;
			}
			return false;
		}

		public bool NextCell(ref int row, ref int column)
		{
			int localRow = row - 2, localColumn = column - 1;
			var columnSearch = localColumn;
			int currentColumn = localColumn;
			int lastRow = localRow;
			int betaRow = ExcelPackage.MaxRows + 1;
			int betaColumn = 0;

			for (int i = 0; i < 2; ++i) // This still over-enumerates
			{
				while (this.Data.NextItem(ref columnSearch))
				{
					var rowSearch = localRow;
					if (this.Data.GetItem(columnSearch)?.Value.NextItem(ref rowSearch) == true)
					{
						if (rowSearch < betaRow)
						{
							betaRow = rowSearch;
							betaColumn = columnSearch;
						}
					}
				}
				columnSearch = -1;
				localRow++;
			}
			row = betaRow + 1;
			column = betaColumn + 1;
			return betaRow != ExcelPackage.MaxRows + 1;
		}

		public bool PrevCell(ref int row, ref int column)
		{
			int localRow = row, localColumn = column - 1;
			var columnSearch = localColumn;
			int currentColumn = localColumn;
			int lastRow = localRow;
			int betaRow = -1;
			int betaColumn = 0;

			for (int i = 0; i < 2; ++i) // This still over-enumerates
			{
				while (this.Data.PreviousItem(ref columnSearch))
				{
					var rowSearch = localRow;
					if (this.Data.GetItem(columnSearch)?.Value.PreviousItem(ref rowSearch) == true)
					{
						if (rowSearch > betaRow)
						{
							betaRow = rowSearch;
							betaColumn = columnSearch;
						}
					}

				}
				columnSearch = ExcelPackage.MaxColumns;
				localRow--;
			}
			row = betaRow + 1;
			column = betaColumn + 1;
			return betaRow != -1;
		}

		public void Delete(int fromRow, int fromCol, int rows, int columns)
		{
			this.Delete(fromRow, fromCol, rows, columns, true);
		}

		public void Delete(int fromRow, int fromCol, int rows, int columns, bool shift)
		{
			if (shift)
			{
				if (fromCol > 0)
					this.Data.ShiftItems(fromCol, -columns);
				if (fromRow > 0)
				{
					int column = -1;
					while (this.Data.NextItem(ref column))
					{
						this.Data.GetItem(column)?.Value.ShiftItems(fromRow, -rows);
					}
				}
			}
			else
				this.Clear(fromRow, fromCol, rows, columns);
		}

		public void Clear(int fromRow, int fromCol, int rows, int columns)
		{
			if (fromCol > 0)
				this.Data.ClearItems(fromCol, columns);
			if (fromRow > 0)
			{
				int column = -1;
				while (this.Data.NextItem(ref column))
				{
					this.Data.GetItem(column)?.Value.ClearItems(fromRow, rows);
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

		public void Insert(int rowFrom, int columnFrom, int rows, int columns)
		{
			if (columnFrom > 0)
				this.Data.ShiftItems(columnFrom, columns);
			if (rowFrom > 0)
			{
				int column = -1;
				while (this.Data.NextItem(ref column))
				{
					this.Data.GetItem(column)?.Value.ShiftItems(rowFrom, rows);
				}
			}
		}

		public void Dispose()
		{
			throw new NotImplementedException();
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

			private ValueHolder?[][] Pages { get; } // protected to enable a test helper to get visibility in here
			#endregion

			#region Constructors
			public PagedStructure(int pageBits)
			{
				this.PageBits = pageBits;
				this.PageSize = 1 << pageBits;
				this.PageMask = this.PageSize - 1;
				this.MaximumIndex = (this.PageSize << pageBits) - 1;
				this.Pages = new ValueHolder?[this.PageSize][];
				this.MinimumUsedIndex = -1;
				this.MaximumUsedIndex = -1;
			}
			#endregion

			#region Public Methods
			public ValueHolder? GetItem(int index)
			{
				if (index < 0 || index > this.MaximumIndex)
					return null;
				var page = index >> this.PageBits;
				var innerIndex = index & this.PageMask;
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
				if (this.MinimumUsedIndex == -1 || this.MaximumUsedIndex == -1)
					return false;

				if (index++ < this.MinimumUsedIndex)
					index = this.MinimumUsedIndex;

				var minimumPage = index >> this.PageBits;
				var minimumInnerIndex = index & this.PageMask;
				var maximumPage = this.MaximumUsedIndex >> this.PageBits;
				var maximumInnerIndex = this.MaximumUsedIndex & this.PageMask;

				for (int page = minimumPage; page <= maximumPage; ++page)
				{
					var currentPage = this.Pages[page];
					if (currentPage != null)
					{
						for (; minimumInnerIndex < this.PageSize; ++minimumInnerIndex)
						{
							var currentItem = currentPage[minimumInnerIndex];
							if (currentItem.HasValue)
							{
								index = page << this.PageBits;
								index = index | minimumInnerIndex;
								return true;
							}
						}
					}
					minimumInnerIndex = 0;
				}
				return false;
			}

			public bool PreviousItem(ref int index)
			{
				if (this.MinimumUsedIndex == -1 || this.MaximumUsedIndex == -1)
					return false;

				if (index-- > this.MaximumUsedIndex)
					index = this.MaximumUsedIndex;

				var minimumPage = this.MinimumUsedIndex >> this.PageBits;
				var minimumInnerIndex = this.MinimumUsedIndex & this.PageMask;
				var maximumPage = index >> this.PageBits;
				var maximumInnerIndex = index & this.PageMask;

				for (int page = maximumPage; page >= minimumPage; --page)
				{
					var currentPage = this.Pages[page];
					if (currentPage != null)
					{
						for (; maximumInnerIndex >= 0; --maximumInnerIndex)
						{
							var currentItem = currentPage[maximumInnerIndex];
							if (currentItem.HasValue)
							{
								index = page << this.PageBits;
								index = index | maximumInnerIndex;
								return true;
							}
						}
					}
					maximumInnerIndex = this.PageSize - 1;
				}
				return false;
			}
			#endregion

			#region Private Methods
			private void UpdateBounds()
			{
				this.MinimumUsedIndex = this.MaximumUsedIndex = -1;
				for (int i = this.MaximumIndex; i >= 0; --i)
				{
					if (this.GetItem(i).HasValue)
					{
						this.MaximumUsedIndex = i;
						break;
					}
				}
				for (int i = 0; i <= this.MaximumIndex; ++i)
				{
					if (this.GetItem(i).HasValue)
					{
						this.MinimumUsedIndex = i;
						break;
					}
				}
			}

			private void SetItem(int index, ValueHolder? item, bool doBoundsUpdate)
			{
				if (index < 0 || index > this.MaximumIndex)
					return;
				var page = index >> this.PageBits;
				var innerIndex = index & this.PageMask;
				var pageArray = this.Pages[page];
				if (null == pageArray)
					this.Pages[page] = pageArray = new ValueHolder?[this.PageSize];
				pageArray[innerIndex] = item;
				if (doBoundsUpdate)
					this.UpdateBounds();
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
			}
			#endregion

			#region Nested Classes
			public class Page
			{
				public int MinimumUsedIndex { get; set; }
				public int MaximumUsedIndex { get; set; }

				private ValueHolder?[] Values { get; }
			}
			#endregion

			#region Test Helpers
			/// <summary>
			/// This is included for testing purposes: DO NOT USE
			/// </summary>
			/// <returns></returns>
			public ValueHolder?[][] GetPages()
			{
				return this.Pages;
			}

			public void LoadPages(ValueHolder?[,] pageData)
			{
				for (int row = 0; row <= pageData.GetUpperBound(0); ++row)
				{
					for (int column = 0; column <= pageData.GetUpperBound(1); ++column)
					{
						var pageArray = this.Pages[row];
						if (null == pageArray)
							this.Pages[row] = pageArray = new ValueHolder?[this.PageSize];
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
		#endregion
	}
}
