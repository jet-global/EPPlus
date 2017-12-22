using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeOpenXml
{
	internal class ZCellStore<T> : ICellStore<T>
	{
		private class Sheet
		{
			// contains 128 pages of columns
			public ColumnsPage[] Pages { get; } = new ColumnsPage[ZCellStore<T>.ColumnPageSize];
		}

		internal class ColumnsPage
		{
			// contains 128 columns from a page 
			public Column[] Columns { get; } = new Column[ZCellStore<T>.ColumnPageSize];
		}

		internal class Column
		{
			// contains 1024 pages of rows
			public RowsPage[] Pages { get; } = new RowsPage[ZCellStore<T>.RowPageSize];
		}

		internal class RowsPage
		{
			// Contains 1024 rows from a page in the column
			public DataWrapper[] Data { get; } = new DataWrapper[ZCellStore<T>.RowPageSize];
		}

		internal class DataWrapper // TODO -- Consider making this a struct
		{
			public T Item { get; set; }
		}

		#region Constants
		private const int RowPageSize = 1024; // The number of rows in a row page
		private const int ColumnPageSize = 128; // The number of columns in a column page
		private const int RowMask = 1023; // 111111111 ( 9 1's )
		private const int ColumnMask = 127; // 111111 ( 6 1's)
		private const int RowShift = 10; // 2 ^ 10 = 1024 so right-shifting by 10 bits divides by 1024
		private const int ColumnShift = 7; // 2 ^ 7 = 128 so right-shifting by 7 bits divides by 128
		#endregion

		private Sheet SheetInternal { get; } = new Sheet();

		private int MaxUsedColumn = 0;
		private int MaxUsedRow = 0;

		#region ICellStore<T> Methods
		public T GetValue(int row, int column)
		{
			this.Exists(row, column, out T value);
			return value;
		}

		public void SetValue(int row, int column, T value)
		{
			if (ZCellStore<T>.GetRowCoordinates(row, out int rowPage, out int rowPageIndex) &&
					ZCellStore<T>.GetColumnCoordinates(column, out int columnPage, out int columnPageIndex))
			{
				var columnObj = this.GetColumn(columnPage, columnPageIndex, true);
				var cell = this.GetRow(columnObj, rowPage, rowPageIndex, true);
				cell.Item = value;
			}
		}

		public bool Exists(int row, int column, out T value)
		{
			value = default(T);
			if (ZCellStore<T>.GetRowCoordinates(row, out int rowPage, out int rowPageIndex) &&
					ZCellStore<T>.GetColumnCoordinates(column, out int columnPage, out int columnPageIndex))
			{
				var columnObj = this.GetColumn(columnPage, columnPageIndex, false);
				if (null == columnObj)
					return false;
				var cell = this.GetRow(columnObj, rowPage, rowPageIndex, false);
				if (null == cell)
					return false;
				value = cell.Item;
				return true;
			}
			return false;
		}

		public bool NextCell(ref int row, ref int column)
		{
			while (row <= this.MaxUsedRow)
			{
				while (++column <= this.MaxUsedColumn)
				{
					if (ZCellStore<T>.GetColumnCoordinates(column, out int columnPage, out int columnPageIndex))
					{
						var columnObj = this.GetColumn(columnPage, columnPageIndex, false);
						if (null != columnObj)
						{
							if (ZCellStore<T>.GetRowCoordinates(row, out int rowPage, out int rowPageIndex))
							{
								var cell = this.GetRow(columnObj, rowPage, rowPageIndex, false);
								if (cell != null)
									return true;
							}
						}
					}
				}
				column = -1;
				++row;
			}
			return false;
		}

		public bool PrevCell(ref int row, ref int column)
		{
			throw new NotImplementedException();
		}

		public void Delete(int fromRow, int fromCol, int rows, int columns)
		{
			throw new NotImplementedException();
		}

		public void Delete(int fromRow, int fromCol, int rows, int columns, bool shift)
		{
			throw new NotImplementedException();
		}

		public bool Exists(int row, int column)
		{
			return this.Exists(row, column, out T value);
		}

		public void Clear(int _fromRow, int _fromCol, int toRow, int toCol)
		{
			throw new NotImplementedException();
		}

		public bool GetDimension(out int fromRow, out int fromCol, out int toRow, out int toCol)
		{
			throw new NotImplementedException();
		}

		public void Insert(int rowFrom, int columnFrom, int rows, int columns)
		{
			throw new NotImplementedException();
		}

		public void Dispose()
		{
			throw new NotImplementedException();
		}
		#endregion

		#region Private Methods

		private Column GetColumn(int columnPage, int columnPageIndex, bool create)
		{
			var page = this.SheetInternal.Pages[columnPage];
			if (null == page)
			{
				if (create)
					this.SheetInternal.Pages[columnPage] = page = new ColumnsPage();
				else
					return null;
			}
			var columnObj = page.Columns[columnPageIndex];
			if (null == columnObj && create)
			{
				page.Columns[columnPageIndex] = columnObj = new Column();
				var column = ZCellStore<T>.RebuildColumn(columnPage, columnPageIndex);
				if (column > this.MaxUsedColumn)
					this.MaxUsedColumn = column;
			}
			return columnObj;
		}

		private DataWrapper GetRow(Column column, int rowPage, int rowPageIndex, bool create)
		{
			var page = column.Pages[rowPage];
			if (null == page)
			{
				if (create)
					column.Pages[rowPage] = page = new RowsPage();
				else
					return null;
			}
			var cell = page.Data[rowPageIndex];
			if (null == cell && create)
			{
				page.Data[rowPageIndex] = cell = new DataWrapper();
				var row = ZCellStore<T>.RebuildRow(rowPage, rowPageIndex);
				if (row > this.MaxUsedRow)
					this.MaxUsedRow = row;
			}
			return cell;
		}
		#endregion

#if DEBUG
		public Tuple<int, int, T>[] GetDump()
		{
			return null;
		}

		public void Initialize(Tuple<int, int, T>[] data)
		{
			foreach (var datum in data)
			{
				var rowValue = datum.Item1;
				var columnValue = datum.Item2;
				var dataValue = datum.Item3;
				ZCellStore<T>.GetColumnCoordinates(columnValue, out int columnPage, out int columnPageIndex);
				var column = this.GetColumn(columnPage, columnPageIndex, true);
				ZCellStore<T>.GetRowCoordinates(rowValue, out int rowPage, out int rowPageIndex);
				var cell = this.GetRow(column, rowPage, rowPageIndex, true);
				cell.Item = dataValue;
			}
		}
#endif

		#region Internal Static Methods
		internal static bool GetRowCoordinates(int row, out int rowPage, out int pageIndex)
		{
			rowPage = pageIndex = 0;
			if (row < 1 || row > ExcelPackage.MaxRows)
				return false;
			// Rows are 1-index in Excel so let's shift down by one to get into our array-space
			--row;
			// Since row pages are in an array of 1024 of we can identify the page by right-shifting 10 (dividing by 1024)
			rowPage = row >> ZCellStore<T>.RowShift;
			// Since each page contains 1024 rows, we can take the least significant bits and bitwise AND them to get the index into the actual page.
			pageIndex = row & ZCellStore<T>.RowMask;
			return true;
		}

		internal static bool GetColumnCoordinates(int column, out int columnPage, out int pageIndex)
		{
			columnPage = pageIndex = 0;
			if (column < 1 || column > ExcelPackage.MaxColumns)
				return false;
			// Columns are 1-index in Excel so let's shift down by one to get into our array-space
			--column;
			// Since column pages are in an array of 128 of we can identify the page by right-shifting 7 (dividing by 128)
			columnPage = column >> ZCellStore<T>.ColumnShift;
			// Since each page contains 128 columns, we can take the least significant bits and bitwise AND them to get the index into the actual page.
			pageIndex = column & ZCellStore<T>.ColumnMask;
			return true;
		}

		internal static int RebuildColumn(int columnPage, int pageIndex)
		{
			int column = columnPage << ZCellStore<T>.ColumnShift;
			column = column | ZCellStore<T>.ColumnMask;
			return column;
		}

		internal static int RebuildRow(int rowPage, int pageIndex)
		{
			int row = rowPage << ZCellStore<T>.RowShift;
			row = row | ZCellStore<T>.RowMask;
			return row;
		}
		#endregion
	}
}
