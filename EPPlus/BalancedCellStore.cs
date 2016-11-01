using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace OfficeOpenXml
{
	[DebuggerDisplay("R {Row} C {Column}")]
	internal class BalancedCellStoreKey
	{
		public BalancedCellStoreKey(int row, int column)
		{
			this.Row = row;
			this.Column = column;
		}
		public int Row { get; set; }

		public int Column { get; set; }

		public override bool Equals(object obj)
		{
			if (obj == null)
				return false;
			var other = obj as BalancedCellStoreKey;
			if (other == null)
				return false;
			else if (this.Row != other.Row)
				return false;
			else if (this.Column != other.Column)
				return false;
			return true;
		}

		public override int GetHashCode()
		{
			return this.Column + this.Row;
		}

		public override string ToString()
		{
			return $"R {this.Row} C {this.Column}";
		}
	}

	internal class BalancedCellStore<T> : ICellStore<T>, IDisposable
	{
		#region Nested Classes

		public class CellStore2KeyComparer : IComparer<BalancedCellStoreKey>
		{
			public int Compare(BalancedCellStoreKey x, BalancedCellStoreKey y)
			{
				if (x.Row != y.Row)
					return x.Row - y.Row;
				else
					return x.Column - y.Column;
			}
		}
		#endregion

		#region Properties
		LeftLeaningRedBlackTree<BalancedCellStoreKey, T> Tree { get; } = new LeftLeaningRedBlackTree<BalancedCellStoreKey, T>(new Comparison<BalancedCellStoreKey>(new CellStore2KeyComparer().Compare));
		#endregion

		#region CellStore Internal API
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
			if (this.Tree.Count == 0)
			{
				fromRow = fromCol = toRow = toCol = 0;
				return false;
			}
			fromRow = Math.Max(this.Tree.MinimumKey.Row, 1);
			fromCol = this.Tree.GetKeys().Min(key => key.Column);
			fromCol = Math.Max(fromCol, 1);
			toRow = Math.Min(Math.Max(this.Tree.MaximumKey.Row, 1), ExcelPackage.MaxRows);
			toCol = this.Tree.GetKeys().Max(key => key.Column);
			toCol = Math.Min(Math.Max(toCol, 1), ExcelPackage.MaxColumns);
			return true;
		}

		/// <summary>
		/// Get the value at a particular location.
		/// </summary>
		/// <param name="row">The row to read a value from.</param>
		/// <param name="column">The column to read a value from.</param>
		/// <returns>The value at the specified row and column.</returns>
		public T GetValue(int row, int column)
		{
			return this.GetValue(new BalancedCellStoreKey(row, column));
		}

		/// <summary>
		/// Determine if a value exists at the given location.
		/// </summary>
		/// <param name="row">The row to look for a value at.</param>
		/// <param name="column">The column to look for a value at.</param>
		/// <returns>True if a value was found at the location; false otherwise.</returns>
		public bool Exists(int row, int column)
		{
			T value;
			return this.Tree.TryGetValueForKey(new BalancedCellStoreKey(row, column), out value);
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
			return this.Tree.TryGetValueForKey(new BalancedCellStoreKey(row, column), out value);
		}

		/// <summary>
		/// Set the value at a particular location.
		/// </summary>
		/// <param name="row">The row of the location to set a value at.</param>
		/// <param name="column">The column of the location to set a value at.</param>
		/// <param name="value">The value to store at the location.</param>
		public void SetValue(int row, int column, T value)
		{
			this.SetValue(new BalancedCellStoreKey(row, column), value);
		}

		/// <summary>
		/// "Insert space" into the cellStore by updating all keys beyond the specified row or column by the specified number of rows or columns.
		/// </summary>
		/// <param name="rowFrom">The row to start updating keys from.</param>
		/// <param name="columnFrom">The columnn to start updating keys from.</param>
		/// <param name="rows">The number of rows being inserted.</param>
		/// <param name="columns">The number of columns being inserted.</param>
		public void Insert(int rowFrom, int columnFrom, int rows, int columns)
		{
			// Reversing is necesary so that the tree is guaranteed to remain in a valid state. 
			foreach (var key in this.Tree.GetKeys().Reverse())
			{
				if (key.Row >= rowFrom)
					key.Row += rows;
				if (key.Column >= columnFrom)
					key.Column += columns;
			}
#if DEBUGGING
			// Verify that we've left the tree in a valid state.
			this.Tree.AssertInvariants();
#endif
		}

		/// <summary>
		/// Removes the values in the specified range without updating cells below or to the right of the specified range.
		/// </summary>
		/// <param name="fromRow">The first row whose cells should be cleared.</param>
		/// <param name="fromCol">The first column whose cells should be cleared.</param>
		/// <param name="toRow">The last row whose values should be cleared.</param>
		/// <param name="toColumn">The last column whose values should be cleared.</param>
		public void Clear(int fromRow, int fromCol, int toRow, int toColumn)
		{
			foreach (var key in this.Tree.GetKeys().ToList())
			{
				if (key.Row >= fromRow && key.Row <= toRow && key.Column >= fromCol && key.Column <= toColumn)
					this.Tree.Remove(key);
			}
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
			this.Delete(fromRow, fromCol, rows, columns, true);
		}

		/// <summary>
		/// Deletes rows and/or columns from the workbook. This deletes all existing nodes in the specified range.
		/// If <paramref name="shift"/> is set to true, all subsequent keys will be updated to reflect their new position after being shifted into the newly-vacated space.
		/// </summary>
		/// <param name="fromRow">The first row to delete.</param>
		/// <param name="fromCol">The first column to delete.</param>
		/// <param name="rows">The number of rows to delete.</param>
		/// <param name="columns">The number of columns to delete.</param>
		/// <param name="shift">Whether or not subsequent rows and columns should be shifted up / left to fill the space that was deleted, or left where they are.</param>
		public void Delete(int fromRow, int fromCol, int rows, int columns, bool shift)
		{
			this.Clear(fromRow, fromCol, fromRow + rows - 1, fromCol + columns + 1);
			if (shift == true)
			{
				foreach (var key in this.Tree.GetKeys().ToArray())
				{
					var newKey = new BalancedCellStoreKey(key.Row, key.Column);
					bool update = false;
					if (fromRow > 0 && newKey.Row >= fromRow)
					{
						newKey.Row -= rows;
						update = true;
					}
					if (fromCol > 0 && newKey.Column >= fromCol)
					{
						newKey.Column -= columns;
					}
					if (update)
					{
						var temp = this.Tree.GetValueForKey(key);
						this.Tree.Remove(key);
						this.Tree.Add(newKey, temp);
					}
				}
			}
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
			if (this.Tree.Count == 0)
				return false;
			int maxRow = this.Tree.MaximumKey.Row;
			int maxColumn = this.Tree.GetKeys().Max(key => key.Column);
			T value;
			for (int c = column; c <= maxColumn; c++)
				for (int r = row; r <= maxRow; r++)
				{
					// Skip the first cell, since that's the current cell.
					if (r == row && c == column)
						continue;
					if (this.Tree.TryGetValueForKey(new BalancedCellStoreKey(r, c), out value))
					{
						row = r;
						column = c;
						return true;
					}
				}
			return false;

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
			if (this.Tree.Count == 0)
				return false;
			int minRow = this.Tree.MinimumKey.Row;
			int minColumn = this.Tree.GetKeys().Min(key => key.Column);
			T value;
			for (int c = column; c >= minColumn; c--)
				for (int r = row; r >= minRow; r--)
				{
					// Skip the first cell, since that's the current cell.
					if (r == row && c == column)
						continue;
					if (this.Tree.TryGetValueForKey(new BalancedCellStoreKey(r, c), out value))
					{
						row = r;
						column = c;
						return true;
					}
				}
			return false;
		}

		/// <summary>
		/// Dispose of any objects allocated by this CellStore.
		/// </summary>
		public void Dispose()
		{
			// Nothing to do here.
		}
		#endregion

		#region Internal Methods to support the BalancedCellStoreEnumerator
		/// <summary>
		/// Get the collection of keys in this CellStore.
		/// </summary>
		/// <returns>An enumerable collection of the keys in the tree.</returns>
		internal IEnumerable<BalancedCellStoreKey> GetKeys()
		{
			return this.Tree.GetKeys();
		}

		/// <summary>
		/// Set an existing value without rebalancing the red-black tree.
		/// </summary>
		/// <param name="key">The jey of the existing value to update.</param>
		/// <param name="value">The new value.</param>
		internal void SetExistingValue(BalancedCellStoreKey key, T value)
		{
			this.Tree.Add(key, value, false);
		}

		/// <summary>
		/// Get the value stored at a particular <see cref="BalancedCellStoreKey"/>.
		/// </summary>
		/// <param name="key">The location of the value to read.</param>
		/// <returns>the value stored at the location, or the default value of the type (typically null or zero) if the value does not exist.</returns>
		internal T GetValue(BalancedCellStoreKey key)
		{
			return this.Tree.GetValueForKey(key);
		}

		/// <summary>
		/// Set the value at a particular location.
		/// </summary>
		/// <param name="key">The location to store the value at.</param>
		/// <param name="value">The value to store.</param>
		internal void SetValue(BalancedCellStoreKey key, T value)
		{
			this.Tree.Add(key, value);
		}
		#endregion

	}

	/// <summary>
	/// An enumerator for the <see cref="BalancedCellStore{T}"/>.
	/// </summary>
	/// <typeparam name="T">The type of values stored in the <see cref="BalancedCellStore{T}"/>.</typeparam>
	internal class BalancedStoreEnumerator<T> : ICellStoreEnumerator<T>
	{
		#region Public Properties
		/// <summary>
		/// The string address, such as A1, of the current cell.
		/// </summary>
		public string CellAddress
		{
			get
			{
				return ExcelAddressBase.GetAddress(this.Row, this.Column);
			}
		}

		object IEnumerator.Current
		{
			get
			{
				this.Reset();
				return this;
			}
		}

		/// <summary>
		/// The row of the current cell.
		/// </summary>
		public int Row
		{
			get
			{
				return this.CurrentKey?.Row ?? this.StartRow;
			}
		}

		/// <summary>
		/// The column number of the current cell.
		/// </summary>
		public int Column
		{
			get
			{
				return this.CurrentKey?.Column ?? this.StartColumn;
			}
		}

		/// <summary>
		/// The value stored at the current cell.
		/// </summary>
		public T Value
		{
			get
			{
				lock (this.CellStore)
				{
					if (this.CurrentKey == null)
						return default(T);
					return this.CellStore.GetValue(this.CurrentKey);
				}
			}
			set
			{
				lock (this.CellStore)
				{
					this.CellStore.SetExistingValue(this.CurrentKey, value);
				}
			}
		}

		/// <summary>
		/// The value stored at the current cell.
		/// </summary>
		public T Current
		{
			get
			{
				return this.Value;
			}
		}
		#endregion

		#region Private Properties
		private BalancedCellStoreKey CurrentKey { get; set; }

		private BalancedCellStore<T> CellStore { get; set; }

		private int StartRow { get; set; }

		private int StartColumn { get; set; }

		private int EndRow { get; set; }

		private int EndColumn { get; set; }

		private bool NoMoreValidKeys { get; set; }

		private IEnumerator<BalancedCellStoreKey> ActualKeyEnumerator { get; set; }
		#endregion

		#region Constructors
		public BalancedStoreEnumerator(BalancedCellStore<T> cellStore) : this(cellStore, 1, 1, ExcelPackage.MaxRows, ExcelPackage.MaxColumns) { }
		public BalancedStoreEnumerator(BalancedCellStore<T> cellStore, int startRow, int startCol, int endRow, int endCol)
		{
			this.CellStore = cellStore;

			this.StartRow = startRow;
			this.StartColumn = startCol;
			this.EndRow = endRow;
			this.EndColumn = endCol;

			this.Reset();
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Dispose of any objects allocated by this enumerator.
		/// </summary>
		public void Dispose()
		{
			// Nothing to see here.
		}

		/// <summary>
		/// Get this <see cref="IEnumerator{T}"/>.
		/// </summary>
		/// <returns>This.</returns>
		public IEnumerator<T> GetEnumerator()
		{
			this.Reset();
			return this;
		}

		/// <summary>
		/// Move to the next the next <see cref="BalancedCellStoreKey"/>.
		/// </summary>
		/// <returns>True if another element exists; false otherwise.</returns>
		public bool MoveNext()
		{
			try
			{
				if (this.NoMoreValidKeys)
					return false;
				bool result;
				do
				{
					result = this.ActualKeyEnumerator.MoveNext();
					this.NoMoreValidKeys = this.AllFutureKeysWillBeOutOfRange(this.ActualKeyEnumerator.Current);
				}
				while (result && !this.NoMoreValidKeys && !this.IsIncludedCell(this.ActualKeyEnumerator.Current));
				if (result)
					this.CurrentKey = this.ActualKeyEnumerator.Current;
				return result && !this.NoMoreValidKeys;
			}
			catch (Exception)
			{
				// Put a breakpoint here to debug exceptions that occur while enumerating.
				throw;
			}
		}

		/// <summary>
		/// Restart this enumerator.
		/// </summary>
		public void Reset()
		{
			this.NoMoreValidKeys = false;
			this.ActualKeyEnumerator = this.CellStore.GetKeys().GetEnumerator();
		}
		#endregion

		#region Private Methods
		private bool IsIncludedCell(BalancedCellStoreKey key)
		{
			return key != null && (key.Column >= this.StartColumn && key.Column <= this.EndColumn) && (key.Row >= this.StartRow && key.Row <= this.EndRow);
		}

		private bool AllFutureKeysWillBeOutOfRange(BalancedCellStoreKey key)
		{
			return key == null || key.Row > this.EndRow || (key.Row == this.EndRow && key.Column > this.EndColumn);
		}

		IEnumerator IEnumerable.GetEnumerator()
		{
			this.Reset();
			return this;
		}
		#endregion
	}
}
