using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace OfficeOpenXml
{
	[DebuggerDisplay("R {Row} C {Column}")]
	internal class CellStore2Key
	{
		public CellStore2Key(int row, int column)
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
			var other = obj as CellStore2Key;
			if (other == null)
				return false;
			if (this.Row != other.Row)
				return false;
			if (this.Column != other.Column)
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

		public class CellStore2KeyComparer : IComparer<CellStore2Key>
		{
			public int Compare(CellStore2Key x, CellStore2Key y)
			{
				if (x.Row != y.Row)
					return x.Row - y.Row;
				else
					return x.Column - y.Column;
			}
		}
		#endregion

		#region Properties
		LeftLeaningRedBlackTree<CellStore2Key, T> Tree { get; } = new LeftLeaningRedBlackTree<CellStore2Key, T>(new Comparison<CellStore2Key>(new CellStore2KeyComparer().Compare));
		#endregion

		#region CellStore Internal API
		public bool GetDimension(out int fromRow, out int fromCol, out int toRow, out int toCol)
		{
			if(this.Tree.Count == 0)
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

		public T GetValue(int row, int column)
		{
			return this.GetValue(new CellStore2Key(row, column));
		}

		internal T GetValue(CellStore2Key key)
		{
			return this.Tree.GetValueForKey(key);
		}

		public bool Exists(int row, int column)
		{
			T value;
			return this.Tree.TryGetValueForKey(new CellStore2Key(row, column), out value);
		}

		public bool Exists(int row, int column, out T value)
		{
			return this.Tree.TryGetValueForKey(new CellStore2Key(row, column), out value);
		}

		public void SetValue(int row, int column, T value)
		{
			this.SetValue(new CellStore2Key(row, column), value);
		}

		internal void SetValue(CellStore2Key key, T value)
		{
			this.Tree.Add(key, value);
		}

		public void Insert(int fromRow, int fromCol, int rows, int columns)
		{
			// Reversing is necesary so that the tree is guaranteed to remain in a valid state. 
			foreach(var key in this.Tree.GetKeys().Reverse())
			{
				if (key.Row >= fromRow)
					key.Row += rows;
				if (key.Column >= fromCol)
					key.Column += columns;
			}
			this.Tree.AssertInvariants();
		}

		public void Clear(int fromRow, int fromCol, int toRow, int toColumn)
		{
			foreach (var key in this.Tree.GetKeys().ToList())
			{
				if (key.Row >= fromRow && key.Row <= toRow && key.Column >= fromCol && key.Column <= toColumn)
					this.Tree.Remove(key);
			}
		}

		public void Delete(int fromRow, int fromCol, int rows, int columns)
		{
			this.Delete(fromRow, fromCol, rows, columns, true);
		}

		public void Delete(int fromRow, int fromCol, int rows, int columns, bool shift)
		{
			this.Clear(fromRow, fromCol, fromRow + rows - 1, fromCol + columns + 1);
			if(shift == true)
			{
				foreach (var key in this.Tree.GetKeys().ToArray())
				{
					var newKey = new CellStore2Key(key.Row, key.Column);
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

		// TODO: Optimize this.
		public bool NextCell(ref int row, ref int col)
		{
			if (this.Tree.Count == 0)
				return false;
			int maxRow = this.Tree.MaximumKey.Row;
			int maxColumn = this.Tree.GetKeys().Max(key => key.Column);
			T value;
			for (int c = col; c <= maxColumn; c++)
				for (int r = row; r <= maxRow; r++)
				{
					// Skip the first cell, since that's the current cell.
					if (r == row && c == col)
						continue;
					if(this.Tree.TryGetValueForKey(new CellStore2Key(r, c), out value))
					{
						row = r;
						col = c;
						return true;
					}
				}
			return false;

		}

		// TODO: Optimize this.
		public bool PrevCell(ref int row, ref int col)
		{
			if (this.Tree.Count == 0)
				return false;
			int minRow = this.Tree.MinimumKey.Row;
			int minColumn = this.Tree.GetKeys().Min(key => key.Column);
			T value;
			for (int c = col; c >= minColumn; c--)
				for (int r = row; r >= minRow; r--)
				{
					// Skip the first cell, since that's the current cell.
					if (r == row && c == col)
						continue;
					if (this.Tree.TryGetValueForKey(new CellStore2Key(r, c), out value))
					{
						row = r;
						col = c;
						return true;
					}
				}
			return false;
		}

		public void Dispose()
		{
			// Nothing to do here?
		}

		internal IEnumerable<CellStore2Key> GetKeyEnumerator()
		{
			return this.Tree.GetKeys();
		}
#endregion

	}

	internal class BalancedStoreEnumerator<T> : ICellStoreEnumerator<T>
	{
#region Properties
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

		public int Row
		{
			get
			{
				return this.CurrentKey?.Row ?? this.StartRow;
			}
		}

		public int Column
		{
			get
			{
				return this.CurrentKey?.Column ?? this.StartColumn;
			}
		}

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
					this.CellStore.SetValue(this.CurrentKey, value);
				}
			}
		}

		public T Current
		{
			get
			{
				return this.Value;
			}
		}

		private CellStore2Key CurrentKey { get; set; }

		private BalancedCellStore<T> CellStore { get; set; }
		private int StartRow { get; set; }
		private int StartColumn { get; set; }
		private int EndRow { get; set; }
		private int EndColumn { get; set; }

		private IEnumerator<CellStore2Key> ActualEnumerator { get; set; }
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
		public void Dispose()
		{
			// Nothing to see here.
		}

		public IEnumerator<T> GetEnumerator()
		{
			this.Reset();
			return this;
		}

		bool OutOfRange = false;
		public bool MoveNext()
		{
			try
			{
				if (this.OutOfRange)
					return false;
				bool result;
				do
				{
					result = this.ActualEnumerator.MoveNext();
					this.OutOfRange = this.AllFutureKeysWillBeOutOfRange(this.ActualEnumerator.Current);
				}
				while (result && !this.OutOfRange && !this.IsIncludedCell(this.ActualEnumerator.Current));
				if (result)
					this.CurrentKey = this.ActualEnumerator.Current;
				return result && !this.OutOfRange;
			}
			catch (Exception ex)
			{
				throw;
			}
		}

		private bool IsIncludedCell(CellStore2Key key)
		{
			return key != null && (key.Column >= this.StartColumn && key.Column <= this.EndColumn) && (key.Row >= this.StartRow && key.Row <= this.EndRow);
		}

		private bool AllFutureKeysWillBeOutOfRange(CellStore2Key key)
		{
			return key == null || key.Row > this.EndRow || (key.Row == this.EndRow && key.Column > this.EndColumn);
		}

		public void Reset()
		{
			this.OutOfRange = false;
			this.ActualEnumerator = this.CellStore.GetKeyEnumerator().GetEnumerator();
		}

		IEnumerator IEnumerable.GetEnumerator()
		{
			this.Reset();
			return this;
		}
#endregion
	}
}
