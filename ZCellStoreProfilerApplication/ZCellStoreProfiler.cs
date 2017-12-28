using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using OfficeOpenXml;

namespace ZCellStoreProfilerApplication
{
	internal class ZCellStoreProfiler<T> : ICellStore<T>
	{
		#region Properties
		private ZCellStore<T> ZCellStore { get; } = new ZCellStore<T>();
		private CellStore<T> CellStore { get; } = new CellStore<T>();

		public Stopwatch ZCellStoreTimer { get; } = new Stopwatch();
		public Stopwatch CellStoreTimer { get; } = new Stopwatch();

		private StringBuilder Log { get; } = new StringBuilder();
		#endregion

		#region Public Methods
		public void EnumerateItems()
		{
			bool zcellstoreExcepted = false;
			bool cellstoreExcepted = false;
			int startRow = 0, startColumn = 0;
			this.ZCellStoreTimer.Start();
			try
			{
				while (this.ZCellStore.NextCell(ref startRow, ref startColumn)) ;
			}
			catch { zcellstoreExcepted = true; }
			this.ZCellStoreTimer.Stop();
			startRow = 0;
			startColumn = 0;
			this.CellStoreTimer.Start();
			try
			{
				while (this.CellStore.NextCell(ref startRow, ref startColumn)) ;
			}
			catch { cellstoreExcepted = true; }
			this.CellStoreTimer.Stop();
			this.Log.AppendLine($"EnumerateItems,{(zcellstoreExcepted ? "exception" : this.ZCellStoreTimer.ElapsedTicks.ToString())},{(cellstoreExcepted ? "exception" : this.CellStoreTimer.ElapsedTicks.ToString())}");
		}

		public void SaveResults(string path)
		{
			File.WriteAllText(path, this.Log.ToString());
		}
		#endregion

		#region ICellStore<T> Members
		public int MaximumRow => this.ZCellStore.MaximumRow;
		public int MaximumColumn => this.ZCellStore.MaximumColumn;

		public void Clear(int _fromRow, int _fromCol, int toRow, int toCol)
		{
			bool zcellstoreExcepted = false;
			bool cellstoreExcepted = false;
			this.ZCellStoreTimer.Start();
			try
			{
				this.ZCellStore.Clear(_fromRow, _fromCol, toRow, toCol);
			}
			catch { zcellstoreExcepted = true; }
			this.ZCellStoreTimer.Stop();
			this.CellStoreTimer.Start();
			try
			{
				this.CellStore.Clear(_fromRow, _fromCol, toRow, toCol);
			}
			catch { cellstoreExcepted = true; }
			this.CellStoreTimer.Stop();
			this.Log.AppendLine($"Clear,{(zcellstoreExcepted ? "exception" : this.ZCellStoreTimer.ElapsedTicks.ToString())},{(cellstoreExcepted ? "exception" : this.CellStoreTimer.ElapsedTicks.ToString())}");
		}

		public void Delete(int fromRow, int fromCol, int rows, int columns)
		{
			bool zcellstoreExcepted = false;
			bool cellstoreExcepted = false;
			this.ZCellStoreTimer.Start();
			try
			{
				this.ZCellStore.Delete(fromRow, fromCol, rows, columns);
			}
			catch { zcellstoreExcepted = true; }
			this.ZCellStoreTimer.Stop();
			this.CellStoreTimer.Start();
			try
			{
				this.CellStore.Delete(fromRow, fromCol, rows, columns);
			}
			catch { cellstoreExcepted = true; }
			this.CellStoreTimer.Stop();
			this.Log.AppendLine($"Delete,{(zcellstoreExcepted ? "exception" : this.ZCellStoreTimer.ElapsedTicks.ToString())},{(cellstoreExcepted ? "exception" : this.CellStoreTimer.ElapsedTicks.ToString())}");
		}

		//public void Delete(int fromRow, int fromCol, int rows, int columns, bool shift)
		//{
		//	bool zcellstoreExcepted = false;
		//	bool cellstoreExcepted = false;
		//	this.ZCellStoreTimer.Start();
		//	try
		//	{
		//		this.ZCellStore.Delete(fromRow, fromCol, rows, columns, shift);
		//	}
		//	catch { zcellstoreExcepted = true; }
		//	this.ZCellStoreTimer.Stop();
		//	this.CellStoreTimer.Start();
		//	try
		//	{
		//		this.CellStore.Delete(fromRow, fromCol, rows, columns, shift);
		//	}
		//	catch { cellstoreExcepted = true; }
		//	this.CellStoreTimer.Stop();
		//	this.Log.AppendLine($"Delete_{shift},{(zcellstoreExcepted ? "exception" : this.ZCellStoreTimer.ElapsedTicks.ToString())},{(cellstoreExcepted ? "exception" : this.CellStoreTimer.ElapsedTicks.ToString())}");
		//}

		public void Dispose()
		{
			throw new NotImplementedException();
		}

		public bool Exists(int row, int column, out T value)
		{
			value = default(T);
			bool zcellstoreExcepted = false;
			bool cellstoreExcepted = false;
			this.ZCellStoreTimer.Start();
			try
			{
				this.ZCellStore.Exists(row, column, out value);
			}
			catch { zcellstoreExcepted = true; }
			this.ZCellStoreTimer.Stop();
			this.CellStoreTimer.Start();
			try
			{
				this.CellStore.Exists(row, column, out value);
			}
			catch { cellstoreExcepted = true; }
			this.CellStoreTimer.Stop();
			this.Log.AppendLine($"Exists_Value,{(zcellstoreExcepted ? "exception" : this.ZCellStoreTimer.ElapsedTicks.ToString())},{(cellstoreExcepted ? "exception" : this.CellStoreTimer.ElapsedTicks.ToString())}");
			return true;
		}

		public bool Exists(int row, int column)
		{
			bool zcellstoreExcepted = false;
			bool cellstoreExcepted = false;
			this.ZCellStoreTimer.Start();
			try
			{
				this.ZCellStore.Exists(row, column);
			}
			catch { zcellstoreExcepted = true; }
			this.ZCellStoreTimer.Stop();
			this.CellStoreTimer.Start();
			try
			{
				this.CellStore.Exists(row, column);
			}
			catch { cellstoreExcepted = true; }
			this.CellStoreTimer.Stop();
			this.Log.AppendLine($"Exists,{(zcellstoreExcepted ? "exception" : this.ZCellStoreTimer.ElapsedTicks.ToString())},{(cellstoreExcepted ? "exception" : this.CellStoreTimer.ElapsedTicks.ToString())}");
			return true;
		}

		public bool GetDimension(out int fromRow, out int fromCol, out int toRow, out int toCol)
		{
			fromRow = fromCol = toRow = toCol = 0;
			bool zcellstoreExcepted = false;
			bool cellstoreExcepted = false;
			this.ZCellStoreTimer.Start();
			try
			{
				this.ZCellStore.GetDimension(out fromRow, out fromCol, out toRow, out toCol);
			}
			catch { zcellstoreExcepted = true; }
			this.ZCellStoreTimer.Stop();
			this.CellStoreTimer.Start();
			try
			{
				this.CellStore.GetDimension(out fromRow, out fromCol, out toRow, out toCol);
			}
			catch { cellstoreExcepted = true; }
			this.CellStoreTimer.Stop();
			this.Log.AppendLine($"GetDimension,{(zcellstoreExcepted ? "exception" : this.ZCellStoreTimer.ElapsedTicks.ToString())},{(cellstoreExcepted ? "exception" : this.CellStoreTimer.ElapsedTicks.ToString())}");
			return false;
		}

		public T GetValue(int row, int column)
		{
			bool zcellstoreExcepted = false;
			bool cellstoreExcepted = false;
			this.ZCellStoreTimer.Start();
			try
			{
				this.ZCellStore.GetValue(row, column);
			}
			catch { zcellstoreExcepted = true; }
			this.ZCellStoreTimer.Stop();
			this.CellStoreTimer.Start();
			try
			{
				this.CellStore.GetValue(row, column);
			}
			catch { cellstoreExcepted = true; }
			this.CellStoreTimer.Stop();
			this.Log.AppendLine($"GetValue,{(zcellstoreExcepted ? "exception" : this.ZCellStoreTimer.ElapsedTicks.ToString())},{(cellstoreExcepted ? "exception" : this.CellStoreTimer.ElapsedTicks.ToString())}");
			return default(T);
		}

		public void Insert(int rowFrom, int columnFrom, int rows, int columns)
		{
			bool zcellstoreExcepted = false;
			bool cellstoreExcepted = false;
			this.ZCellStoreTimer.Start();
			try
			{
				this.ZCellStore.Insert(rowFrom, columnFrom, rows, columns);
			}
			catch { zcellstoreExcepted = true; }
			this.ZCellStoreTimer.Stop();
			this.CellStoreTimer.Start();
			try
			{
				this.CellStore.Insert(rowFrom, columnFrom, rows, columns);
			}
			catch { cellstoreExcepted = true; }
			this.CellStoreTimer.Stop();
			this.Log.AppendLine($"Insert,{(zcellstoreExcepted ? "exception" : this.ZCellStoreTimer.ElapsedTicks.ToString())},{(cellstoreExcepted ? "exception" : this.CellStoreTimer.ElapsedTicks.ToString())}");
		}

		public bool NextCell(ref int row, ref int column)
		{
			bool zcellstoreExcepted = false;
			bool cellstoreExcepted = false;
			int localRow = row, localColumn = column;
			this.ZCellStoreTimer.Start();
			try
			{
				this.ZCellStore.NextCell(ref localRow, ref localColumn);
			}
			catch { zcellstoreExcepted = true; }
			this.ZCellStoreTimer.Stop();
			this.CellStoreTimer.Start();
			try
			{
				this.CellStore.NextCell(ref row, ref column);
			}
			catch { cellstoreExcepted = true; }
			this.CellStoreTimer.Stop();
			this.Log.AppendLine($"NextCell,{(zcellstoreExcepted ? "exception" : this.ZCellStoreTimer.ElapsedTicks.ToString())},{(cellstoreExcepted ? "exception" : this.CellStoreTimer.ElapsedTicks.ToString())}");
			return false;
		}

		public bool PrevCell(ref int row, ref int column)
		{
			bool zcellstoreExcepted = false;
			bool cellstoreExcepted = false;
			int localRow = row, localColumn = column;
			this.ZCellStoreTimer.Start();
			try
			{
				this.ZCellStore.PrevCell(ref localRow, ref localColumn);
			}
			catch { zcellstoreExcepted = true; }
			this.ZCellStoreTimer.Stop();
			this.CellStoreTimer.Start();
			try
			{
				this.CellStore.PrevCell(ref row, ref column);
			}
			catch { cellstoreExcepted = true; }
			this.CellStoreTimer.Stop();
			this.Log.AppendLine($"PrevCell,{(zcellstoreExcepted ? "exception" : this.ZCellStoreTimer.ElapsedTicks.ToString())},{(cellstoreExcepted ? "exception" : this.CellStoreTimer.ElapsedTicks.ToString())}");
			return false;
		}

		public void SetValue(int row, int column, T value)
		{
			bool zcellstoreExcepted = false;
			bool cellstoreExcepted = false;
			this.ZCellStoreTimer.Start();
			try
			{
				this.ZCellStore.SetValue(row, column, value);
			}
			catch { zcellstoreExcepted = true; }
			this.ZCellStoreTimer.Stop();
			this.CellStoreTimer.Start();
			try
			{
				this.CellStore.SetValue(row, column, value);
			}
			catch { cellstoreExcepted = true; }
			this.CellStoreTimer.Stop();
			this.Log.AppendLine($"SetValue,{(zcellstoreExcepted ? "exception" : this.ZCellStoreTimer.ElapsedTicks.ToString())},{(cellstoreExcepted ? "exception" : this.CellStoreTimer.ElapsedTicks.ToString())}");
		}

		public ICellStoreEnumerator<T> GetEnumerator()
		{
			throw new NotImplementedException();
		}

		public ICellStoreEnumerator<T> GetEnumerator(int startRow, int startColumn, int endRow, int endColumn)
		{
			throw new NotImplementedException();
		}
		#endregion
	}
}
