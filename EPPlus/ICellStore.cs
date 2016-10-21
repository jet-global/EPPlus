using System;
using System.Collections.Generic;

namespace OfficeOpenXml
{
	internal static class CellStoreDelegates<T>
	{
		internal delegate void SetRangeValueDelegate(List<T> list, int index, int row, int column, object value);

		internal delegate void SetValueDelegate(List<T> list, int index, object value);
	}

	internal interface ICellStore<T> : IDisposable
	{
		T GetValue(int row, int column);
		void SetValue(int row, int column, T value);
		bool Exists(int row, int column, out T value);
		bool NextCell(ref int row, ref int column);
		void Delete(int _fromRow, int _fromCol, int v1, int v2);
		bool Exists(int row, int col);
		bool PrevCell(ref int row, ref int col);
		// SetRangeValueSpecial should be refactored to eliminate or vastly simplify / generalize the delegate
		void SetRangeValueSpecial(int _fromRow, int _fromCol, int v1, int v2, CellStoreDelegates<T>.SetRangeValueDelegate setValueDelegate, object values);
		// SetValueSpecial should be refactored to eliminate or vastly simplify / generalize the delegate
		void SetValueSpecial(int row, int col, CellStoreDelegates<T>.SetValueDelegate setStyleInnerUpdate, object value);
		void Clear(int _fromRow, int _fromCol, int toRow, int toCol);
		void Delete(int fromRow, int fromCol, int rows, int cols, bool shift);
		bool GetDimension(out int fromRow, out int fromCol, out int toRow, out int toCol);
		void Insert(int rowFrom, int v1, int rows, int v2);
	}

	internal interface ICellStoreEnumerator<T> : IEnumerable<T>, IEnumerator<T>
	{
		string CellAddress { get; }
		int Column { get; }
		int Row { get; }
		T Value { get; set; }
	}

	internal class CellStoreEnumeratorFactory<T>
	{
		public static ICellStoreEnumerator<T> GetNewEnumerator(ICellStore<T> cellStore)
		{
			var specificCellStore = cellStore as CellStore<T>;
			if(specificCellStore != null)
				return new CellStore<T>.CellsStoreEnumerator<T>(specificCellStore);
			throw new NotImplementedException($"No CellStoreEnumerator accepts the type {cellStore.GetType()}.");
		}

		public static ICellStoreEnumerator<T> GetNewEnumerator(ICellStore<T> cells, int startRow, int startColumn, int endRow, int endColumn)
		{
			var specificCellStore = cells as CellStore<T>;
			if (specificCellStore != null)
				return new CellStore<T>.CellsStoreEnumerator<T>(specificCellStore, startRow, startColumn, endRow, endColumn);
			throw new NotImplementedException($"No CellStoreEnumerator accepts the type {cells.GetType()}.");
		}
	}
}