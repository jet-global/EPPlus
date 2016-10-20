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
		bool GetNextCell(ref int row, ref int colPos, int minColPos, int maxRow, int maxColPos);
		int GetPosition(int _startCol);
		int GetColumnPosIndex(int colPos);
		T GetValue(int row, int column);
		void SetValue(int row, int column, T value);
		bool GetPrevCell(ref int row, ref int colPos, int minRow, int minColPos, int maxColPos);
		bool Exists(int row, int column, out T value);
		bool NextCell(ref int row, ref int column);
		void Delete(int _fromRow, int _fromCol, int v1, int v2);
		bool Exists(int row, int col);
		bool PrevCell(ref int row, ref int col);
		void SetRangeValueSpecial(int _fromRow, int _fromCol, int v1, int v2, CellStoreDelegates<T>.SetRangeValueDelegate setValueDelegate, object values);
		void SetValueSpecial(int row, int col, CellStoreDelegates<T>.SetValueDelegate setStyleInnerUpdate, object value);
		void Clear(int _fromRow, int _fromCol, int toRow, int toCol);
		void Delete(int fromRow, int fromCol, int rows, int cols, bool shift);
		bool GetDimension(out int fromRow, out int fromCol, out int toRow, out int toCol);
		void Insert(int rowFrom, int v1, int rows, int v2);
	}
}