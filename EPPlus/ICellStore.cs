﻿using System;
using System.Collections.Generic;

namespace OfficeOpenXml
{
	internal static class CellStoreDelegates<T>
	{
		internal delegate void SetRangeValueDelegate(List<T> list, int index, int row, int column, object value);

		internal delegate void SetValueDelegate(List<T> list, int index, object value);
	}

	/// <summary>
	/// Represents a generic interface that defines how a CellStore data structure can be accessed and updated. 
	/// </summary>
	/// <typeparam name="T">The type of object being stored in the CellStore.</typeparam>
	internal interface ICellStore<T> : IDisposable
	{
		/// <summary>
		/// Get the value at a particular location.
		/// </summary>
		/// <param name="row">The row to read a value from.</param>
		/// <param name="column">The column to read a value from.</param>
		/// <returns></returns>
		T GetValue(int row, int column);

		/// <summary>
		/// Set the value at a particular location.
		/// </summary>
		/// <param name="row">The row of the location to set a value at.</param>
		/// <param name="column">The column of the location to set a value at.</param>
		/// <param name="value">The value to store at the location.</param>
		void SetValue(int row, int column, T value);

		/// <summary>
		/// Determine if a value exists at the given location, and return it as an out parameter if found.
		/// </summary>
		/// <param name="row">The row to look for a value at.</param>
		/// <param name="column">The column to look for a value at.</param>
		/// <param name="value">The value found, if one exists.</param>
		/// <returns>True if a value was found at the location; false otherwise.</returns>
		bool Exists(int row, int column, out T value);

		/// <summary>
		/// Get the location of the next cell after the given row and column, if one exists.
		/// Relies on the assumption that cells are sorted by column and then sorted within each column by row.
		/// </summary>
		/// <param name="row">The row to start searching from, and also the new location's row, if one exists.</param>
		/// <param name="column">The column to start searching from, and also the new location's column, if one exists.</param>
		/// <returns>True if a next cell has been found and the row and column parameters have been updated; false otherwise.</returns>
		bool NextCell(ref int row, ref int column);

		/// <summary>
		/// Get the location of the first cell before the given row and column, if one exists.
		/// Relies on the assumption that cells are sorted by column and then sorted within each column by row.
		/// </summary>
		/// <param name="row">The row to start searching from, and also the new location's row, if one exists.</param>
		/// <param name="column">The column to start searching from, and also the new location's column, if one exists.</param>
		/// <returns>True if a previous cell has been found and the row and column parameters have been updated; false otherwise.</returns>
		bool PrevCell(ref int row, ref int column);

		/// <summary>
		/// Deletes rows and/or columns from the workbook. This deletes all existing nodes in the specified range, and updates the keys of all subsequent nodes to reflect their new positions after being shifted to fill the newly-vacated space.
		/// </summary>
		/// <param name="fromRow">The first row to delete.</param>
		/// <param name="fromCol">The first column to delete.</param>
		/// <param name="rows">The number of rows to delete.</param>
		/// <param name="columns">The number of columns to delete.</param>
		void Delete(int fromRow, int fromCol, int rows, int columns);


		/// <summary>
		/// Deletes rows and/or columns from the workbook. This deletes all existing nodes in the specified range.
		/// If <paramref name="shift"/> is set to true, all subsequent keys will be updated to reflect their new position after being shifted into the newly-vacated space.
		/// </summary>
		/// <param name="fromRow">The first row to delete.</param>
		/// <param name="fromCol">The first column to delete.</param>
		/// <param name="rows">The number of rows to delete.</param>
		/// <param name="columns">The number of columns to delete.</param>
		/// <param name="shift">Whether or not subsequent rows and columns should be shifted up / left to fill the space that was deleted, or left where they are.</param>
		void Delete(int fromRow, int fromCol, int rows, int columns, bool shift);

		/// <summary>
		/// Determine if a value exists at the given location.
		/// </summary>
		/// <param name="row">The row to look for a value at.</param>
		/// <param name="column">The column to look for a value at.</param>
		/// <returns>True if a value was found at the location; false otherwise.</returns>
		bool Exists(int row, int column);

		/// <summary>
		/// Removes the values in the specified range without updating cells below or to the right of the specified range.
		/// </summary>
		/// <param name="_fromRow">The first row whose cells should be cleared.</param>
		/// <param name="_fromCol">The first column whose cells should be cleared.</param>
		/// <param name="toRow">The last row whose values should be cleared.</param>
		/// <param name="toCol">The last column whose values should be cleared.</param>
		void Clear(int _fromRow, int _fromCol, int toRow, int toCol);

		/// <summary>
		/// Get the range of cells contained in this collection.
		/// </summary>
		/// <param name="fromRow">The first row contained in this collection.</param>
		/// <param name="fromCol">The first column contained in this collection.</param>
		/// <param name="toRow">The last row contained in this collection.</param>
		/// <param name="toCol">The last column contained in this collection.</param>
		/// <returns>True if the collection contains at least one cell (and therefore has a dimension); false if the collection contains no cells and thus has no dimension.</returns>
		bool GetDimension(out int fromRow, out int fromCol, out int toRow, out int toCol);

		/// <summary>
		/// "Insert space" into the cellStore by updating all keys beyond the specified row or column by the specified number of rows or columns.
		/// </summary>
		/// <param name="rowFrom">The row to start updating keys from.</param>
		/// <param name="columnFrom">The columnn to start updating keys from.</param>
		/// <param name="rows">The number of rows being inserted.</param>
		/// <param name="columns">The number of columns being inserted.</param>
		void Insert(int rowFrom, int columnFrom, int rows, int columns);
	}

	/// <summary>
	/// An interface defining the properties necessary when enumerating a CellStore.
	/// </summary>
	/// <typeparam name="T"></typeparam>
	internal interface ICellStoreEnumerator<T> : IEnumerable<T>, IEnumerator<T>
	{
		/// <summary>
		/// The string address, such as A1, of the current cell.
		/// </summary>
		string CellAddress { get; }

		/// <summary>
		/// The column number of the current cell.
		/// </summary>
		int Column { get; }

		/// <summary>
		/// The row of the current cell.
		/// </summary>
		int Row { get; }

		/// <summary>
		/// The value stored at the current cell.
		/// </summary>
		T Value { get; set; }
	}

	/// <summary>
	/// A factory class for generating CellStoreEnumerators for a given <see cref="ICellStore{T}"/>.
	/// </summary>
	/// <typeparam name="T">The type of the value being stored in the <see cref="ICellStore{T}"/>.</typeparam>
	internal class CellStoreEnumeratorFactory<T>
	{
		/// <summary>
		/// Enumerate the entire CellStore.
		/// </summary>
		/// <param name="cellStore">The CellStore to enumerate.</param>
		/// <returns>An <see cref="ICellStoreEnumerator{T}"/> that enumerates the entire <paramref name="cellStore"/>.</returns>
		public static ICellStoreEnumerator<T> GetNewEnumerator(ICellStore<T> cellStore)
		{
			//var specificCellStore = cellStore as CellStore<T>;
			//if(specificCellStore != null)
			//	return new CellStore<T>.CellsStoreEnumerator<T>(specificCellStore);
			var newCellStore = cellStore as BalancedCellStore<T>;
			if (newCellStore != null)
				return new BalancedStoreEnumerator<T>(newCellStore);
			throw new NotImplementedException($"No CellStoreEnumerator accepts the type {cellStore.GetType()}.");
		}

		/// <summary>
		/// Enumerate only the part of the <paramref name="cellStore"/> that lie within the given bounds.
		/// </summary>
		/// <param name="cellStore">The <see cref="ICellStore{T}"/> to partially enumerate.</param>
		/// <param name="startRow">The minimum row of cells to enumerate.</param>
		/// <param name="startColumn">The minimum column of cells to enumerate.</param>
		/// <param name="endRow">The maximum row to enumerate.</param>
		/// <param name="endColumn">The maximum column to enumerate.</param>
		/// <returns>An enumerator that only returns values within the given range.</returns>
		public static ICellStoreEnumerator<T> GetNewEnumerator(ICellStore<T> cellStore, int startRow, int startColumn, int endRow, int endColumn)
		{
			//var specificCellStore = cellStore as CellStore<T>;
			//if (specificCellStore != null)
			//	return new CellStore<T>.CellsStoreEnumerator<T>(specificCellStore, startRow, startColumn, endRow, endColumn);
			var newCellStore = cellStore as BalancedCellStore<T>;
			if (newCellStore != null)
				return new BalancedStoreEnumerator<T>(newCellStore, startRow, startColumn, endRow, endColumn);
			throw new NotImplementedException($"No CellStoreEnumerator accepts the type {cellStore.GetType()}.");
		}
	}
}