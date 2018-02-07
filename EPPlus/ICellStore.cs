using System;
using System.Collections.Generic;

namespace OfficeOpenXml
{
	/// <summary>
	/// Represents a generic interface that defines how a CellStore data structure can be accessed and updated. 
	/// </summary>
	/// <typeparam name="T">The type of object being stored in the CellStore.</typeparam>
	internal interface ICellStore<T> : IDisposable
	{
		#region Properties
		/// <summary>
		/// Gets the maximum row of the cellstore.
		/// </summary>
		int MaximumRow { get; }

		/// <summary>
		/// Gets the maximum column of the cellstore.
		/// </summary>
		int MaximumColumn { get; }
		#endregion

		#region Methods
		/// <summary>
		/// Get the value at a particular location.
		/// </summary>
		/// <param name="row">The row to read a value from.</param>
		/// <param name="column">The column to read a value from.</param>
		/// <returns>The value in the cell coordinate.</returns>
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

		/// <summary>
		/// Gets a default enumerator for the cellstore.
		/// </summary>
		/// <returns>The default enumerator for the cellstore.</returns>
		ICellStoreEnumerator<T> GetEnumerator();

		/// <summary>
		/// Gets a contained enumerator for the cellstore.
		/// </summary>
		/// <param name="startRow">The minimum row to enumerate.</param>
		/// <param name="startColumn">The minimum column to enumerate.</param>
		/// <param name="endRow">The maximum row to enumerate.</param>
		/// <param name="endColumn">The maximum column to enumerate.</param>
		/// <returns>The constrained enumerator for the cellstore.</returns>
		ICellStoreEnumerator<T> GetEnumerator(int startRow, int startColumn, int endRow, int endColumn);
		#endregion
	}

	/// <summary>
	/// Represents an interface used to store flags with cell metadata.
	/// </summary>
	internal interface IFlagStore : ICellStore<byte>
	{
		#region Methods
		/// <summary>
		/// Adds or removes the given <paramref name="cellFlags"/> value based on <paramref name="value"/>.
		/// </summary>
		/// <param name="Row">The cell row to set a flag for.</param>
		/// <param name="Col">The cell column to set a flag for.</param>
		/// <param name="value">A boolean value indicating whether to add or remove the flag value.</param>
		/// <param name="cellFlags">The flags to set for the cell.</param>
		void SetFlagValue(int Row, int Col, bool value, CellFlags cellFlags);

		/// <summary>
		/// Gets the flag values from a given cell.
		/// </summary>
		/// <param name="Row">The cell row to get a flag for.</param>
		/// <param name="Col">The cell column to get a flag for.</param>
		/// <param name="cellFlags">The flags to query for.</param>
		/// <returns>True if the flags are set; otherwise false.</returns>
		bool GetFlagValue(int Row, int Col, CellFlags cellFlags);
		#endregion
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
	/// This class will build the requested ICellStore type.
	/// Currently it uses a hardcoded flag to toggle between the new implementation
	/// and the original. Soon the old CellStore will be deleted but this structure 
	/// will remain in place just so there is one place to always call into for a new 
	/// ICellStore. This will simplify future development should we ever revisit this.
	/// </summary>
	internal static class CellStore
	{
		#region Properties
		static bool UseZCellStore { get; } = true;
		#endregion

		#region Public Static Methods
		/// <summary>
		/// Returns a new <see cref="ICellStore{T}"/>.
		/// </summary>
		/// <typeparam name="T">The type the new cellstore will contain.</typeparam>
		/// <returns>The new <see cref="ICellStore{T}"/>.</returns>
		public static ICellStore<T> Build<T>()
		{
			if (CellStore.UseZCellStore)
				return new ZCellStore<T>();
			return new CellStore<T>();
		}

		/// <summary>
		/// Returns a new <see cref="IFlagStore"/>.
		/// </summary>
		/// <returns>The new <see cref="IFlagStore"/>.</returns>
		public static IFlagStore BuildFlagStore()
		{
			if (CellStore.UseZCellStore)
				return new ZFlagStore();
			return new FlagCellStore();
		}
		#endregion
	}
}