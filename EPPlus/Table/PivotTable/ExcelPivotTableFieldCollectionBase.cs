using System;
using System.Collections.Generic;

namespace OfficeOpenXml.Table.PivotTable
{
	/// <summary>
	/// Abstract base collection class for pivot table fields.
	/// </summary>
	/// <typeparam name="T">An instance of {T}.</typeparam>
	public abstract class ExcelPivotTableFieldCollectionBase<T> : IEnumerable<T>
	{
		#region Class Variables
		/// <summary>
		/// The pivot table.
		/// </summary>
		protected ExcelPivotTable myTable;

		/// <summary>
		/// A list of fields.
		/// </summary>
		internal List<T> myList = new List<T>();
		#endregion

		#region Properties
		/// <summary>
		/// Gets the generic enumerator of the list.
		/// </summary>
		/// <returns>The enumerator.</returns>
		public IEnumerator<T> GetEnumerator()
		{
			return myList.GetEnumerator();
		}

		/// <summary>
		/// Gets the specified type enumerator of the list.
		/// </summary>
		/// <returns>The enumerator.</returns>
		System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
		{
			return myList.GetEnumerator();
		}

		/// <summary>
		/// Gets the count of fields in the pivot table.
		/// </summary>
		public int Count
		{
			get
			{
				return myList.Count;
			}
		}

		/// <summary>
		/// Gets the field at the given index.
		/// </summary>
		/// <param name="Index">The position of the field in the list.</param>
		/// <returns>The pivot table field.</returns>
		public T this[int Index]
		{
			get
			{
				if (Index < 0 || Index >= myList.Count)
					throw (new ArgumentOutOfRangeException("Index out of range"));
				return myList[Index];
			}
		}
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="ExcelPivotTableFieldCollection"/>.
		/// </summary>
		/// <param name="table">The existing pivot table.</param>
		internal ExcelPivotTableFieldCollectionBase(ExcelPivotTable table)
		{
			myTable = table;
		}
		#endregion

		#region Methods
		/// <summary>
		/// Adds a field to the collection.
		/// </summary>
		/// <param name="field"></param>
		internal void AddInternal(T field)
		{
			myList.Add(field);
		}

		/// <summary>
		/// Clears the field collection.
		/// </summary>
		internal void Clear()
		{
			myList.Clear();
		}
		#endregion
	}
}