using System.Collections.Generic;

namespace OfficeOpenXml.Table.PivotTable.Sorting
{
	/// <summary>
	/// Base class for comparing date strings in cache field group items.
	/// </summary>
	public abstract class DateComparerBase : IComparer<string>
	{
		#region Properties
		/// <summary>
		/// A dictionary of string to int that defines the sort position of a date string.
		/// </summary>
		protected abstract IReadOnlyDictionary<string, int> Mapping { get; }
		#endregion

		#region Public Methods
		/// <summary>
		/// Compares two strings.
		/// </summary>
		/// <param name="x">The left string to compare.</param>
		/// <param name="y">The right string to compare.</param>
		/// <returns>A compare value.</returns>
		public int Compare(string x, string y)
		{
			if (this.Mapping.ContainsKey(x))
			{
				if (this.Mapping.ContainsKey(y))
				{
					var xIndex = this.Mapping[x];
					var yIndex = this.Mapping[y];
					if (xIndex < yIndex)
						return -1;
					if (xIndex > yIndex)
						return 1;
					return 0;
				}
				return -1;
			}
			if (this.Mapping.ContainsKey(y))
				return 1;
			return string.Compare(x, y);
		}
		#endregion
	}
}
