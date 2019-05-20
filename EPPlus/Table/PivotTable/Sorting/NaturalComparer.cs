using System.Collections.Generic;

namespace OfficeOpenXml.Table.PivotTable.Sorting
{
	/// <summary>
	/// Compares alphanumeric strings, comparing numeric strings as numbers.
	/// </summary>
	public class NaturalComparer : IComparer<string>
	{
		#region Public Methods
		/// <summary>
		/// Compares two strings.
		/// </summary>
		/// <param name="x">The first string to compare.</param>
		/// <param name="y">The second string to compare.</param>
		/// <returns>The comparison result.</returns>
		public int Compare(string x, string y)
		{
			bool xIsNumeric = double.TryParse(x, out var xDouble);
			bool yIsNumeric = double.TryParse(y, out var yDouble);
			if (xIsNumeric && yIsNumeric)
			{
				if (xDouble < yDouble)
					return -1;
				else if (yDouble < xDouble)
					return 1;
				return 0;
			}
			if (xIsNumeric)
				return -1;
			if (yIsNumeric)
				return 1;
			return string.Compare(x, y);
		}
		#endregion
	}
}
