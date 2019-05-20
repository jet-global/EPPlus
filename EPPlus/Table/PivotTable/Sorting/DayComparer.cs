using System.Collections.Generic;

namespace OfficeOpenXml.Table.PivotTable.Sorting
{
	/// <summary>
	/// Compares day of week strings in cache field group items.
	/// </summary>
	public class DayComparer : DateComparerBase
	{
		#region Class Variables
		IReadOnlyDictionary<string, int> myMapping;
		#endregion

		#region Properties
		/// <summary>
		/// A mapping of abbreviated days of the week to their sort position.
		/// </summary>
		protected override IReadOnlyDictionary<string, int> Mapping
		{
			get
			{
				if (myMapping == null)
					myMapping = new Dictionary<string, int>
					{
						{ "Sun", 0 },
						{ "Mon", 1 },
						{ "Tue", 2 },
						{ "Wed", 3 },
						{ "Thu", 4 },
						{ "Fri", 5 },
						{ "Sat", 6 }
					};
				return myMapping;
			}
		}
		#endregion
	}
}
