using System.Collections.Generic;

namespace OfficeOpenXml.Table.PivotTable.Sorting
{
	/// <summary>
	/// Compares month strings in cache field group items.
	/// </summary>
	public class MonthComparer : DateComparerBase
	{
		#region Class Variables
		IReadOnlyDictionary<string, int> myMapping;
		#endregion

		#region Properties
		/// <summary>
		/// A mapping of the abbreviated months to their sort position.
		/// </summary>
		protected override IReadOnlyDictionary<string, int> Mapping
		{
			get
			{
				if (myMapping == null)
					myMapping = new Dictionary<string, int>
					{
						{ "Jan", 0 },
						{ "Feb", 1 },
						{ "Mar", 2 },
						{ "Apr", 3 },
						{ "May", 4 },
						{ "Jun", 5 },
						{ "Jul", 6 },
						{ "Aug", 7 },
						{ "Sep", 8 },
						{ "Oct", 9 },
						{ "Nov", 10 },
						{ "Dec", 11 }
					};
				return myMapping;
			}
		}
		#endregion
	}
}

