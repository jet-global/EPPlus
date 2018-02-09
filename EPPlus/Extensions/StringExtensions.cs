using System;

namespace OfficeOpenXml.Extensions
{
	/// <summary>
	/// A class containing extension methods for <see cref="string"/>.
	/// </summary>
	public static class StringExtensions
	{
		#region Public Static Methods
		/// <summary>
		/// Compares two <see cref="string"/> objects for equality.
		/// Case is ignored, and null and empty strings are considered equal.
		/// </summary>
		/// <param name="originalString">The <see cref="string"/> object to compare to <paramref name="stringToCompare"/>.</param>
		/// <param name="stringToCompare">The <see cref="string"/> object to compare to <paramref name="originalString"/>.</param>
		/// <returns>True if the <see cref="string"/> objects are equivalent; otherwise, false.</returns>
		public static bool IsEquivalentTo(this string originalString, string stringToCompare)
		{
			return string.Equals(originalString ?? string.Empty, stringToCompare ?? string.Empty,
				StringComparison.CurrentCultureIgnoreCase);
		}
		#endregion
	}
}
