using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.Excel.Functions;

namespace EPPlusTest.FormulaParsing.TestHelpers
{
	/// <summary>
	/// A static class containing helper methods for unit tests.
	/// </summary>
	public static class FunctionsHelper
	{
		#region Public Static Methods
		/// <summary>
		/// Creates an <see cref="IEnumerable{FunctionArgument}"/> from the specified <paramref name="arguments"/>.
		/// </summary>
		/// <param name="arguments">The arguments to convert into <see cref="FunctionArgument"/>s.</param>
		/// <returns>The specified <paramref name="arguments"/> as a collection of <see cref="FunctionArgument"/>s.</returns>
		public static IEnumerable<FunctionArgument> CreateArgs(params object[] arguments)
		{
			return arguments.Select(a => new FunctionArgument(a)).ToList();
		}
		#endregion
	}
}
