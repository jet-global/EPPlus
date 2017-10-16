using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup
{
	/// <summary>
	/// A function which takes the same arguments as the OFFSET( ) function and calculates the address offset.
	/// </summary>
	/// <remarks>
	/// This function is used internally to calculate an OFFSET( ) function's dependencies before calculation.
	/// </remarks>
	public class OffsetAddress : LookupFunction
	{
		#region Constants
		public const string Name = "OFFSETADDRESS";
		#endregion

		#region ExcelFunction Overrides
		/// <summary>
		/// Executes the function with the specified <paramref name="arguments"/> in the specified <paramref name="context"/>.
		/// </summary>
		/// <param name="arguments">The arguments with which to evaluate the function.</param>
		/// <param name="context">The context in which to evaluate the function.</param>
		/// <returns>An address range <see cref="CompileResult"/> if successful, otherwise an error result.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			var functionArguments = arguments as FunctionArgument[] ?? arguments.ToArray();
			if (this.ArgumentsAreValid(functionArguments, 3, out eErrorType argumentError) == false)
				return new CompileResult(argumentError);
			ExcelAddress offset = base.CalculateOffset(functionArguments, context);
			if (offset == null)
				return new CompileResult(eErrorType.Ref);
			return new CompileResult(offset.FullAddress, DataType.String);
		}
		#endregion
	}
}
