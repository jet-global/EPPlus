using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Logical
{
	/// <summary>
	/// Represents the Excel SWITCH logical function.
	/// </summary>
	public class Switch : ExcelFunction
	{
		#region ExcelFunction Overrides
		/// <summary>
		/// Executes the SWITCH function.
		/// </summary>
		/// <param name="arguments">The arguments to supply to the function.</param>
		/// <param name="context">The context to evaluate the function in.</param>
		/// <returns>The <see cref="CompileResult"/> result of executing the function.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (this.ArgumentsAreValid(arguments, 3, out eErrorType argumentError) == false)
				return new CompileResult(eErrorType.Value);
			var value = arguments.First().ValueFirst;
			for (int i = 1; i + 1 < arguments.Count(); i += 2)
			{
				var valueToMatch = arguments.ElementAt(i).Value;
				if (valueToMatch.Equals(value))
					return new CompileResultFactory().Create(arguments.ElementAt(i + 1).Value);
			}
			if (arguments.Count() % 2 == 0)
				return new CompileResultFactory().Create(arguments.ElementAt(arguments.Count() - 1).Value);
			return new CompileResult(eErrorType.NA);
		}
		#endregion
	}
}

