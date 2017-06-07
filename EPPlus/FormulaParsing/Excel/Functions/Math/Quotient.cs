using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	/// <summary>
	/// This class contains the formula for dividing two arguments.
	/// </summary>
	public class Quotient : ExcelFunction
	{
		/// <summary>
		/// Takes two user specified arguments and divides the first by the second. 
		/// </summary>
		/// <param name="arguments">The user specified arguments to divide.</param>
		/// <param name="context">Not used, but needed to override the method.</param>
		/// <returns>The first argument divided by the second argument as an integer value.</returns>
		public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			if (ValidateArguments(arguments, 2) == false)
				return new CompileResult(eErrorType.Value);
			var numeratorCandidate = arguments.ElementAt(0).Value;
			var denominatorCandidate = arguments.ElementAt(1).Value;
			double candidateAsDouble;

			if (numeratorCandidate == null ||denominatorCandidate == null)
				return new CompileResult(eErrorType.NA);
			if (!ConvertUtil.TryParseNumericString(numeratorCandidate, out candidateAsDouble))
				if (!ConvertUtil.TryParseDateObjectToOADate(numeratorCandidate, out candidateAsDouble))
					return new CompileResult(eErrorType.Value);
			if (!ConvertUtil.TryParseNumericString(denominatorCandidate, out candidateAsDouble))
				if (!ConvertUtil.TryParseDateObjectToOADate(denominatorCandidate, out candidateAsDouble))
					return new CompileResult(eErrorType.Value);

			var num = ArgToDecimal(arguments, 0);
			var denom = ArgToDecimal(arguments, 1);
			if (denom == 0.0)
				return new CompileResult(eErrorType.Div0);
			var result = (int)(num / denom);
			return CreateResult(result, DataType.Integer);
		}
	}
}
