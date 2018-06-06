using System;
using System.Collections.Generic;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Math
{
	/// <summary>
	/// A base class for coalescing arguments for the MIN and MAX Excel functions.
	/// </summary>
	public abstract class MinMaxBase : HiddenValuesHandlingFunction
	{
		#region Constructors
		/// <summary>
		/// Instantiates a new <see cref="MinMaxBase"/> base class.
		/// </summary>
		public MinMaxBase()
		{
			// Min/Max functions handle values differently depending on whether or not they
			// originated from a range argument or a static argument, so the compiler must
			// resolve all range arguments as such in order to allow us to distinguish between range
			// and static values.
			this.ResolveArgumentsAsRange = true;
		}
		#endregion

		#region Protected Methods
		/// <summary>
		/// Coalesces the potentially numeric MIN/MAX-able values in the specified <paramref name="arguments"/>.
		/// </summary>
		/// <param name="arguments">The arguments containing potential numeric values.</param>
		/// <param name="context">The context with which the calling function executes within.</param>
		/// <param name="error">The type of error that occurs, if any.</param>
		/// <returns>A collection of numeric objects.</returns>
		protected IEnumerable<double> GroupArguments(IEnumerable<FunctionArgument> arguments, ParsingContext context, out eErrorType? error)
		{
			error = null;
			if (this.ArgumentsAreValid(arguments, 1, out eErrorType errorValue) == false)
			{
				error = errorValue;
				return null;
			}
			var argumentValueList = this.ArgsToObjectEnumerable(false, arguments, context);
			foreach (var item in argumentValueList)
			{
				if (item is ExcelErrorValue excelErrorValue)
				{
					error = excelErrorValue.Type;
					return null;
				}
			}
			return this.GetNumberArguments(arguments, out error);
		}
		#endregion

		#region Private Methods
		private List<double> GetNumberArguments(IEnumerable<FunctionArgument> arguments, out eErrorType? error)
		{
			// Values that are retrieved from cell references must be numbers or dates to be included in MIN or MAX.
			// Strings that can be converted to either a date or a number are included if they are static arguments.
			List<double> doubleArgs = new List<double>();
			error = null;
			foreach (var argument in arguments)
			{
				if (argument.ExcelStateFlagIsSet(ExcelCellState.HiddenCell))
					continue;
				else if (argument.IsExcelRange)
				{
					foreach (var cell in argument.ValueAsRangeInfo)
					{
						if (this.TryGetDoubleValue(cell.Value, out double doubleValue))
							doubleArgs.Add(doubleValue);
					}
				}
				else if (this.TryGetDoubleValue(argument.Value, out double doubleValue))
					doubleArgs.Add(doubleValue);
				else if (argument.Value is string stringValue)
				{
					if (ConvertUtil.TryParseNumericString(stringValue, out double result))
						doubleArgs.Add(result);
					else if (ConvertUtil.TryParseDateString(stringValue, out System.DateTime dateResult))
						doubleArgs.Add(dateResult.ToOADate());
					else
					{
						error = eErrorType.Value;
						return null;
					}
				}
				else if (argument.Value is bool boolValue)
					doubleArgs.Add(boolValue ? 1 : 0);
				else if (argument.Value is IEnumerable<object> enumerableValues)
				{
					foreach (var value in enumerableValues)
					{
						if (value is FunctionArgument argValue && this.TryGetDoubleValue(argValue.Value, out doubleValue))
							doubleArgs.Add(doubleValue);
					}
				}
			}
			return doubleArgs;
		}

		private bool TryGetDoubleValue(object value, out double doubleValue)
		{
			doubleValue = 0;
			if (value == null)
				return false;
			if ((value.GetType().IsPrimitive && value is bool == false) || value is decimal)
				doubleValue = Convert.ToDouble(value);
			else if (value is System.DateTime dateTimeArg)
				doubleValue = dateTimeArg.ToOADate();
			else
				return false;
			return true;
		}
		#endregion
	}
}
