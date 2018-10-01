/* Copyright (C) 2011  Jan Källman
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "as is" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
 * Author							Change						Date
 *******************************************************************************
 * Mats Alm   		                Added		                2013-12-03
 *******************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions
{
	/// <summary>
	/// Base class for Excel function implementations.
	/// </summary>
	public abstract class ExcelFunction
	{
		#region Class Variables
		private readonly ArgumentCollectionUtil _argumentCollectionUtil;
		private readonly ArgumentParsers _argumentParsers;
		private readonly CompileResultValidators _compileResultValidators;
		#endregion
		
		#region Properties
		/// <summary>
		/// Indicates whether or not the function's compiler should resolve arguments as ranges.
		/// </summary>
		public bool ResolveArgumentsAsRange { get; protected set; } = false;
		#endregion

		#region Constructors
		/// <summary>
		/// Default constructor for a <see cref="ExcelFunction"/>.
		/// </summary>
		public ExcelFunction()
			 : this(new ArgumentCollectionUtil(), new ArgumentParsers(), new CompileResultValidators())
		{
		}

		/// <summary>
		/// Instantiates a new <see cref="ExcelFunction"/> with the specified parameters.
		/// </summary>
		/// <param name="argumentCollectionUtil">The argument collection utility.</param>
		/// <param name="argumentParsers">The argument parsers to use.</param>
		/// <param name="compileResultValidators">The compile result validators to use.</param>
		public ExcelFunction(
			 ArgumentCollectionUtil argumentCollectionUtil,
			 ArgumentParsers argumentParsers,
			 CompileResultValidators compileResultValidators)
		{
			_argumentCollectionUtil = argumentCollectionUtil;
			_argumentParsers = argumentParsers;
			_compileResultValidators = compileResultValidators;
		}
		#endregion

		#region Public Abstract Methods
		/// <summary>
		/// Executes the function.
		/// </summary>
		/// <param name="arguments">Arguments to the function, each argument can contain primitive types, lists or <see cref="ExcelDataProvider.IRangeInfo">Excel ranges</see></param>
		/// <param name="context">The <see cref="ParsingContext"/> contains various data that can be useful in functions.</param>
		/// <returns>A <see cref="CompileResult"/> containing the calculated value</returns>
		public abstract CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context);
		#endregion

		#region Public Virtual Methods
		/// <summary>
		/// If overridden, this method is called before Execute is called.
		/// </summary>
		/// <param name="context"></param>
		public virtual void BeforeInvoke(ParsingContext context) { }
		#endregion

		#region Protected Methods
		protected object GetFirstValue(IEnumerable<FunctionArgument> val)
		{
			var arg = ((IEnumerable<FunctionArgument>)val).FirstOrDefault();
			if (arg.Value is ExcelDataProvider.IRangeInfo)
			{
				//var r=((ExcelDataProvider.IRangeInfo)arg);
				var r = arg.ValueAsRangeInfo;
				return r.GetValue(r.Address._fromRow, r.Address._fromCol);
			}
			else
			{
				return arg == null ? null : arg.Value;
			}
		}

		/// <summary>
		/// This functions validates that the supplied <paramref name="arguments"/> contains at least
		/// (the value of) <paramref name="minLength"/> elements. If one of the arguments is an
		/// <see cref="ExcelDataProvider.IRangeInfo">Excel range</see> the number of cells in
		/// that range will be counted as well. Additionally, if any of the given arguments are an
		/// <see cref="ExcelErrorValue"/>, that error value will be passed back in <paramref name="errorValue"/>.
		/// </summary>
		/// <param name="arguments">The arguments to validate.</param>
		/// <param name="minLength">The expected minimum number of elements in <paramref name="arguments"/>.</param>
		/// <param name="errorValue">The <see cref="eErrorType"/> contained in the first <see cref="ExcelErrorValue"/> encountered
		///								if this method returns false. The default value is the value given for <paramref name="errorOnInvalidCount"/>.</param>
		/// <param name="errorOnInvalidCount">The desired <see cref="eErrorType"/> to receive if this method returns false.</param>
		/// <returns>Returns true if there are at least the <paramref name="minLength"/> number of arguments present 
		///			 and none of the arguments contain an <see cref="ExcelErrorValue"/>, and returns false if otherwise.</returns>
		protected bool ArgumentsAreValid(IEnumerable<FunctionArgument> arguments, int minLength, out eErrorType errorValue, eErrorType errorOnInvalidCount = eErrorType.Value)
		{
			errorValue = errorOnInvalidCount;
			if (!this.ArgumentCountIsValid(arguments, minLength))
				return false;
			var argumentContainingError = arguments.FirstOrDefault(arg => arg.ValueIsExcelError);
			if (argumentContainingError != null)
			{
				errorValue = argumentContainingError.ValueAsExcelErrorValue.Type;
				return false;
			}
			return true;
		}

		/// <summary>
		/// This functions validates that the supplied <paramref name="arguments"/> contains at least
		/// (the value of) <paramref name="minLength"/> elements. If one of the arguments is an
		/// <see cref="ExcelDataProvider.IRangeInfo">Excel range</see> the number of cells in
		/// that range will be counted as well.
		/// </summary>
		/// <param name="arguments">The arguments to be evaluated.</param>
		/// <param name="minLength">The minimum number of arguments that need to be present.</param>
		/// <returns>Returns true if there are at least the <paramref name="minLength"/> number of arguments present, and returns false if otherwise.</returns>
		protected bool ArgumentCountIsValid(IEnumerable<FunctionArgument> arguments, int minLength)
		{
			return arguments?.Count() >= minLength;
		}

		/// <summary>
		/// Returns the value of the argument att the position of the 0-based
		/// <paramref name="index"/> as an integer.
		/// </summary>
		/// <param name="arguments">The list of function arguments where our input to parse is.</param>
		/// <param name="index">The index of the arguments to try to parse to an integer.</param>
		/// <returns>Value of the argument as an integer.</returns>
		/// <exception cref="ExcelErrorValueException"></exception>
		protected int ArgToInt(IEnumerable<FunctionArgument> arguments, int index)
		{
			var val = arguments.ElementAt(index).ValueFirst;
			if (val == null)
				return 0;
			return (int)_argumentParsers.GetParser(DataType.Integer).Parse(val);
		}

		/// <summary>
		/// Attempts to parse an argument into an integer value.
		/// </summary>
		/// <param name="arguments">The list of function arguments where our input to parse is.</param>
		/// <param name="index"> The index of the arguments to try to parse to an integer.</param>
		/// <param name="value">The resulting value if the parse was successful. If not the value is the minimum integer value.</param>
		/// <returns></returns>
		protected bool TryGetArgAsInt(IEnumerable<FunctionArgument> arguments, int index, out int value)
		{
			var intCandidate = arguments.ElementAt(index).Value;
			value = int.MinValue;
			if (intCandidate == null)
			{
				value = 0;
				return true;
			}
			else if (intCandidate is int numberInt)
			{
				value = numberInt;
				return true;
			}
			else if (intCandidate is double numberDouble)
			{
				value = (int)numberDouble;
				return true;
			}
			else if (intCandidate is string)
			{
				if (ConvertUtil.TryParseNumericString(intCandidate, out double result))
				{
					value = (int)result;
					return true;
				}
			}
			if (ConvertUtil.TryParseDateObject(intCandidate, out System.DateTime date, out eErrorType? error))
			{
				value = (int)date.ToOADate();
				return true;
			}
			return false;
		}

		/// <summary>
		/// Returns the value of the argument att the position of the 0-based
		/// <paramref name="index"/> as a string.
		/// </summary>
		/// <param name="arguments"></param>
		/// <param name="index"></param>
		/// <returns>Value of the argument as a string.</returns>
		protected string ArgToString(IEnumerable<FunctionArgument> arguments, int index)
		{
			var obj = arguments.ElementAt(index).ValueFirst;
			return obj != null ? obj.ToString() : string.Empty;
		}

		/// <summary>
		/// Returns the value of the argument att the position of the 0-based
		/// </summary>
		/// <param name="obj"></param>
		/// <returns>Value of the argument as a double.</returns>
		/// <exception cref="ExcelErrorValueException"></exception>
		protected double ArgToDecimal(object obj)
		{
			return (double)_argumentParsers.GetParser(DataType.Decimal).Parse(obj);
		}

		/// <summary>
		/// Returns the value of the argument att the position of the 0-based
		/// <paramref name="index"/> as a <see cref="System.Double"/>.
		/// </summary>
		/// <param name="arguments"></param>
		/// <param name="index"></param>
		/// <returns>Value of the argument as an integer.</returns>
		/// <exception cref="ExcelErrorValueException"></exception>
		protected double ArgToDecimal(IEnumerable<FunctionArgument> arguments, int index)
		{
			return ArgToDecimal(arguments.ElementAt(index).Value);
		}

		protected double Divide(double left, double right)
		{
			if (System.Math.Abs(right - 0d) < double.Epsilon)
			{
				throw new ExcelErrorValueException(eErrorType.Div0);
			}
			return left / right;
		}

		protected bool IsNumericString(object value)
		{
			if (value == null || string.IsNullOrEmpty(value.ToString())) return false;
			return Regex.IsMatch(value.ToString(), @"^[\d]+(\,[\d])?");
		}

		/// <summary>
		/// If the specified <paramref name="value"/> is a boolean value its value will be returned.
		/// If the <paramref name="value"/> is an integer value, true will be returned if its
		/// value is not 0, otherwise false.
		/// </summary>
		/// <param name="value">The value to parse to a boolean value.</param>
		/// <returns>True if the value coalesces to true, false otherwise.</returns>
		protected bool ArgToBool(object value)
		{
			return (bool)_argumentParsers.GetParser(DataType.Boolean).Parse(value ?? string.Empty);
		}

		/// <summary>
		/// If the specified <paramref name="argument"/>'s value is a boolean value its value will be returned.
		/// If the specified <paramref name="argument"/>'s value is an integer value, true will be returned if its
		/// value is not 0, otherwise false.
		/// </summary>
		/// <param name="argument">The argument to parse to a boolean value.</param>
		/// <returns>True if the <paramref name="argument"/>'s value coalesces to true, false otherwise.</returns>
		protected bool ArgToBool(FunctionArgument argument)
		{
			return this.ArgToBool(argument.Value);
		}

		/// <summary>
		/// If the argument is a collection, its first value will be returned.
		/// If the argument is not a collection, the argument will be returned.
		/// </summary>
		/// <param name="argument"></param>
		/// <returns></returns>
		protected FunctionArgument GetFirstArgument(FunctionArgument argument)
		{
			if (argument.Value is IEnumerable<FunctionArgument> enumerableArgument)
				return enumerableArgument.First();
			return argument;
		}

		/// <summary>
		/// If the argument is a collection, its first value will be returned.
		/// If the argument is not a collection, the argument will be returned.
		/// </summary>
		/// <param name="argument"></param>
		/// <returns></returns>
		protected object GetFirstArgument(object argument)
		{
			if (argument is IEnumerable<object> enumerableArgument)
				return enumerableArgument.First();
			return argument;
		}

		protected bool IsNumeric(object val)
		{
			if (val == null) return false;
			return (val.GetType().IsPrimitive || val is double || val is decimal || val is System.DateTime || val is TimeSpan);
		}

		/// <summary>
		/// Helper method for comparison of two doubles.
		/// </summary>
		/// <param name="d1"></param>
		/// <param name="d2"></param>
		/// <returns></returns>
		protected bool AreEqual(double d1, double d2)
		{
			return System.Math.Abs(d1 - d2) < double.Epsilon;
		}

		/// <summary>
		/// Will return the arguments as an enumerable of doubles.
		/// </summary>
		/// <param name="arguments"></param>
		/// <param name="context"></param>
		/// <returns></returns>
		protected virtual IEnumerable<double> ArgsToDoubleEnumerable(IEnumerable<FunctionArgument> arguments,
																						 ParsingContext context)
		{
			return ArgsToDoubleEnumerable(false, arguments, context);
		}

		/// <summary>
		/// Will return the arguments as an enumerable of doubles.
		/// </summary>
		/// <param name="ignoreHiddenCells">If a cell is hidden and this value is true the value of that cell will be ignored</param>
		/// <param name="ignoreErrors">If a cell contains an error, that error will be ignored if this method is set to true</param>
		/// <param name="arguments"></param>
		/// <param name="context"></param>
		/// <returns></returns>
		protected virtual IEnumerable<double> ArgsToDoubleEnumerable(bool ignoreHiddenCells, bool ignoreErrors, IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			return _argumentCollectionUtil.ArgsToDoubleEnumerable(ignoreHiddenCells, ignoreErrors, arguments, context);
		}

		/// <summary>
		/// Will return the arguments as an enumerable of doubles.
		/// </summary>
		/// <param name="ignoreHiddenCells">If a cell is hidden and this value is true the value of that cell will be ignored</param>
		/// <param name="arguments"></param>
		/// <param name="context"></param>
		/// <returns></returns>
		protected virtual IEnumerable<double> ArgsToDoubleEnumerable(bool ignoreHiddenCells, IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			return ArgsToDoubleEnumerable(ignoreHiddenCells, true, arguments, context);
		}

		/// <summary>
		/// Will return the arguments as an enumerable of objects.
		/// </summary>
		/// <param name="ignoreHiddenCells">If a cell is hidden and this value is true the value of that cell will be ignored</param>
		/// <param name="arguments"></param>
		/// <param name="context"></param>
		/// <returns></returns>
		protected virtual IEnumerable<object> ArgsToObjectEnumerable(bool ignoreHiddenCells, IEnumerable<FunctionArgument> arguments, ParsingContext context)
		{
			return _argumentCollectionUtil.ArgsToObjectEnumerable(ignoreHiddenCells, arguments, context);
		}

		/// <summary>
		/// Use this method to create a result to return from Excel functions. 
		/// </summary>
		/// <param name="result"></param>
		/// <param name="dataType"></param>
		/// <returns></returns>
		protected CompileResult CreateResult(object result, DataType dataType)
		{
			var validator = _compileResultValidators.GetValidator(dataType);

			if(!validator.TryValidateObjValueIsNotNaNOrinfinity(result, out eErrorType error))
				return new CompileResult(error);
			else
				validator.Validate(result);

			return new CompileResult(result, dataType);
		}

		/// <summary>
		/// Use this method to apply a function on a collection of arguments. The <paramref name="result"/>
		/// should be modifyed in the supplied <paramref name="action"/> and will contain the result
		/// after this operation has been performed.
		/// </summary>
		/// <param name="collection"></param>
		/// <param name="result"></param>
		/// <param name="action"></param>
		/// <returns></returns>
		protected virtual double CalculateCollection(IEnumerable<FunctionArgument> collection, double result, Func<FunctionArgument, double, double> action)
		{
			return _argumentCollectionUtil.CalculateCollection(collection, result, action);
		}

		/// <summary>
		/// if the supplied <paramref name="arg">argument</paramref> contains an Excel error
		/// an <see cref="ExcelErrorValueException"/> with that errorcode will be thrown
		/// </summary>
		/// <param name="arg"></param>
		/// <exception cref="ExcelErrorValueException"></exception>
		protected void CheckForAndHandleExcelError(FunctionArgument arg)
		{
			if (arg.ValueIsExcelError)
			{
				throw (new ExcelErrorValueException(arg.ValueAsExcelErrorValue));
			}
		}

		/// <summary>
		/// If the supplied <paramref name="cell"/> contains an Excel error
		/// an <see cref="ExcelErrorValueException"/> with that errorcode will be thrown
		/// </summary>
		/// <param name="cell"></param>
		protected void CheckForAndHandleExcelError(ExcelDataProvider.ICellInfo cell)
		{
			if (cell.IsExcelError)
			{
				throw (new ExcelErrorValueException(ExcelErrorValue.Parse(cell.Value.ToString())));
			}
		}

		protected CompileResult GetResultByObject(object result)
		{
			if (IsNumeric(result))
			{
				return CreateResult(result, DataType.Decimal);
			}
			if (result is string)
			{
				return CreateResult(result, DataType.String);
			}
			if (result is ExcelErrorValue)
			{
				return CreateResult(result, DataType.ExcelError);
			}
			if (result == null)
			{
				return CompileResult.Empty;
			}
			return CreateResult(result, DataType.Enumerable);
		}
	}
	#endregion
}