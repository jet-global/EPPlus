using System;
using System.Collections.Generic;

namespace OfficeOpenXml
{
	/// <summary>
	/// Represents the error types in excel.
	/// </summary>
	public enum eErrorType
	{
		/// <summary>
		/// Division by zero
		/// </summary>
		Div0 = -2146826281,
		/// <summary>
		/// Not applicable
		/// </summary>
		NA = -2146826246,
		/// <summary>
		/// Name error
		/// </summary>
		Name = -2146826259,
		/// <summary>
		/// Null error
		/// </summary>
		Null = -2146826288,
		/// <summary>
		/// Num error
		/// </summary>
		Num = -2146826252,
		/// <summary>
		/// Reference error
		/// </summary>
		Ref = -2146826265,
		/// <summary>
		/// Value error
		/// </summary>
		Value = -2146826273
	}

	/// <summary>
	/// Represents an Excel error.
	/// </summary>
	/// <seealso cref="eErrorType"/>
	public class ExcelErrorValue
	{
		#region Properties
		/// <summary>
		/// The error type
		/// </summary>
		public eErrorType Type { get; private set; }
		#endregion

		#region Constructors
		private ExcelErrorValue(eErrorType type)
		{
			if (type == default(eErrorType))
				throw new ArgumentException($"{nameof(type)} must be a valid error type.");
			this.Type = type;
		}
		#endregion

		#region Internal Static Methods
		/// <summary>
		/// Creates a new <see cref="ExcelErrorValue"/> with the specified <paramref name="errorType"/>.
		/// </summary>
		/// <param name="errorType">The type of error to create.</param>
		/// <returns>An <see cref="ExcelErrorValue"/>.</returns>
		internal static ExcelErrorValue Create(eErrorType errorType)
		{
			return new ExcelErrorValue(errorType);
		}

		/// <summary>
		/// Parse the specified <paramref name="val"/> to an <see cref="ExcelErrorValue"/>.
		/// </summary>
		/// <param name="val">The value to attempt to parse.</param>
		/// <returns>An <see cref="ExcelErrorValue"/> matching the the specified <paramref name="val"/>.</returns>
		internal static ExcelErrorValue Parse(string val)
		{
			if (Values.TryGetErrorType(val, out eErrorType errorType))
				return new ExcelErrorValue(errorType);
			if (string.IsNullOrEmpty(val))
				throw new ArgumentNullException(nameof(val));
			throw new ArgumentException("Not a valid error value: " + val);
		}
		#endregion

		#region Object Overrides
		/// <summary>
		/// Returns the string representation of the error type.
		/// </summary>
		/// <returns>A string representation of the error type.</returns>
		public override string ToString()
		{
			switch (this.Type)
			{
				case eErrorType.Div0:
					return Values.Div0;
				case eErrorType.NA:
					return Values.NA;
				case eErrorType.Name:
					return Values.Name;
				case eErrorType.Null:
					return Values.Null;
				case eErrorType.Num:
					return Values.Num;
				case eErrorType.Ref:
					return Values.Ref;
				case eErrorType.Value:
					return Values.Value;
				default:
					throw new ArgumentException("Invalid error type");
			}
		}

		/// <summary>
		/// Compares the specified <paramref name="obj"/> to this for equality.
		/// </summary>
		/// <param name="obj">The object to compare for equality.</param>
		/// <returns>True if the objects are equal, otherwise false.</returns>
		public override bool Equals(object obj)
		{
			if (!(obj is ExcelErrorValue errorValueObj))
				return false;
			return errorValueObj.Type == this.Type;
		}

		/// <summary>
		/// Gets the hash code for this object.
		/// </summary>
		/// <returns>An integer hash code for this object.</returns>
		public override int GetHashCode()
		{
			return 2049151605 + this.Type.GetHashCode();
		}
		#endregion

		#region Nested Classes
		/// <summary>
		/// Handles the convertion between <see cref="eErrorType"/> and the string values
		/// used by Excel.
		/// </summary>
		public static class Values
		{
			#region Constants
			/// <summary>
			/// The string representation of a divide by zero error.
			/// </summary>
			public const string Div0 = "#DIV/0!";
			/// <summary>
			/// The string representation of a not applicable error.
			/// </summary>
			public const string NA = "#N/A";
			/// <summary>
			/// The string representation of a unknown name error.
			/// </summary>
			public const string Name = "#NAME?";
			/// <summary>
			/// The string representation of a null error.
			/// </summary>
			public const string Null = "#NULL!";
			/// <summary>
			/// The string representation of an invalid number error.
			/// </summary>
			public const string Num = "#NUM!";
			/// <summary>
			/// The string representation of an invalid reference error.
			/// </summary>
			public const string Ref = "#REF!";
			/// <summary>
			/// The string representation of a value error.
			/// </summary>
			public const string Value = "#VALUE!";
			#endregion

			#region Static Class Variables
			private static Dictionary<string, eErrorType> _values = new Dictionary<string, eErrorType>()
			{
				{Div0, eErrorType.Div0},
				{NA, eErrorType.NA},
				{Name, eErrorType.Name},
				{Null, eErrorType.Null},
				{Num, eErrorType.Num},
				{Ref, eErrorType.Ref},
				{Value, eErrorType.Value}
			};
			#endregion

			#region Public Methods
			/// <summary>
			/// Tries to convert a string to an <see cref="eErrorType"/>.
			/// </summary>
			/// <param name="candidate">The string to convert to an error value.</param>
			/// <param name="eErrorType">The converted <see cref="eErrorType"/>.</param>
			/// <returns>True if succesfully converted, otherwise false.</returns>
			public static bool TryGetErrorType(string candidate, out eErrorType eErrorType)
			{
				return _values.TryGetValue(candidate, out eErrorType);
			}
			#endregion
		}
		#endregion
	}
}
