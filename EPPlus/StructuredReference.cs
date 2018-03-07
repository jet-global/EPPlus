using System;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeOpenXml.FormulaParsing.Utilities;

namespace OfficeOpenXml
{
	/// <summary>
	/// NOTES:: Structured references in the Excel GUI are different than the ones
	/// in the XML. This class is designed to support parsing the XML variation in order
	/// to support function evaluation. In practice, the differences are not significant,
	/// mostly the XML is fully-qualified as opposed to unqualified.
	/// </summary>
	public class StructuredReference
	{
		#region Constants
		private const string MalformedExceptionMessage = "Malformed structured reference.";
		#endregion

		#region Properties
		/// <summary>
		/// Gets the table name for the structured reference.
		/// </summary>
		public string TableName { get; private set; }

		/// <summary>
		/// Gets the item specifiers for the structured reference.
		/// </summary>
		public ItemSpecifiers ItemSpecifiers { get; private set; }

		/// <summary>
		/// Gets the start column name for the structured reference.
		/// </summary>
		public string StartColumn { get; private set; }

		/// <summary>
		/// Gets the end column for the structured reference if it references multiple columns.
		/// </summary>
		public string EndColumn { get; private set; }

		private string OriginalReference { get; set; }
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="StructuredReference"/>.
		/// </summary>
		/// <param name="structuredReference">The reference string to parse.</param>
		public StructuredReference(string structuredReference)
		{
			if (string.IsNullOrEmpty(structuredReference))
				throw new ArgumentNullException(nameof(structuredReference));
			if (!Regex.IsMatch(structuredReference, RegexConstants.StructuredReference, RegexOptions.IgnoreCase | RegexOptions.Multiline))
				throw new ArgumentException(StructuredReference.MalformedExceptionMessage);
			this.OriginalReference = structuredReference;
			var letters = structuredReference.ToArray();
			int currentIndex = 0;
			char currentChar = letters[currentIndex++];
			StringBuilder tableName = new StringBuilder();
			do
			{
				tableName.Append(currentChar);
				currentChar = letters[currentIndex++];
			} while (currentChar != '[');
			this.MovePastWhitespace(letters, ref currentIndex);
			this.TableName = tableName.ToString();
			bool structuredReferenceIsSingleComponent = !Regex.IsMatch(new string(letters.Skip(currentIndex).ToArray()), "[^']\\[");
			while (currentIndex < letters.Length)
			{
				var component = this.BuildComponent(letters, ref currentIndex, structuredReferenceIsSingleComponent);
				switch (component.ToLower())
				{
					case "#all":
						this.ItemSpecifiers |= ItemSpecifiers.All;
						break;
					case "#data":
						this.ItemSpecifiers |= ItemSpecifiers.Data;
						break;
					case "#headers":
						this.ItemSpecifiers |= ItemSpecifiers.Headers;
						break;
					case "#totals":
						this.ItemSpecifiers |= ItemSpecifiers.Totals;
						break;
					case "#this row":
						this.ItemSpecifiers |= ItemSpecifiers.ThisRow;
						break;
					default:
						if (this.StartColumn == null)
							this.StartColumn = component;
						else
							this.EndColumn = component;
						break;
				}
				this.MoveToNextComponent(letters, ref currentIndex);
			}
			// Set default specifiers if none were specified
			if (this.ItemSpecifiers == default(ItemSpecifiers))
				this.ItemSpecifiers = ItemSpecifiers.Data;
			if (this.EndColumn == null)
				this.EndColumn = this.StartColumn;
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Checks that the item specifiers are valid for the structured reference.
		/// </summary>
		/// <returns>True if the item specifiers are valid, otherwise false.</returns>
		/// <remarks>
		/// Only one item specifier can be used for a structured reference except for 
		/// (#Data and #Headers) and (#Data and #Totals).
		/// </remarks>
		public bool HasValidItemSpecifiers()
		{
			if (this.ItemSpecifiers == default(ItemSpecifiers))
				return false;
			else if ((this.ItemSpecifiers & (this.ItemSpecifiers - 1)) != 0 
				&& this.ItemSpecifiers != (ItemSpecifiers.Data | ItemSpecifiers.Headers) 
				&& this.ItemSpecifiers != (ItemSpecifiers.Data | ItemSpecifiers.Totals))
			{
				return false;
			}
			return true;
		}
		#endregion

		#region Private Methods
		private void MovePastWhitespace(char[] letters, ref int currentIndex)
		{
			while (currentIndex < letters.Count() && char.IsWhiteSpace(letters[currentIndex]))
			{
				currentIndex++;
			}
		}

		private void MoveToNextComponent(char[] letters, ref int currentIndex)
		{
			this.MovePastWhitespace(letters, ref currentIndex);
			if (currentIndex < letters.Count() && (letters[currentIndex] == ',' || letters[currentIndex] == ':' || letters[currentIndex] == ']'))
				currentIndex++;
			this.MovePastWhitespace(letters, ref currentIndex);
		}

		private string BuildComponent(char[] letters, ref int currentIndex, bool structuredReferenceIsSingleComponent)
		{
			var letter = letters[currentIndex];
			bool inBrackets = letter == '[';
			if (inBrackets)
				currentIndex++;
			bool isEscaped = false;
			const char escapeCharacter = '\'';
			StringBuilder component = new StringBuilder();
			for (; currentIndex < letters.Length; currentIndex++)
			{
				letter = letters[currentIndex];
				// Stopping conditions
				if (letter == ']' && !isEscaped)
					break;
				else if (!structuredReferenceIsSingleComponent && !inBrackets && letter == ',')
					break;
				// Error cases
				if (!inBrackets && !isEscaped && (letter == '[' || letter == ']'))
					throw new ArgumentException(StructuredReference.MalformedExceptionMessage);
				// Building component
				if (letter != escapeCharacter)
				{
					component.Append(letter);
					isEscaped = false;
				}
				else if (letter == escapeCharacter && !isEscaped)
					isEscaped = true;
				else if (letter == escapeCharacter && isEscaped)
				{
					component.Append(letter);
					isEscaped = false;
				}
			}
			currentIndex++;
			return component.ToString();
		}
		#endregion
	}

	#region Enums
	/// <summary>
	/// NOTE:: Item Specifiers are Pascal Cased and Excel will revert any input
	/// to be so. Still, we will handle all casing because users can still type 
	/// it in however they'd like.
	/// </summary>
	[Flags]
	public enum ItemSpecifiers : int
	{
		All = 1,
		Data = 2,
		Headers = 4,
		Totals = 8,
		ThisRow = 16
	}
	#endregion
}
