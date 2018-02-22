using System;
using System.Linq;
using System.Text;

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
			this.TableName = tableName.ToString();
			// If there is no specifier then only a single bracket exists
			if (letters[currentIndex] != '[')
				currentIndex--;
			while (currentIndex < letters.Length)
			{
				var component = this.BuildComponent(letters, ref currentIndex);
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
				currentIndex++;
			}
			// Set default specifiers if none were specified
			if (this.ItemSpecifiers == 0)
				this.ItemSpecifiers = ItemSpecifiers.Data;
			if (this.EndColumn == null)
				this.EndColumn = this.StartColumn;
			// TODO :: some validation on the resulting parsed structured reference.
		}
		#endregion

		#region Private Methods
		private string BuildComponent(char[] letters, ref int currentIndex)
		{
			var letter = letters[currentIndex++];
			if (letter != '[')
				throw new ArgumentException("Malformed structured reference");
			StringBuilder component = new StringBuilder();
			for (;currentIndex < letters.Length; currentIndex++)
			{
				letter = letters[currentIndex];
				if (letter == ']')
					break;
				component.Append(letter);
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
