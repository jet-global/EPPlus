using System;

namespace OfficeOpenXml.Internationalization
{
	/// <summary>
	/// Attribute for <see cref="StringResources"/> properties that can be translated.
	/// </summary>
	[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
	public class StringResource : Attribute
	{
		#region Properties
		/// <summary>
		/// Gets the comment for the string resource.
		/// </summary>
		/// <remarks>
		/// Used to provide special instructions to translators. 
		/// </remarks>
		public string Comment { get; }
		#endregion

		#region Constructors
		/// <summary>
		/// Instantiates a new <see cref="StringResource"/> object.
		/// </summary>
		/// <param name="comment">The comment to attach to the attribute (optional).</param>
		public StringResource(string comment = null)
		{
			this.Comment = comment;
		}
		#endregion
	}
}
