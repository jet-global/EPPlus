using System.Xml;

namespace OfficeOpenXml
{
	/// <summary>
	/// Represents an Excel Theme XML element.
	/// </summary>
	public class ExcelTheme
	{
		#region Properties
		private XmlNamespaceManager NamespaceManager { get; }

		private XmlNode TopNode { get; }
		#endregion

		#region Constructors
		/// <summary>
		/// Constructor.
		/// </summary>
		/// <param name="document">The document containing the "Theme" xml.</param>
		/// <param name="namespaceManager">The namespace manager.</param>
		public ExcelTheme(XmlDocument document, XmlNamespaceManager namespaceManager)
		{
			namespaceManager.AddNamespace("a", ExcelPackage.schemaDrawings);
			this.NamespaceManager = namespaceManager;
			this.TopNode = document.SelectSingleNode("a:theme", this.NamespaceManager);
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Gets the color value at the specified index of the theme's color scheme.
		/// </summary>
		/// <param name="index">The index of the color in the color theme.</param>
		/// <returns>The color scheme.</returns>
		public string GetColorScheme(int index)
		{
			try
			{
				var colorSchemeNode = this.TopNode.SelectSingleNode("a:themeElements/a:clrScheme", this.NamespaceManager);
				var colorScheme = colorSchemeNode.ChildNodes[index];
				var value = colorScheme.FirstChild.Attributes["val"].Value;
				return value;
			}
			catch { return null; }
		}

		#endregion
	}
}
