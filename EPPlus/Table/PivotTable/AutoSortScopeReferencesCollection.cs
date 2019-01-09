using System.Collections.Generic;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
	/// <summary>
	/// Collection class for <see cref="AutoSortScopeReference"/>.
	/// </summary>
	public class AutoSortScopeReferencesCollection : ExcelPivotTableFieldCollectionBase<AutoSortScopeReference>
	{
		#region Constructors
		/// <summary>
		/// Creates an instance of a <see cref="AutoSortScopeReferencesCollection"/>.
		/// </summary>
		/// <param name="namespaceManager">The namespace manager.</param>
		/// <param name="node">The xml node.</param>
		/// <param name="table">The existing pivot table.</param>
		public AutoSortScopeReferencesCollection(XmlNamespaceManager namespaceManager, XmlNode node, ExcelPivotTable table) : base(namespaceManager, node, table) { }
		#endregion

		#region XmlCollectionBase Overrides
		/// <summary>
		/// Loads the <see cref="AutoSortScopeReference"/>s from the xml document.
		/// </summary>
		/// <returns>The collection of <see cref="AutoSortScopeReference"/>s.</returns>
		protected override List<AutoSortScopeReference> LoadItems()
		{
			var references = new List<AutoSortScopeReference>();
			foreach (XmlNode reference in base.TopNode.SelectNodes("d:reference", this.NameSpaceManager))
			{
				references.Add(new AutoSortScopeReference(this.NameSpaceManager, reference.FirstChild));
			}
			return references;
		}
		#endregion
	}
}
