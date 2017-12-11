using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;

namespace OfficeOpenXml
{
	/// <summary>
	/// Contains external references.
	/// </summary>
	public class ExternalReferenceCollection : XmlHelper
	{
		#region Properties
		/// <summary>
		/// Gets a read-only list of the external references in the workbook.
		/// </summary>
		public IReadOnlyList<ExternalReference> References => this.PrivateReferences;

		private List<ExternalReference> PrivateReferences { get; } = new List<ExternalReference>();
		#endregion

		#region Constructors
		/// <summary>
		/// Creates an instance of an <see cref="ExternalReferenceCollection"/>.
		/// </summary>
		/// <param name="resolveName">A function that can resolve a node ID to a reference name.</param>
		/// <param name="rootNode">The <see cref="XmlNode"/> containing the references.</param>
		/// <param name="namespaceManager">The <see cref="XmlNamespaceManager"/> for external references.</param>
		internal ExternalReferenceCollection(Func<string, string> resolveName, XmlNode rootNode, XmlNamespaceManager namespaceManager) : base(namespaceManager, rootNode)
		{
			if (resolveName == null)
				throw new ArgumentNullException(nameof(resolveName));
			XmlNodeList nodeList = this.TopNode.SelectNodes("//d:externalReference", this.NameSpaceManager);
			if (nodeList != null)
			{
				int id = 1;
				foreach (XmlNode node in nodeList)
				{
					string rID = node.Attributes["r:id"].Value;
					var referenceName = resolveName(rID);
					this.PrivateReferences.Add(new ExternalReference(id++, referenceName, node));
				}
			}
		}
		#endregion

		#region Public Methods
		/// <summary>
		/// Removes an external reference from the workbook by ID.
		/// </summary>
		/// <param name="id">The ID of the reference to delete.</param>
		public void DeleteReference(int id)
		{
			var reference = this.PrivateReferences.FirstOrDefault(r => r.Id == id);
			if (reference != null)
			{
				this.PrivateReferences.Remove(reference);
				this.TopNode.RemoveChild(reference.Node);
			}
		}
		#endregion

		#region Nested Classes
		/// <summary>
		/// Defines an external workbook reference.
		/// </summary>
		public class ExternalReference
		{
			#region Properties
			/// <summary>
			/// Gets the ID of the external reference.
			/// </summary>
			public int Id { get; }

			/// <summary>
			/// Gets the Name of the external reference.
			/// </summary>
			public string Name { get; }

			/// <summary>
			/// Gets the <see cref="XmlNode"/> of the external reference.
			/// </summary>
			internal XmlNode Node { get; }
			#endregion

			#region Constructors
			/// <summary>
			/// Creates an instance of an <see cref="ExternalReference"/>.
			/// </summary>
			/// <param name="id">The ID of the reference.</param>
			/// <param name="name">The name of the reference.</param>
			/// <param name="node">The node containing the reference in the references collection.</param>
			internal ExternalReference(int id, string name, XmlNode node)
			{
				if (id < 1)
					throw new ArgumentOutOfRangeException(nameof(id));
				if (string.IsNullOrEmpty(name))
					throw new ArgumentNullException(nameof(name));
				if (node == null)
					throw new ArgumentNullException(nameof(node));
				this.Id = id;
				this.Name = name;
				this.Node = node;
			}
			#endregion
		}
		#endregion
	}
}
