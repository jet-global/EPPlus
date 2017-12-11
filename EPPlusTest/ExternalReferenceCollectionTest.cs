using System;
using System.Xml;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using static OfficeOpenXml.ExternalReferenceCollection;

namespace EPPlusTest
{
	[TestClass]
	public class ExternalReferenceCollectionTest
	{
		#region Constructor Tests
		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void ExternalReferenceCollectionNullNameResolverThrowsException()
		{
			new ExternalReferenceCollection(null, null, null);
		}
		#endregion

		#region References Tests
		[TestMethod]
		public void References()
		{
			var document = new XmlDocument();
			document.LoadXml(@"<externalReferences xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">
	<externalReference r:id=""rId1"" />
	<externalReference r:id=""rId2"" />
	<externalReference r:id=""rId3"" />
</externalReferences>");
			var referenceCollection = new ExternalReferenceCollection(id =>
			{
				switch (id)
				{
					case "rId1": return "External Reference 1";
					case "rId2": return "External Reference 2";
					case "rId3": return "External Reference 3";
					default: throw new InvalidOperationException();
				}
			}, document.ChildNodes[0], this.GetNamespaceManager(document.NameTable));
			Assert.AreEqual(3, referenceCollection.References.Count);
			Assert.AreEqual(1, referenceCollection.References[0].Id);
			Assert.AreEqual("External Reference 1", referenceCollection.References[0].Name);
			Assert.AreEqual(2, referenceCollection.References[1].Id);
			Assert.AreEqual("External Reference 2", referenceCollection.References[1].Name);
			Assert.AreEqual(3, referenceCollection.References[2].Id);
			Assert.AreEqual("External Reference 3", referenceCollection.References[2].Name);
		}
		#endregion

		#region DeleteReference Tests
		[TestMethod]
		public void DeleteReference()
		{
			var document = new XmlDocument();
			document.LoadXml(@"<externalReferences xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">
	<externalReference r:id=""rId1"" />
	<externalReference r:id=""rId2"" />
	<externalReference r:id=""rId3"" />
</externalReferences>");
			var referenceCollection = new ExternalReferenceCollection(id =>
			{
				switch (id)
				{
					case "rId1": return "External Reference 1";
					case "rId2": return "External Reference 2";
					case "rId3": return "External Reference 3";
					default: throw new InvalidOperationException();
				}
			}, document.ChildNodes[0], this.GetNamespaceManager(document.NameTable));
			Assert.AreEqual(3, referenceCollection.References.Count);
			referenceCollection.DeleteReference(2);
			Assert.AreEqual(2, referenceCollection.References.Count);
			Assert.AreEqual(1, referenceCollection.References[0].Id);
			Assert.AreEqual("External Reference 1", referenceCollection.References[0].Name);
			Assert.AreEqual(3, referenceCollection.References[1].Id);
			Assert.AreEqual("External Reference 3", referenceCollection.References[1].Name);
			Assert.AreEqual(2, document.ChildNodes[0].ChildNodes.Count);
		}

		[TestMethod]
		public void DeleteReferenceReferenceNotFoundDoesNothing()
		{
			var document = new XmlDocument();
			document.LoadXml(@"<externalReferences xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">
	<externalReference r:id=""rId1"" />
	<externalReference r:id=""rId2"" />
	<externalReference r:id=""rId3"" />
</externalReferences>");
			var referenceCollection = new ExternalReferenceCollection(id =>
			{
				switch (id)
				{
					case "rId1": return "External Reference 1";
					case "rId2": return "External Reference 2";
					case "rId3": return "External Reference 3";
					default: throw new InvalidOperationException();
				}
			}, document.ChildNodes[0], this.GetNamespaceManager(document.NameTable));
			Assert.AreEqual(3, referenceCollection.References.Count);
			referenceCollection.DeleteReference(0);
			Assert.AreEqual(1, referenceCollection.References[0].Id);
			Assert.AreEqual("External Reference 1", referenceCollection.References[0].Name);
			Assert.AreEqual(2, referenceCollection.References[1].Id);
			Assert.AreEqual("External Reference 2", referenceCollection.References[1].Name);
			Assert.AreEqual(3, referenceCollection.References[2].Id);
			Assert.AreEqual("External Reference 3", referenceCollection.References[2].Name);
			Assert.AreEqual(3, document.ChildNodes[0].ChildNodes.Count);
		}
		#endregion

		#region ExternalReference Tests
		[TestMethod]
		[ExpectedException(typeof(ArgumentOutOfRangeException))]
		public void ExternalReferenceIdOutOfRangeThrowsException()
		{
			new ExternalReference(0, "name", new XmlDocument().ChildNodes[0]);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void ExternalReferenceNullNameThrowsException()
		{
			new ExternalReference(1, null, new XmlDocument().ChildNodes[0]);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void ExternalReferenceEmptyNameThrowsException()
		{
			new ExternalReference(1, string.Empty, new XmlDocument().ChildNodes[0]);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void ExternalReferenceNullXmlNodeThrowsException()
		{
			new ExternalReference(1, "name", null);
		}
		#endregion

		#region Helper Methods
		private XmlNamespaceManager GetNamespaceManager(XmlNameTable nameTable)
		{
			var namespaceManager = new XmlNamespaceManager(nameTable);
			namespaceManager.AddNamespace("d", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
			return namespaceManager;
		}
		#endregion
	}
}
