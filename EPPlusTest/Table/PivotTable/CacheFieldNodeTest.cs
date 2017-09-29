using System;
using System.Xml;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Table.PivotTable;

namespace EPPlusTest.Table.PivotTable
{
	[TestClass]
	public class CacheFieldNodeTest
	{
		#region Constructor Tests
		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void SlicerRangeItemNullXmlNodeThrowsException()
		{
			new CacheFieldNode(null, this.CreateDefaultNSM());
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void SlicerRangeItemNullNamespaceManagerThrowsException()
		{
			new CacheFieldNode(this.CreateCacheFieldNode(), null);
		}
		#endregion

		#region Name Tests
		[TestMethod]
		public void SlicerRangeNodeName()
		{
			var node = this.CreateCacheFieldNode();
			var cacheFieldNode = new CacheFieldNode(node, this.CreateDefaultNSM());
			Assert.AreEqual("Customer No.", cacheFieldNode.Name);
			cacheFieldNode.Name = "Other Name";
			Assert.AreEqual("Other Name", cacheFieldNode.Name);
			Assert.AreEqual($@"<cacheField name=""Other Name"" numFmtId=""49"" xmlns=""{ExcelPackage.schemaMain}""><sharedItems count=""3""><s v=""10000"" /><s v=""20000"" /><s v=""30000"" /></sharedItems></cacheField>", node.OuterXml);
		}
		#endregion

		#region NumFormatId Tests
		[TestMethod]
		public void SlicerRangeNodeNumFormatId()
		{
			var node = this.CreateCacheFieldNode();
			var cacheFieldNode = new CacheFieldNode(node, this.CreateDefaultNSM());
			Assert.AreEqual("49", cacheFieldNode.NumFormatId);
			cacheFieldNode.NumFormatId = "30";
			Assert.AreEqual("30", cacheFieldNode.NumFormatId);
			Assert.AreEqual($@"<cacheField name=""Customer No."" numFmtId=""30"" xmlns=""{ExcelPackage.schemaMain}""><sharedItems count=""3""><s v=""10000"" /><s v=""20000"" /><s v=""30000"" /></sharedItems></cacheField>", node.OuterXml);
		}
		#endregion

		#region Items Tests
		[TestMethod]
		public void SlicerRangeNodeItems()
		{
			var node = this.CreateCacheFieldNode();
			var cacheFieldNode = new CacheFieldNode(node, this.CreateDefaultNSM());
			Assert.AreEqual(3, cacheFieldNode.Items.Count);
			Assert.AreEqual("10000", cacheFieldNode.Items[0].Value);
			Assert.AreEqual("20000", cacheFieldNode.Items[1].Value);
			Assert.AreEqual("30000", cacheFieldNode.Items[2].Value);
		}
		#endregion

		#region Helper Methods
		private XmlNode CreateCacheFieldNode()
		{
			var xmlDoc = new XmlDocument(this.CreateDefaultNSM().NameTable);
			xmlDoc.LoadXml(
			$@"<cacheField name=""Customer No."" numFmtId=""49"" xmlns=""{ExcelPackage.schemaMain}"">
				<sharedItems count=""3"">
					<s v=""10000""/>
					<s v=""20000""/>
					<s v=""30000""/>
				</sharedItems>
			</cacheField>");
			return xmlDoc.FirstChild;
		}

		private XmlNamespaceManager CreateDefaultNSM()
		{
			//  Create a NamespaceManager to handle the default namespace,
			//  and create a prefix for the default namespace:
			NameTable nt = new NameTable();
			var ns = new XmlNamespaceManager(nt);
			ns.AddNamespace(string.Empty, ExcelPackage.schemaMain);
			ns.AddNamespace("d", ExcelPackage.schemaMain);
			ns.AddNamespace("r", ExcelPackage.schemaRelationships);
			ns.AddNamespace("c", ExcelPackage.schemaChart);
			ns.AddNamespace("vt", ExcelPackage.schemaVt);
			// extended properties (app.xml)
			ns.AddNamespace("xp", ExcelPackage.schemaExtended);
			// custom properties
			ns.AddNamespace("ctp", ExcelPackage.schemaCustom);
			// core properties
			ns.AddNamespace("cp", ExcelPackage.schemaCore);
			// core property namespaces
			ns.AddNamespace("dc", ExcelPackage.schemaDc);
			ns.AddNamespace("dcterms", ExcelPackage.schemaDcTerms);
			ns.AddNamespace("dcmitype", ExcelPackage.schemaDcmiType);
			ns.AddNamespace("xsi", ExcelPackage.schemaXsi);
			ns.AddNamespace("x14", ExcelPackage.schemaMain2009);
			ns.AddNamespace("xm", ExcelPackage.schemaOfficeMain2006);
			return ns;
		}
		#endregion
	}
}
