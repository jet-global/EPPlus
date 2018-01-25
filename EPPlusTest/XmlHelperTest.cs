using System.Xml;
using System.Xml.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest
{
	[TestClass]
	public class XmlHelperTest
	{
		[TestMethod]
		public void TransformValuesInNodeTest()
		{
			var documentNode = new XElement("documentelement",
					new XElement("basenodetype", "val")
			  );

			var doc = new XmlDocument();
			doc.LoadXml(documentNode.ToString());

			XmlHelper.TransformValuesInNode(doc, null, s => "updated " + s, ".//basenodetype");
			var node1 = doc.ChildNodes[0].ChildNodes[0];
			Assert.AreEqual("updated val", node1.FirstChild.Value);
			Assert.AreEqual("updated val", node1.InnerText);
			Assert.AreEqual("updated val", node1.InnerXml);
		}

		[TestMethod]
		public void TransformValuesInNodeMultipleTest()
		{
			var documentNode = new XElement("documentelement",
					new XElement("basenodetype",
						new XElement("secondarynodetype", "value11"),
						new XElement("secondarynodetype", "value12"),
						new XElement("anothernodetype", "value13")
					),
					new XElement("basenodetype",
						new XElement("secondarynodetype", "value21")
					)
			  );

			var doc = new XmlDocument();
			doc.LoadXml(documentNode.ToString());

			var basenode1 = doc.ChildNodes[0].ChildNodes[0];
			var basenode2 = doc.ChildNodes[0].ChildNodes[1];
			XmlHelper.TransformValuesInNode(basenode1, null, s => "updated " + s, ".//secondarynodetype");
			var node11 = basenode1.ChildNodes[0];
			var node12 = basenode1.ChildNodes[1];
			var node13 = basenode1.ChildNodes[2];
			var node21 = basenode2.ChildNodes[0];
			Assert.AreEqual("updated value11", node11.FirstChild.Value);
			Assert.AreEqual("updated value12", node12.FirstChild.Value);
			Assert.AreEqual("value13", node13.FirstChild.Value);
			Assert.AreEqual("value21", node21.FirstChild.Value);
		}

		[TestMethod]
		public void TransformAttributesInNodeTest()
		{
			var node1 =
			  new XElement("basenodetype",
					new XAttribute("attribute1", "value1"),
					new XAttribute("attribute2", "value2")
			  );
			var doc = new XmlDocument();
			doc.LoadXml(node1.ToString());
			XmlHelper.TransformAttributesInNode(doc, null, s => "updated " + s, ".//basenodetype", "attribute1");
			Assert.AreEqual("updated value1", doc.ChildNodes[0].Attributes["attribute1"].Value);
			Assert.AreEqual("value2", doc.ChildNodes[0].Attributes["attribute2"].Value);
		}

		[TestMethod]
		public void TransformAttributesInNodeMultiplesTest()
		{
			var node1 =
			  new XElement("basenodetype",
					new XAttribute("attribute1", "value1"),
					new XAttribute("attribute2", "value2"),
					new XElement("secondarynodetype",
						new XAttribute("attribute1", "value11")
					),
					new XElement("basenodetype",
						new XAttribute("attribute1", "value21")
					)
			  );
			var doc = new XmlDocument();
			doc.LoadXml(node1.ToString());
			XmlHelper.TransformAttributesInNode(doc, null, s => "updated " + s, ".//basenodetype", "attribute1");
			Assert.AreEqual("updated value1", doc.ChildNodes[0].Attributes["attribute1"].Value);
			Assert.AreEqual("value11", doc.ChildNodes[0].ChildNodes[0].Attributes["attribute1"].Value);
			Assert.AreEqual("updated value21", doc.ChildNodes[0].ChildNodes[1].Attributes["attribute1"].Value);
			Assert.AreEqual("value2", doc.ChildNodes[0].Attributes["attribute2"].Value);
		}
	}
}
