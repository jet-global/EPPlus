using System;
using System.Xml;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Table.PivotTable;

namespace EPPlusTest.Table.PivotTable
{
	[TestClass]
	public class AutoSortScopeReferenceTest
	{
		#region Constructor Tests
		[TestMethod]
		public void AutoSortScopeReference()
		{
			XmlDocument document = new XmlDocument();
			document.LoadXml(@"<references count=""1""><reference field=""4294967294"" count=""1"" selected=""0""><x v=""0""/></reference></references>");
			var node = document.SelectSingleNode("//x");
			var namespaceManager = TestUtility.CreateDefaultNSM();
			var reference = new AutoSortScopeReference(namespaceManager, node);
			Assert.AreEqual(PivotCacheRecordType.x, reference.Type);
			Assert.AreEqual("0", reference.Value);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void AutoSortScopeReferenceNullNamespaceManager()
		{
			XmlDocument document = new XmlDocument();
			document.LoadXml(@"<references count=""1""><reference field=""4294967294"" count=""1"" selected=""0""><x v=""0""/></reference></references>");
			var node = document.SelectSingleNode("//x");
			new AutoSortScopeReference(null, node);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void AutoSortScopeReferenceNullNode()
		{
			new AutoSortScopeReference(TestUtility.CreateDefaultNSM(), null);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentException))]
		public void AutoSortScopeReferenceIncorrectParentNode()
		{
			XmlDocument document = new XmlDocument();
			document.LoadXml(@"<references count=""1""><reference field=""4294967294"" count=""1"" selected=""0""><x v=""0""/></reference></references>");
			var node = document.SelectSingleNode("//reference");
			new AutoSortScopeReference(TestUtility.CreateDefaultNSM(), node);
		}
		#endregion
	}
}
