using System;
using System.Xml;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Table.PivotTable.Formats;

namespace EPPlusTest.Table.PivotTable.Formats
{
	[TestClass]
	public class PivotAreaTest
	{
		[TestMethod]
		public void PivotAreaConstructorTest()
		{
			var xmlDoc = new XmlDocument(TestUtility.CreateDefaultNSM().NameTable);
			xmlDoc.LoadXml(@"
					<pivotArea type=""data"" outline=""0"" collapsedLevelsAreSubtotals=""1"" fieldPosition=""0"" xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
						<references count=""2"">
							<reference field=""4294967294"" count=""1"" selected=""0"">
								<x v=""1""/>
							</reference>
							<reference field=""1"" count=""1"" selected=""0"">
								<x v=""0""/>
							</reference>
						</references>
					</pivotArea>");
			var pivotArea = new PivotArea(TestUtility.CreateDefaultNSM(), xmlDoc.FirstChild);
			Assert.IsNotNull(pivotArea);
			Assert.AreEqual(PivotAreaType.Data, pivotArea.RuleType);
			Assert.IsFalse(pivotArea.Outline);
			Assert.IsTrue(pivotArea.CollapsedLevelsAreSubtotals);
			Assert.AreEqual(2, pivotArea.ReferencesCollection.Count);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void PivotAreaNullNamespaceManagerTest()
		{
			var xmlDoc = new XmlDocument(TestUtility.CreateDefaultNSM().NameTable);
			xmlDoc.LoadXml(@"
					<pivotArea type=""data"" outline=""0"" collapsedLevelsAreSubtotals=""1"" fieldPosition=""0"" xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
						<references count=""2"">
							<reference field=""4294967294"" count=""1"" selected=""0"">
								<x v=""1""/>
							</reference>
							<reference field=""1"" count=""1"" selected=""0"">
								<x v=""0""/>
							</reference>
						</references>
					</pivotArea>");
			new PivotArea(null, xmlDoc.FirstChild);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void PivotAreaNullNodeTest()
		{
			new PivotArea(TestUtility.CreateDefaultNSM(), null);
		}
	}
}
