using System;
using System.Xml;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Table.PivotTable.Formats;

namespace EPPlusTest.Table.PivotTable.Formats
{
	[TestClass]
	public class PivotAreaCollectionTest
	{
		#region Constructor Tests
		[TestMethod]
		public void PivotAreaCollectionConstructorTest()
		{
			var xmlDoc = new XmlDocument(TestUtility.CreateDefaultNSM().NameTable);
			xmlDoc.LoadXml(@"
				<pivotAreas count=""2"" xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
					<pivotArea type=""data"" outline=""0"" collapsedLevelsAreSubtotals=""1"" fieldPosition=""0"">
						<references count=""2"">
							<reference field=""4294967294"" count=""1"" selected=""0"">
								<x v=""1""/>
							</reference>
							<reference field=""1"" count=""1"" selected=""0"">
								<x v=""0""/>
							</reference>
						</references>
					</pivotArea>
					<pivotArea type=""data"" outline=""0"" collapsedLevelsAreSubtotals=""1"" fieldPosition=""0"">
						<references count=""4"">
							<reference field=""4294967294"" count=""1"" selected=""0"">
								<x v=""1""/>
							</reference>
							<reference field=""1"" count=""1"" selected=""0"">
								<x v=""1""/>
							</reference>
							<reference field=""2"" count=""1"" selected=""0"">
								<x v=""0""/>
							</reference>
							<reference field=""3"" count=""1"" selected=""0"">
								<x v=""0""/>
							</reference>
						</references>
					</pivotArea>
				</pivotAreas>");
			var collection = new PivotAreasCollection(TestUtility.CreateDefaultNSM(), xmlDoc.FirstChild);
			Assert.IsNotNull(collection);
			Assert.AreEqual(2, collection.Count);
			Assert.AreEqual(PivotAreaType.Data, collection[0].RuleType);
			Assert.IsFalse(collection[0].Outline);
			Assert.IsTrue(collection[0].CollapsedLevelsAreSubtotals);
			Assert.AreEqual(PivotAreaType.Data, collection[1].RuleType);
			Assert.IsFalse(collection[1].Outline);
			Assert.IsTrue(collection[1].CollapsedLevelsAreSubtotals);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void PivotAreaCollectionNullNamespaceManagerTest()
		{
			var xmlDoc = new XmlDocument(TestUtility.CreateDefaultNSM().NameTable);
			xmlDoc.LoadXml(@"
				<pivotAreas count=""2"" xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">
					<pivotArea type=""data"" outline=""0"" collapsedLevelsAreSubtotals=""1"" fieldPosition=""0"">
						<references count=""2"">
							<reference field=""4294967294"" count=""1"" selected=""0"">
								<x v=""1""/>
							</reference>
							<reference field=""1"" count=""1"" selected=""0"">
								<x v=""0""/>
							</reference>
						</references>
					</pivotArea>
					<pivotArea type=""data"" outline=""0"" collapsedLevelsAreSubtotals=""1"" fieldPosition=""0"">
						<references count=""4"">
							<reference field=""4294967294"" count=""1"" selected=""0"">
								<x v=""1""/>
							</reference>
							<reference field=""1"" count=""1"" selected=""0"">
								<x v=""1""/>
							</reference>
							<reference field=""2"" count=""1"" selected=""0"">
								<x v=""0""/>
							</reference>
							<reference field=""3"" count=""1"" selected=""0"">
								<x v=""0""/>
							</reference>
						</references>
					</pivotArea>
				</pivotAreas>");
			new PivotAreasCollection(null, xmlDoc.FirstChild);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void PivotAreaCollectionNullNodeTest()
		{
			new PivotAreasCollection(TestUtility.CreateDefaultNSM(), null);
		}
		#endregion
	}
}
