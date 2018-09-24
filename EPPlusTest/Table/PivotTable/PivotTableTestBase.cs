using System.Xml;
using OfficeOpenXml.Table.PivotTable;

namespace EPPlusTest.Table.PivotTable
{
	public abstract class PivotTableTestBase
	{
		#region Helper Methods
		public CacheFieldNode GetTestCacheFieldNode()
		{
			XmlDocument document = new XmlDocument();
			document.LoadXml(@"<cacheField xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" name=""Item"" numFmtId=""0""><sharedItems count=""2""><s v=""Bike""/><s v=""Car""/></sharedItems></cacheField>");
			var ns = TestUtility.CreateDefaultNSM();
			return new CacheFieldNode(document.SelectSingleNode("//d:cacheField", ns), ns);
		}

		public CacheRecordNode GetTestCacheRecordNode()
		{
			XmlDocument document = new XmlDocument();
			document.LoadXml(@"<pivotCacheRecords xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" count=""1""><r><n v=""20100076""/><x v=""0""/></r></pivotCacheRecords>");
			var ns = TestUtility.CreateDefaultNSM();
			return new CacheRecordNode(document.SelectSingleNode("//d:r", ns), ns);
		}
		#endregion
	}
}
