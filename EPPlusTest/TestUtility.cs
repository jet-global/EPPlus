using System.Xml;
using OfficeOpenXml;

namespace EPPlusTest
{
	internal static class TestUtility
	{
		public static XmlNamespaceManager CreateDefaultNSM()
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
	}
}