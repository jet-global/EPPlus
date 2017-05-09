using System.Xml;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Style.Dxf;

namespace EPPlusTest.Style.Dxf
{
	[TestClass]
	public class ExcelDxfStyleTest
	{
		#region Font Tests
		[TestMethod]
		public void FontDefaults()
		{
			var nodeXml = "<dxf xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
			"<font>" +
				"<b/>" +
				"<i/>" +
				"<color theme=\"0\"/>" +
			"</font>" +
			"<fill>" +
				"<patternFill>" +
					"<bgColor rgb=\"FF002060\"/>" +
				"</patternFill>" +
			"</fill>" +
		"</dxf>";
			using (var excelPackage = new ExcelPackage())
			{
				var document = new XmlDocument();
				document.LoadXml(nodeXml);
				XmlNode dxfNode = document.DocumentElement;
				var namespaceManager = new XmlNamespaceManager(document.NameTable);
				var excelDxfStyle = new ExcelDxfStyleConditionalFormatting(excelPackage.Workbook.NameSpaceManager, dxfNode, excelPackage.Workbook.Styles);
				Assert.IsTrue(excelDxfStyle.Font.Bold.Value);
				Assert.IsTrue(excelDxfStyle.Font.Italic.Value);
			}
		}

		[TestMethod]
		public void FontNotBold()
		{
			var nodeXml = "<dxf xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
			"<font>" +
				"<b val=\"0\"/>" +
				"<i/>" +
				"<color theme=\"0\"/>" +
			"</font>" +
			"<fill>" +
				"<patternFill>" +
					"<bgColor rgb=\"FF002060\"/>" +
				"</patternFill>" +
			"</fill>" +
		"</dxf>";
			using (var excelPackage = new ExcelPackage())
			{
				var document = new XmlDocument();
				document.LoadXml(nodeXml);
				XmlNode dxfNode = document.DocumentElement;
				var namespaceManager = new XmlNamespaceManager(document.NameTable);
				var excelDxfStyle = new ExcelDxfStyleConditionalFormatting(excelPackage.Workbook.NameSpaceManager, dxfNode, excelPackage.Workbook.Styles);
				Assert.IsFalse(excelDxfStyle.Font.Bold.Value);
				Assert.IsTrue(excelDxfStyle.Font.Italic.Value);
			}
		}

		[TestMethod]
		public void FontNotItalic()
		{
			var nodeXml = "<dxf xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
			"<font>" +
				"<b/>" +
				"<i val=\"0\"/>" +
				"<color theme=\"0\"/>" +
			"</font>" +
			"<fill>" +
				"<patternFill>" +
					"<bgColor rgb=\"FF002060\"/>" +
				"</patternFill>" +
			"</fill>" +
		"</dxf>";
			using (var excelPackage = new ExcelPackage())
			{
				var document = new XmlDocument();
				document.LoadXml(nodeXml);
				XmlNode dxfNode = document.DocumentElement;
				var namespaceManager = new XmlNamespaceManager(document.NameTable);
				var excelDxfStyle = new ExcelDxfStyleConditionalFormatting(excelPackage.Workbook.NameSpaceManager, dxfNode, excelPackage.Workbook.Styles);
				Assert.IsTrue(excelDxfStyle.Font.Bold.Value);
				Assert.IsFalse(excelDxfStyle.Font.Italic.Value);
			}
		}

		[TestMethod]
		public void FontNotBoldAndNotItalic()
		{
			var nodeXml = "<dxf xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
			"<font>" +
				"<b val=\"0\"/>" +
				"<i val=\"0\"/>" +
				"<color theme=\"0\"/>" +
			"</font>" +
			"<fill>" +
				"<patternFill>" +
					"<bgColor rgb=\"FF002060\"/>" +
				"</patternFill>" +
			"</fill>" +
		"</dxf>";
			using (var excelPackage = new ExcelPackage())
			{
				var document = new XmlDocument();
				document.LoadXml(nodeXml);
				XmlNode dxfNode = document.DocumentElement;
				var namespaceManager = new XmlNamespaceManager(document.NameTable);
				var excelDxfStyle = new ExcelDxfStyleConditionalFormatting(excelPackage.Workbook.NameSpaceManager, dxfNode, excelPackage.Workbook.Styles);
				Assert.IsFalse(excelDxfStyle.Font.Bold.Value);
				Assert.IsFalse(excelDxfStyle.Font.Italic.Value);
			}
		}
		#endregion
	}
}
