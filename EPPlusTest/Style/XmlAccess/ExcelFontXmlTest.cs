using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Style.XmlAccess;

namespace EPPlusTest.Style.XmlAccess
{
	[TestClass]
	public class ExcelFontXmlTest
	{
		#region Class Constants
		const string TestFont = "Rockwell";
		#endregion

		#region GetFontHeight Tests
		[TestMethod]
		public void GetFontHeight()
		{
			var height = ExcelFontXml.GetFontHeight(ExcelFontXmlTest.TestFont, 8);
			Assert.AreEqual(20, height);
		}

		[TestMethod]
		public void GetFontHeightHandlesAtSymbol()
		{
			var height = ExcelFontXml.GetFontHeight($"@{ExcelFontXmlTest.TestFont}", 8);
			Assert.AreEqual(20, height);
		}

		[TestMethod]
		public void GetFontHeightFakeFontDefaultsToCalibri()
		{
			var defaultedHeight = ExcelFontXml.GetFontHeight("whatever", 8);
			var calibriHeight = ExcelFontXml.GetFontHeight("Calibri", 8);
			Assert.AreEqual(calibriHeight, defaultedHeight);
		}

		[TestMethod]
		public void GetFontHeightTinyFontSizeDefaultsToMinimumKnownSize()
		{
			var defaultedHeight = ExcelFontXml.GetFontHeight("Calibri", 2);
			var calibriHeight = ExcelFontXml.GetFontHeight("Calibri", 6);
			Assert.AreEqual(calibriHeight, defaultedHeight);
		}

		[TestMethod]
		public void GetFontHeightHugeFontSizeDefaultsToMaximumKnownSize()
		{
			var defaultedHeight = ExcelFontXml.GetFontHeight("Calibri", 5000);
			var calibriHeight = ExcelFontXml.GetFontHeight("Calibri", 256);
			Assert.AreEqual(calibriHeight, defaultedHeight);
		}
		#endregion
	}
}
