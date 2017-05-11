using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using System;

namespace EPPlusTest.FormulaParsing
{
	[TestClass]
	public class EpplusExcelDataProviderTest
	{
		#region GetFormat Tests
		[TestMethod]
		public void GetFormatBuiltinFormat()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				var provider = new EpplusExcelDataProvider(package);
				Assert.AreEqual("31-Dec-17", provider.GetFormat(new DateTime(2017, 12, 31), "d-mmm-yy"));
			}
		}

		[TestMethod]
		public void GetFormatCustomFormat()
		{
			using (var package = new ExcelPackage())
			{
				var sheet = package.Workbook.Worksheets.Add("Sheet");
				var provider = new EpplusExcelDataProvider(package);
				Assert.AreEqual("12312017", provider.GetFormat(new DateTime(2017, 12, 31), "mmddyyyy"));
			}
		}
		#endregion
	}
}
