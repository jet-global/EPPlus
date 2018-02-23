using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;

namespace EPPlusTest.FormulaParsing
{
	[TestClass]
	public class EpplusExcelDataProviderTest
	{
		#region Constants
		private const string TableName = "MyTable";
		private const string Header1 = "Header1";
		private const string Header2 = "Header2";
		private const string Header3 = "Header3";
		private const string Header4 = "Header4";
		#endregion

		#region Properties
		private EpplusExcelDataProvider DataProviderWithDataAndHeaders { get; set; }
		private EpplusExcelDataProvider DataProviderWithDataAndTotals { get; set; }
		private EpplusExcelDataProvider DataProviderWithDataHeadersAndTotals { get; set; }
		#endregion

		#region Test Setup
		[TestInitialize]
		public void SetUp()
		{
			var package = new ExcelPackage();
			var worksheet = package.Workbook.Worksheets.Add("Sheet1");
			this.BuildTableHeaders(worksheet);
			this.BuildTableData(worksheet);
			var table = worksheet.Tables.Add(new ExcelAddress("Sheet1", 3, 3, 9, 6), EpplusExcelDataProviderTest.TableName);
			table.ShowHeader = true;
			table.ShowTotal = false;
			Assert.AreEqual("C3:F9", table.Address.Address);
			this.DataProviderWithDataAndHeaders = new EpplusExcelDataProvider(package);

			package = new ExcelPackage();
			worksheet = package.Workbook.Worksheets.Add("Sheet1");
			this.BuildTableData(worksheet);
			this.BuildTableTotals(worksheet);
			table = worksheet.Tables.Add(new ExcelAddress("Sheet1", 4, 3, 9, 6), EpplusExcelDataProviderTest.TableName);
			table.ShowHeader = false;
			// Note: This adds a row to the table's address
			table.ShowTotal = true;
			Assert.AreEqual("C4:F10", table.Address.Address);
			this.DataProviderWithDataAndTotals = new EpplusExcelDataProvider(package);

			package = new ExcelPackage();
			worksheet = package.Workbook.Worksheets.Add("Sheet1");
			this.BuildTableHeaders(worksheet);
			this.BuildTableData(worksheet);
			this.BuildTableTotals(worksheet);
			table = worksheet.Tables.Add(new ExcelAddress("Sheet1", 3, 3, 9, 6), EpplusExcelDataProviderTest.TableName);
			table.ShowHeader = true;
			// Note: This adds a row to the table's address
			table.ShowTotal = true;
			Assert.AreEqual("C3:F10", table.Address.Address);
			this.DataProviderWithDataHeadersAndTotals = new EpplusExcelDataProvider(package);
		}

		[TestCleanup]
		public void Cleanup()
		{
			this.DataProviderWithDataAndHeaders.Dispose();
			this.DataProviderWithDataAndTotals.Dispose();
		}
		#endregion

		#region ResolveStructuredReference Tests
		#region #ThisRow Tests
		[TestMethod]
		public void ResolveStructuredReferenceThisRowRightOfTable()
		{
			var reference = new StructuredReference($"{EpplusExcelDataProviderTest.TableName}[[#This Row],[{EpplusExcelDataProviderTest.Header1}]]");
			var result = this.DataProviderWithDataAndHeaders.ResolveStructuredReference(reference, "Sheet1", 7, 7);
			Assert.AreEqual(7, result.Address.Start.Row);
			Assert.AreEqual(3, result.Address.Start.Column);
			Assert.AreEqual(7, result.Address.End.Row);
			Assert.AreEqual(3, result.Address.End.Column);
			Assert.AreEqual("h1_r4", result.GetOffset(0, 0));
		}

		[TestMethod]
		public void ResolveStructuredReferenceThisRowRange()
		{
			var reference = new StructuredReference($"{EpplusExcelDataProviderTest.TableName}[[#This Row],[{EpplusExcelDataProviderTest.Header1}]:[{EpplusExcelDataProviderTest.Header3}]]");
			var result = this.DataProviderWithDataAndHeaders.ResolveStructuredReference(reference, "Sheet1", 7, 7);
			Assert.AreEqual(7, result.Address.Start.Row);
			Assert.AreEqual(3, result.Address.Start.Column);
			Assert.AreEqual(7, result.Address.End.Row);
			Assert.AreEqual(5, result.Address.End.Column);
			Assert.AreEqual("h1_r4", result.GetOffset(0, 0));
			Assert.AreEqual("h2_r4", result.GetOffset(0, 1));
			Assert.AreEqual("h3_r4", result.GetOffset(0, 2));
		}

		[TestMethod]
		public void ResolveStructuredReferenceThisRowLeftOfTable()
		{
			var reference = new StructuredReference($"{EpplusExcelDataProviderTest.TableName}[[#This Row],[{EpplusExcelDataProviderTest.Header1}]]");
			var result = this.DataProviderWithDataAndHeaders.ResolveStructuredReference(reference, "Sheet1", 7, 1);
			Assert.AreEqual(7, result.Address.Start.Row);
			Assert.AreEqual(3, result.Address.Start.Column);
			Assert.AreEqual(7, result.Address.End.Row);
			Assert.AreEqual(3, result.Address.End.Column);
			Assert.AreEqual("h1_r4", result.GetOffset(0, 0));
		}

		[TestMethod]
		public void ResolveStructuredReferenceThisRowAboveTable()
		{
			var reference = new StructuredReference($"{EpplusExcelDataProviderTest.TableName}[[#This Row],[{EpplusExcelDataProviderTest.Header1}]]");
			var result = this.DataProviderWithDataAndHeaders.ResolveStructuredReference(reference, "Sheet1", 2, 4);
			Assert.IsNull(result);
		}

		[TestMethod]
		public void ResolveStructuredReferenceThisRowBelowTable()
		{
			var reference = new StructuredReference($"{EpplusExcelDataProviderTest.TableName}[[#This Row],[{EpplusExcelDataProviderTest.Header1}]]");
			var result = this.DataProviderWithDataAndHeaders.ResolveStructuredReference(reference, "Sheet1", 12, 4);
			Assert.IsNull(result);
		}

		[TestMethod]
		public void ResolveStructuredReferenceThisRowHeaderRow()
		{
			var reference = new StructuredReference($"{EpplusExcelDataProviderTest.TableName}[[#This Row],[{EpplusExcelDataProviderTest.Header1}]]");
			var result = this.DataProviderWithDataAndHeaders.ResolveStructuredReference(reference, "Sheet1", 3, 7);
			Assert.IsNull(result);
		}
		#endregion

		#region #Data Tests
		[TestMethod]
		public void ResolveStructuredReferenceDataIgnoresHeaders()
		{
			var reference = new StructuredReference($"{EpplusExcelDataProviderTest.TableName}[[#Data],[{EpplusExcelDataProviderTest.Header4}]]");
			var result = this.DataProviderWithDataAndHeaders.ResolveStructuredReference(reference, "Sheet1", 15, 7);
			Assert.AreEqual(4, result.Address.Start.Row);
			Assert.AreEqual(6, result.Address.Start.Column);
			Assert.AreEqual(9, result.Address.End.Row);
			Assert.AreEqual(6, result.Address.End.Column);
			Assert.AreEqual("h4_r1", result.GetOffset(0, 0));
			Assert.AreEqual("h4_r2", result.GetOffset(1, 0));
			Assert.AreEqual("h4_r3", result.GetOffset(2, 0));
			Assert.AreEqual("h4_r4", result.GetOffset(3, 0));
			Assert.AreEqual("h4_r5", result.GetOffset(4, 0));
			Assert.AreEqual("h4_r6", result.GetOffset(5, 0));
		}

		[TestMethod]
		public void ResolveStructuredReferenceDataIgnoresTotals()
		{
			var reference = new StructuredReference($"{EpplusExcelDataProviderTest.TableName}[[#Data]]");
			var result = this.DataProviderWithDataAndTotals.ResolveStructuredReference(reference, "Sheet1", 15, 7);
			Assert.AreEqual(4, result.Address.Start.Row);
			Assert.AreEqual(3, result.Address.Start.Column);
			Assert.AreEqual(9, result.Address.End.Row);
			Assert.AreEqual(6, result.Address.End.Column);
			Assert.AreEqual("h1_r1", result.GetOffset(0, 0));
			Assert.AreEqual("h1_r2", result.GetOffset(1, 0));
			Assert.AreEqual("h1_r3", result.GetOffset(2, 0));
			Assert.AreEqual("h1_r4", result.GetOffset(3, 0));
			Assert.AreEqual("h1_r5", result.GetOffset(4, 0));
			Assert.AreEqual("h1_r6", result.GetOffset(5, 0));
		}

		[TestMethod]
		public void ResolveStructuredReferenceDataMultipleColumns()
		{
			var reference = new StructuredReference($"{EpplusExcelDataProviderTest.TableName}[[#Data],[{EpplusExcelDataProviderTest.Header3}]:[{EpplusExcelDataProviderTest.Header4}]]");
			var result = this.DataProviderWithDataAndHeaders.ResolveStructuredReference(reference, "Sheet1", 15, 7);
			Assert.AreEqual(4, result.Address.Start.Row);
			Assert.AreEqual(5, result.Address.Start.Column);
			Assert.AreEqual(9, result.Address.End.Row);
			Assert.AreEqual(6, result.Address.End.Column);
			Assert.AreEqual("h3_r1", result.GetOffset(0, 0));
			Assert.AreEqual("h3_r2", result.GetOffset(1, 0));
			Assert.AreEqual("h3_r3", result.GetOffset(2, 0));
			Assert.AreEqual("h3_r4", result.GetOffset(3, 0));
			Assert.AreEqual("h3_r5", result.GetOffset(4, 0));
			Assert.AreEqual("h3_r6", result.GetOffset(5, 0));
			Assert.AreEqual("h4_r1", result.GetOffset(0, 1));
			Assert.AreEqual("h4_r2", result.GetOffset(1, 1));
			Assert.AreEqual("h4_r3", result.GetOffset(2, 1));
			Assert.AreEqual("h4_r4", result.GetOffset(3, 1));
			Assert.AreEqual("h4_r5", result.GetOffset(4, 1));
			Assert.AreEqual("h4_r6", result.GetOffset(5, 1));
		}

		[TestMethod]
		public void ResolveStructuredReferenceDataEntireTable()
		{
			var reference = new StructuredReference($"{EpplusExcelDataProviderTest.TableName}[#Data]");
			var result = this.DataProviderWithDataAndHeaders.ResolveStructuredReference(reference, "Sheet1", 15, 7);
			Assert.AreEqual(4, result.Address.Start.Row);
			Assert.AreEqual(3, result.Address.Start.Column);
			Assert.AreEqual(9, result.Address.End.Row);
			Assert.AreEqual(6, result.Address.End.Column);
			Assert.AreEqual("h1_r1", result.GetOffset(0, 0));
			Assert.AreEqual("h1_r2", result.GetOffset(1, 0));
			Assert.AreEqual("h1_r3", result.GetOffset(2, 0));
			Assert.AreEqual("h1_r4", result.GetOffset(3, 0));
			Assert.AreEqual("h1_r5", result.GetOffset(4, 0));
			Assert.AreEqual("h1_r6", result.GetOffset(5, 0));
			Assert.AreEqual("h4_r6", result.GetOffset(5, 3));
		}
		#endregion

		#region Data and Headers Tests
		[TestMethod]
		public void ResolveStructuredReferenceDataWithHeadersThatExist()
		{
			var reference = new StructuredReference($"{EpplusExcelDataProviderTest.TableName}[[#Data],[#Headers]]");
			var result = this.DataProviderWithDataAndHeaders.ResolveStructuredReference(reference, "Sheet1", 15, 7);
			Assert.AreEqual(3, result.Address.Start.Row);
			Assert.AreEqual(3, result.Address.Start.Column);
			Assert.AreEqual(9, result.Address.End.Row);
			Assert.AreEqual(6, result.Address.End.Column);
			Assert.AreEqual(EpplusExcelDataProviderTest.Header1, result.GetOffset(0, 0));
			Assert.AreEqual(EpplusExcelDataProviderTest.Header2, result.GetOffset(0, 1));
			Assert.AreEqual(EpplusExcelDataProviderTest.Header3, result.GetOffset(0, 2));
			Assert.AreEqual(EpplusExcelDataProviderTest.Header4, result.GetOffset(0, 3));
			Assert.AreEqual("h1_r1", result.GetOffset(1, 0));
			Assert.AreEqual("h1_r2", result.GetOffset(2, 0));
			Assert.AreEqual("h1_r3", result.GetOffset(3, 0));
			Assert.AreEqual("h1_r4", result.GetOffset(4, 0));
			Assert.AreEqual("h1_r5", result.GetOffset(5, 0));
			Assert.AreEqual("h1_r6", result.GetOffset(6, 0));
		}

		[TestMethod]
		public void ResolveStructuredReferenceDataWithHeadersThatDoNotExist()
		{
			var reference = new StructuredReference($"{EpplusExcelDataProviderTest.TableName}[[#Data],[#Headers]]");
			var result = this.DataProviderWithDataAndTotals.ResolveStructuredReference(reference, "Sheet1", 15, 7);
			Assert.AreEqual(4, result.Address.Start.Row);
			Assert.AreEqual(3, result.Address.Start.Column);
			Assert.AreEqual(9, result.Address.End.Row);
			Assert.AreEqual(6, result.Address.End.Column);
			Assert.AreEqual("h1_r1", result.GetOffset(0, 0));
			Assert.AreEqual("h1_r2", result.GetOffset(1, 0));
			Assert.AreEqual("h1_r3", result.GetOffset(2, 0));
			Assert.AreEqual("h1_r4", result.GetOffset(3, 0));
			Assert.AreEqual("h1_r5", result.GetOffset(4, 0));
			Assert.AreEqual("h1_r6", result.GetOffset(5, 0));
			Assert.AreEqual("h4_r6", result.GetOffset(5, 3));
		}

		[TestMethod]
		public void ResolveStructuredReferenceDataWithHeadersColumnSpecified()
		{
			var reference = new StructuredReference($"{EpplusExcelDataProviderTest.TableName}[[#Data],[#Headers],[Header2]]");
			var result = this.DataProviderWithDataAndHeaders.ResolveStructuredReference(reference, "Sheet1", 15, 7);
			Assert.AreEqual(3, result.Address.Start.Row);
			Assert.AreEqual(4, result.Address.Start.Column);
			Assert.AreEqual(9, result.Address.End.Row);
			Assert.AreEqual(4, result.Address.End.Column);
			Assert.AreEqual(EpplusExcelDataProviderTest.Header2, result.GetOffset(0, 0));
			Assert.AreEqual("h2_r1", result.GetOffset(1, 0));
			Assert.AreEqual("h2_r2", result.GetOffset(2, 0));
			Assert.AreEqual("h2_r3", result.GetOffset(3, 0));
			Assert.AreEqual("h2_r4", result.GetOffset(4, 0));
			Assert.AreEqual("h2_r5", result.GetOffset(5, 0));
			Assert.AreEqual("h2_r6", result.GetOffset(6, 0));
		}

		[TestMethod]
		public void ResolveStructuredReferenceDataWithHeadersColumnRangeSpecified()
		{
			var reference = new StructuredReference($"{EpplusExcelDataProviderTest.TableName}[[#Data],[#Headers],[Header2]:[Header4]]");
			var result = this.DataProviderWithDataAndHeaders.ResolveStructuredReference(reference, "Sheet1", 15, 7);
			Assert.AreEqual(3, result.Address.Start.Row);
			Assert.AreEqual(4, result.Address.Start.Column);
			Assert.AreEqual(9, result.Address.End.Row);
			Assert.AreEqual(6, result.Address.End.Column);
			Assert.AreEqual(EpplusExcelDataProviderTest.Header2, result.GetOffset(0, 0));
			Assert.AreEqual(EpplusExcelDataProviderTest.Header3, result.GetOffset(0, 1));
			Assert.AreEqual(EpplusExcelDataProviderTest.Header4, result.GetOffset(0, 2));
			Assert.AreEqual("h2_r1", result.GetOffset(1, 0));
			Assert.AreEqual("h2_r2", result.GetOffset(2, 0));
			Assert.AreEqual("h2_r3", result.GetOffset(3, 0));
			Assert.AreEqual("h2_r4", result.GetOffset(4, 0));
			Assert.AreEqual("h2_r5", result.GetOffset(5, 0));
			Assert.AreEqual("h2_r6", result.GetOffset(6, 0));
			Assert.AreEqual("h4_r6", result.GetOffset(6, 2));
		}
		#endregion

		#region Data and Totals Tests
		[TestMethod]
		public void ResolveStructuredReferenceDataWithTotals()
		{
			var reference = new StructuredReference($"{EpplusExcelDataProviderTest.TableName}[[#Data],[#Totals]]");
			var result = this.DataProviderWithDataAndTotals.ResolveStructuredReference(reference, "Sheet1", 15, 7);
			Assert.AreEqual(4, result.Address.Start.Row);
			Assert.AreEqual(3, result.Address.Start.Column);
			Assert.AreEqual(10, result.Address.End.Row);
			Assert.AreEqual(6, result.Address.End.Column);
			Assert.AreEqual("h1_r1", result.GetOffset(0, 0));
			Assert.AreEqual("h1_r2", result.GetOffset(1, 0));
			Assert.AreEqual("h1_r3", result.GetOffset(2, 0));
			Assert.AreEqual("h1_r4", result.GetOffset(3, 0));
			Assert.AreEqual("h1_r5", result.GetOffset(4, 0));
			Assert.AreEqual("h1_r6", result.GetOffset(5, 0));
			Assert.AreEqual("h1_t", result.GetOffset(6, 0));
			Assert.AreEqual("h4_t", result.GetOffset(6, 3));
		}

		[TestMethod]
		public void ResolveStructuredReferenceDataWithTotalsColumnSpecified()
		{
			var reference = new StructuredReference($"{EpplusExcelDataProviderTest.TableName}[[#Data],[#Totals],[Header2]]");
			var result = this.DataProviderWithDataHeadersAndTotals.ResolveStructuredReference(reference, "Sheet1", 15, 7);
			Assert.AreEqual(4, result.Address.Start.Row);
			Assert.AreEqual(4, result.Address.Start.Column);
			Assert.AreEqual(10, result.Address.End.Row);
			Assert.AreEqual(4, result.Address.End.Column);
			Assert.AreEqual("h2_r1", result.GetOffset(0, 0));
			Assert.AreEqual("h2_r2", result.GetOffset(1, 0));
			Assert.AreEqual("h2_r3", result.GetOffset(2, 0));
			Assert.AreEqual("h2_r4", result.GetOffset(3, 0));
			Assert.AreEqual("h2_r5", result.GetOffset(4, 0));
			Assert.AreEqual("h2_r6", result.GetOffset(5, 0));
			Assert.AreEqual("h2_t", result.GetOffset(6, 0));
		}

		[TestMethod]
		public void ResolveStructuredReferenceDataWithTotalsColumnRangeSpecified()
		{
			var reference = new StructuredReference($"{EpplusExcelDataProviderTest.TableName}[[#Data],[#Totals],[Header2]:[Header3]");
			var result = this.DataProviderWithDataHeadersAndTotals.ResolveStructuredReference(reference, "Sheet1", 15, 7);
			Assert.AreEqual(4, result.Address.Start.Row);
			Assert.AreEqual(4, result.Address.Start.Column);
			Assert.AreEqual(10, result.Address.End.Row);
			Assert.AreEqual(5, result.Address.End.Column);
			Assert.AreEqual("h2_r1", result.GetOffset(0, 0));
			Assert.AreEqual("h2_r2", result.GetOffset(1, 0));
			Assert.AreEqual("h2_r3", result.GetOffset(2, 0));
			Assert.AreEqual("h2_r4", result.GetOffset(3, 0));
			Assert.AreEqual("h2_r5", result.GetOffset(4, 0));
			Assert.AreEqual("h2_r6", result.GetOffset(5, 0));
			Assert.AreEqual("h2_t", result.GetOffset(6, 0));
			Assert.AreEqual("h3_r1", result.GetOffset(0, 1));
			Assert.AreEqual("h3_r2", result.GetOffset(1, 1));
			Assert.AreEqual("h3_r3", result.GetOffset(2, 1));
			Assert.AreEqual("h3_r4", result.GetOffset(3, 1));
			Assert.AreEqual("h3_r5", result.GetOffset(4, 1));
			Assert.AreEqual("h3_r6", result.GetOffset(5, 1));
			Assert.AreEqual("h3_t", result.GetOffset(6, 1));
		}

		public void ResolveStructuredReferenceDataWithTotalsThatDoNotExist()
		{
			var reference = new StructuredReference($"{EpplusExcelDataProviderTest.TableName}[[#Data],[#Totals]]");
			var result = this.DataProviderWithDataAndHeaders.ResolveStructuredReference(reference, "Sheet1", 15, 7);
			Assert.AreEqual(4, result.Address.Start.Row);
			Assert.AreEqual(3, result.Address.Start.Column);
			Assert.AreEqual(9, result.Address.End.Row);
			Assert.AreEqual(6, result.Address.End.Column);
			Assert.AreEqual("h1_r1", result.GetOffset(0, 0));
			Assert.AreEqual("h1_r2", result.GetOffset(1, 0));
			Assert.AreEqual("h1_r3", result.GetOffset(2, 0));
			Assert.AreEqual("h1_r4", result.GetOffset(3, 0));
			Assert.AreEqual("h1_r5", result.GetOffset(4, 0));
			Assert.AreEqual("h1_r6", result.GetOffset(5, 0));
		}

		[TestMethod]
		public void ResolveStructuredReferenceDataWithTotalsThatDoNotExistColumnSpecified()
		{
			var reference = new StructuredReference($"{EpplusExcelDataProviderTest.TableName}[[#Data],[#Totals],[Header2]]");
			var result = this.DataProviderWithDataAndHeaders.ResolveStructuredReference(reference, "Sheet1", 15, 7);
			Assert.AreEqual(4, result.Address.Start.Row);
			Assert.AreEqual(4, result.Address.Start.Column);
			Assert.AreEqual(9, result.Address.End.Row);
			Assert.AreEqual(4, result.Address.End.Column);
			Assert.AreEqual("h2_r1", result.GetOffset(0, 0));
			Assert.AreEqual("h2_r2", result.GetOffset(1, 0));
			Assert.AreEqual("h2_r3", result.GetOffset(2, 0));
			Assert.AreEqual("h2_r4", result.GetOffset(3, 0));
			Assert.AreEqual("h2_r5", result.GetOffset(4, 0));
			Assert.AreEqual("h2_r6", result.GetOffset(5, 0));
		}

		[TestMethod]
		public void ResolveStructuredReferenceDataWithTotalsThatDoNotExistColumnRangeSpecified()
		{
			var reference = new StructuredReference($"{EpplusExcelDataProviderTest.TableName}[[#Data],[#Totals],[Header2]:[Header3]");
			var result = this.DataProviderWithDataAndHeaders.ResolveStructuredReference(reference, "Sheet1", 15, 7);
			Assert.AreEqual(4, result.Address.Start.Row);
			Assert.AreEqual(4, result.Address.Start.Column);
			Assert.AreEqual(9, result.Address.End.Row);
			Assert.AreEqual(5, result.Address.End.Column);
			Assert.AreEqual("h2_r1", result.GetOffset(0, 0));
			Assert.AreEqual("h2_r2", result.GetOffset(1, 0));
			Assert.AreEqual("h2_r3", result.GetOffset(2, 0));
			Assert.AreEqual("h2_r4", result.GetOffset(3, 0));
			Assert.AreEqual("h2_r5", result.GetOffset(4, 0));
			Assert.AreEqual("h2_r6", result.GetOffset(5, 0));
			Assert.AreEqual("h3_r1", result.GetOffset(0, 1));
			Assert.AreEqual("h3_r2", result.GetOffset(1, 1));
			Assert.AreEqual("h3_r3", result.GetOffset(2, 1));
			Assert.AreEqual("h3_r4", result.GetOffset(3, 1));
			Assert.AreEqual("h3_r5", result.GetOffset(4, 1));
			Assert.AreEqual("h3_r6", result.GetOffset(5, 1));
		}
		#endregion

		#region Headers Tests
		[TestMethod]
		public void ResolveStructuredReferenceHeadersThatExist()
		{
			var reference = new StructuredReference($"{EpplusExcelDataProviderTest.TableName}[#Headers]");
			var result = this.DataProviderWithDataAndHeaders.ResolveStructuredReference(reference, "Sheet1", 15, 7);
			Assert.AreEqual(3, result.Address.Start.Row);
			Assert.AreEqual(3, result.Address.Start.Column);
			Assert.AreEqual(3, result.Address.End.Row);
			Assert.AreEqual(6, result.Address.End.Column);
			Assert.AreEqual(EpplusExcelDataProviderTest.Header1, result.GetOffset(0, 0));
			Assert.AreEqual(EpplusExcelDataProviderTest.Header2, result.GetOffset(0, 1));
			Assert.AreEqual(EpplusExcelDataProviderTest.Header3, result.GetOffset(0, 2));
			Assert.AreEqual(EpplusExcelDataProviderTest.Header4, result.GetOffset(0, 3));
		}

		[TestMethod]
		public void ResolveStructuredReferenceHeadersThatDoNotExist()
		{
			var reference = new StructuredReference($"{EpplusExcelDataProviderTest.TableName}[#Headers]");
			var result = this.DataProviderWithDataAndTotals.ResolveStructuredReference(reference, "Sheet1", 15, 7);
			Assert.AreEqual("#REF!", result.Address.Address);
		}

		[TestMethod]
		public void ResolveStructuredReferenceHeadersColumnSpecified()
		{
			var reference = new StructuredReference($"{EpplusExcelDataProviderTest.TableName}[[#Headers],[Header2]]");
			var result = this.DataProviderWithDataAndHeaders.ResolveStructuredReference(reference, "Sheet1", 15, 7);
			Assert.AreEqual(3, result.Address.Start.Row);
			Assert.AreEqual(4, result.Address.Start.Column);
			Assert.AreEqual(3, result.Address.End.Row);
			Assert.AreEqual(4, result.Address.End.Column);
			Assert.AreEqual(EpplusExcelDataProviderTest.Header2, result.GetOffset(0, 0));
		}

		[TestMethod]
		public void ResolveStructuredReferenceHeadersColumnRangeSpecified()
		{
			var reference = new StructuredReference($"{EpplusExcelDataProviderTest.TableName}[[#Headers],[Header2]:[Header4]]");
			var result = this.DataProviderWithDataAndHeaders.ResolveStructuredReference(reference, "Sheet1", 15, 7);
			Assert.AreEqual(3, result.Address.Start.Row);
			Assert.AreEqual(4, result.Address.Start.Column);
			Assert.AreEqual(3, result.Address.End.Row);
			Assert.AreEqual(6, result.Address.End.Column);
			Assert.AreEqual(EpplusExcelDataProviderTest.Header2, result.GetOffset(0, 0));
			Assert.AreEqual(EpplusExcelDataProviderTest.Header3, result.GetOffset(0, 1));
			Assert.AreEqual(EpplusExcelDataProviderTest.Header4, result.GetOffset(0, 2));
		}
		#endregion

		#region Totals Tests
		[TestMethod]
		public void ResolveStructuredReferenceTotals()
		{
			var reference = new StructuredReference($"{EpplusExcelDataProviderTest.TableName}[#Totals]");
			var result = this.DataProviderWithDataAndTotals.ResolveStructuredReference(reference, "Sheet1", 15, 7);
			Assert.AreEqual(10, result.Address.Start.Row);
			Assert.AreEqual(3, result.Address.Start.Column);
			Assert.AreEqual(10, result.Address.End.Row);
			Assert.AreEqual(6, result.Address.End.Column);
			Assert.AreEqual("h1_t", result.GetOffset(0, 0));
			Assert.AreEqual("h2_t", result.GetOffset(0, 1));
			Assert.AreEqual("h3_t", result.GetOffset(0, 2));
			Assert.AreEqual("h4_t", result.GetOffset(0, 3));
		}

		[TestMethod]
		public void ResolveStructuredReferenceTotalsColumnSpecified()
		{
			var reference = new StructuredReference($"{EpplusExcelDataProviderTest.TableName}[[#Totals],[Header2]]");
			var result = this.DataProviderWithDataHeadersAndTotals.ResolveStructuredReference(reference, "Sheet1", 15, 7);
			Assert.AreEqual(10, result.Address.Start.Row);
			Assert.AreEqual(4, result.Address.Start.Column);
			Assert.AreEqual(10, result.Address.End.Row);
			Assert.AreEqual(4, result.Address.End.Column);
			Assert.AreEqual("h2_t", result.GetOffset(0, 0));
		}

		[TestMethod]
		public void ResolveStructuredReferenceTotalsColumnRangeSpecified()
		{
			var reference = new StructuredReference($"{EpplusExcelDataProviderTest.TableName}[[#Totals],[Header2]:[Header3]");
			var result = this.DataProviderWithDataHeadersAndTotals.ResolveStructuredReference(reference, "Sheet1", 15, 7);
			Assert.AreEqual(10, result.Address.Start.Row);
			Assert.AreEqual(4, result.Address.Start.Column);
			Assert.AreEqual(10, result.Address.End.Row);
			Assert.AreEqual(5, result.Address.End.Column);
			Assert.AreEqual("h2_t", result.GetOffset(0, 0));
			Assert.AreEqual("h3_t", result.GetOffset(0, 1));
		}

		public void ResolveStructuredReferenceTotalsThatDoNotExist()
		{
			var reference = new StructuredReference($"{EpplusExcelDataProviderTest.TableName}[#Totals]");
			var result = this.DataProviderWithDataAndHeaders.ResolveStructuredReference(reference, "Sheet1", 15, 7);
			Assert.AreEqual("#REF!", result.Address.Address);
		}

		[TestMethod]
		public void ResolveStructuredReferenceTotalsThatDoNotExistColumnSpecified()
		{
			var reference = new StructuredReference($"{EpplusExcelDataProviderTest.TableName}[[#Totals],[Header2]]");
			var result = this.DataProviderWithDataAndHeaders.ResolveStructuredReference(reference, "Sheet1", 15, 7);
			Assert.AreEqual("#REF!", result.Address.Address);
		}

		[TestMethod]
		public void ResolveStructuredReferenceTotalsThatDoNotExistColumnRangeSpecified()
		{
			var reference = new StructuredReference($"{EpplusExcelDataProviderTest.TableName}[[#Totals],[Header2]:[Header3]");
			var result = this.DataProviderWithDataAndHeaders.ResolveStructuredReference(reference, "Sheet1", 15, 7);
			Assert.AreEqual("#REF!", result.Address.Address);
		}
		#endregion

		#region All Tests
		[TestMethod]
		public void ResolveStructuredReferenceAll()
		{
			var reference = new StructuredReference($"{EpplusExcelDataProviderTest.TableName}[#All]");
			var result = this.DataProviderWithDataHeadersAndTotals.ResolveStructuredReference(reference, "Sheet1", 15, 7);
			Assert.AreEqual(3, result.Address.Start.Row);
			Assert.AreEqual(3, result.Address.Start.Column);
			Assert.AreEqual(10, result.Address.End.Row);
			Assert.AreEqual(6, result.Address.End.Column);
			Assert.AreEqual(EpplusExcelDataProviderTest.Header1, result.GetOffset(0, 0));
			Assert.AreEqual(EpplusExcelDataProviderTest.Header2, result.GetOffset(0, 1));
			Assert.AreEqual(EpplusExcelDataProviderTest.Header3, result.GetOffset(0, 2));
			Assert.AreEqual(EpplusExcelDataProviderTest.Header4, result.GetOffset(0, 3));
			Assert.AreEqual("h1_r1", result.GetOffset(1, 0));
			Assert.AreEqual("h1_r2", result.GetOffset(2, 0));
			Assert.AreEqual("h1_r3", result.GetOffset(3, 0));
			Assert.AreEqual("h1_r4", result.GetOffset(4, 0));
			Assert.AreEqual("h1_r5", result.GetOffset(5, 0));
			Assert.AreEqual("h1_r6", result.GetOffset(6, 0));
			Assert.AreEqual("h4_r6", result.GetOffset(6, 3));
			Assert.AreEqual("h1_t", result.GetOffset(7, 0));
			Assert.AreEqual("h4_t", result.GetOffset(7, 3));
		}

		[TestMethod]
		public void ResolveStructuredReferenceAllColumnSpecified()
		{
			var reference = new StructuredReference($"{EpplusExcelDataProviderTest.TableName}[[#All],[Header2]]");
			var result = this.DataProviderWithDataHeadersAndTotals.ResolveStructuredReference(reference, "Sheet1", 15, 7);
			Assert.AreEqual(3, result.Address.Start.Row);
			Assert.AreEqual(4, result.Address.Start.Column);
			Assert.AreEqual(10, result.Address.End.Row);
			Assert.AreEqual(4, result.Address.End.Column);
			Assert.AreEqual(EpplusExcelDataProviderTest.Header2, result.GetOffset(0, 0));
			Assert.AreEqual("h2_r1", result.GetOffset(1, 0));
			Assert.AreEqual("h2_r2", result.GetOffset(2, 0));
			Assert.AreEqual("h2_r3", result.GetOffset(3, 0));
			Assert.AreEqual("h2_r4", result.GetOffset(4, 0));
			Assert.AreEqual("h2_r5", result.GetOffset(5, 0));
			Assert.AreEqual("h2_r6", result.GetOffset(6, 0));
			Assert.AreEqual("h2_t", result.GetOffset(7, 0));
		}

		[TestMethod]
		public void ResolveStructuredReferenceAllColumnRangeSpecified()
		{
			var reference = new StructuredReference($"{EpplusExcelDataProviderTest.TableName}[[#All],[Header2]:[Header3]");
			var result = this.DataProviderWithDataHeadersAndTotals.ResolveStructuredReference(reference, "Sheet1", 15, 7);
			Assert.AreEqual(3, result.Address.Start.Row);
			Assert.AreEqual(4, result.Address.Start.Column);
			Assert.AreEqual(10, result.Address.End.Row);
			Assert.AreEqual(5, result.Address.End.Column);
			Assert.AreEqual(EpplusExcelDataProviderTest.Header2, result.GetOffset(0, 0));
			Assert.AreEqual(EpplusExcelDataProviderTest.Header3, result.GetOffset(0, 1));
			Assert.AreEqual("h2_r1", result.GetOffset(1, 0));
			Assert.AreEqual("h2_r2", result.GetOffset(2, 0));
			Assert.AreEqual("h2_r3", result.GetOffset(3, 0));
			Assert.AreEqual("h2_r4", result.GetOffset(4, 0));
			Assert.AreEqual("h2_r5", result.GetOffset(5, 0));
			Assert.AreEqual("h2_r6", result.GetOffset(6, 0));
			Assert.AreEqual("h2_t", result.GetOffset(7, 0));
			Assert.AreEqual("h3_r1", result.GetOffset(1, 1));
			Assert.AreEqual("h3_r2", result.GetOffset(2, 1));
			Assert.AreEqual("h3_r3", result.GetOffset(3, 1));
			Assert.AreEqual("h3_r4", result.GetOffset(4, 1));
			Assert.AreEqual("h3_r5", result.GetOffset(5, 1));
			Assert.AreEqual("h3_r6", result.GetOffset(6, 1));
			Assert.AreEqual("h3_t", result.GetOffset(7, 1));
		}

		[TestMethod]
		public void ResolveStructuredReferenceAllHeadersDoNotExist()
		{
			var reference = new StructuredReference($"{EpplusExcelDataProviderTest.TableName}[#All]");
			var result = this.DataProviderWithDataAndTotals.ResolveStructuredReference(reference, "Sheet1", 15, 7);
			Assert.AreEqual(4, result.Address.Start.Row);
			Assert.AreEqual(3, result.Address.Start.Column);
			Assert.AreEqual(10, result.Address.End.Row);
			Assert.AreEqual(6, result.Address.End.Column);
			Assert.AreEqual("h1_r1", result.GetOffset(0, 0));
			Assert.AreEqual("h1_r2", result.GetOffset(1, 0));
			Assert.AreEqual("h1_r3", result.GetOffset(2, 0));
			Assert.AreEqual("h1_r4", result.GetOffset(3, 0));
			Assert.AreEqual("h1_r5", result.GetOffset(4, 0));
			Assert.AreEqual("h1_r6", result.GetOffset(5, 0));
			Assert.AreEqual("h4_r6", result.GetOffset(5, 3));
			Assert.AreEqual("h1_t", result.GetOffset(6, 0));
			Assert.AreEqual("h4_t", result.GetOffset(6, 3));
		}

		[TestMethod]
		public void ResolveStructuredReferenceAllTotalsDoNotExist()
		{
			var reference = new StructuredReference($"{EpplusExcelDataProviderTest.TableName}[#All]");
			var result = this.DataProviderWithDataAndHeaders.ResolveStructuredReference(reference, "Sheet1", 15, 7);
			Assert.AreEqual(3, result.Address.Start.Row);
			Assert.AreEqual(3, result.Address.Start.Column);
			Assert.AreEqual(9, result.Address.End.Row);
			Assert.AreEqual(6, result.Address.End.Column);
			Assert.AreEqual(EpplusExcelDataProviderTest.Header1, result.GetOffset(0, 0));
			Assert.AreEqual(EpplusExcelDataProviderTest.Header2, result.GetOffset(0, 1));
			Assert.AreEqual(EpplusExcelDataProviderTest.Header3, result.GetOffset(0, 2));
			Assert.AreEqual(EpplusExcelDataProviderTest.Header4, result.GetOffset(0, 3));
			Assert.AreEqual("h1_r1", result.GetOffset(1, 0));
			Assert.AreEqual("h1_r2", result.GetOffset(2, 0));
			Assert.AreEqual("h1_r3", result.GetOffset(3, 0));
			Assert.AreEqual("h1_r4", result.GetOffset(4, 0));
			Assert.AreEqual("h1_r5", result.GetOffset(5, 0));
			Assert.AreEqual("h1_r6", result.GetOffset(6, 0));
			Assert.AreEqual("h4_r6", result.GetOffset(6, 3));
		}
		#endregion
		#endregion

		#region Helper Methods
		private void BuildTableHeaders(ExcelWorksheet worksheet)
		{
			worksheet.Cells[3, 3].Value = EpplusExcelDataProviderTest.Header1;
			worksheet.Cells[3, 4].Value = EpplusExcelDataProviderTest.Header2;
			worksheet.Cells[3, 5].Value = EpplusExcelDataProviderTest.Header3;
			worksheet.Cells[3, 6].Value = EpplusExcelDataProviderTest.Header4;
		}

		private void BuildTableTotals(ExcelWorksheet worksheet)
		{
			worksheet.Cells[10, 3].Value = "h1_t";
			worksheet.Cells[10, 4].Value = "h2_t";
			worksheet.Cells[10, 5].Value = "h3_t";
			worksheet.Cells[10, 6].Value = "h4_t";
		}

		private void BuildTableData(ExcelWorksheet worksheet)
		{
			worksheet.Cells[4, 3].Value = "h1_r1";
			worksheet.Cells[4, 4].Value = "h2_r1";
			worksheet.Cells[4, 5].Value = "h3_r1";
			worksheet.Cells[4, 6].Value = "h4_r1";
			worksheet.Cells[5, 3].Value = "h1_r2";
			worksheet.Cells[5, 4].Value = "h2_r2";
			worksheet.Cells[5, 5].Value = "h3_r2";
			worksheet.Cells[5, 6].Value = "h4_r2";
			worksheet.Cells[6, 3].Value = "h1_r3";
			worksheet.Cells[6, 4].Value = "h2_r3";
			worksheet.Cells[6, 5].Value = "h3_r3";
			worksheet.Cells[6, 6].Value = "h4_r3";
			worksheet.Cells[7, 3].Value = "h1_r4";
			worksheet.Cells[7, 4].Value = "h2_r4";
			worksheet.Cells[7, 5].Value = "h3_r4";
			worksheet.Cells[7, 6].Value = "h4_r4";
			worksheet.Cells[8, 3].Value = "h1_r5";
			worksheet.Cells[8, 4].Value = "h2_r5";
			worksheet.Cells[8, 5].Value = "h3_r5";
			worksheet.Cells[8, 6].Value = "h4_r5";
			worksheet.Cells[9, 3].Value = "h1_r6";
			worksheet.Cells[9, 4].Value = "h2_r6";
			worksheet.Cells[9, 5].Value = "h3_r6";
			worksheet.Cells[9, 6].Value = "h4_r6";
		}
		#endregion
	}
}
