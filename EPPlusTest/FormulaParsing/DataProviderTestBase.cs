using OfficeOpenXml;

namespace EPPlusTest.FormulaParsing
{
	public class DataProviderTestBase
	{
		#region Constants
		protected const string TableName = "MyTable";
		protected const string Header1 = "Header1";
		protected const string Header2 = "Header2";
		protected const string Header3 = "Header3";
		protected const string Header4 = "Header4";
		#endregion

		#region Helper Methods
		protected void BuildTableHeaders(ExcelWorksheet worksheet)
		{
			worksheet.Cells[3, 3].Value = DataProviderTestBase.Header1;
			worksheet.Cells[3, 4].Value = DataProviderTestBase.Header2;
			worksheet.Cells[3, 5].Value = DataProviderTestBase.Header3;
			worksheet.Cells[3, 6].Value = DataProviderTestBase.Header4;
		}

		protected void BuildTableTotals(ExcelWorksheet worksheet)
		{
			worksheet.Cells[10, 3].Value = "h1_t";
			worksheet.Cells[10, 4].Value = "h2_t";
			worksheet.Cells[10, 5].Value = "h3_t";
			worksheet.Cells[10, 6].Value = "h4_t";
		}

		protected void BuildTableData(ExcelWorksheet worksheet, bool asNumbers = false)
		{
			worksheet.Cells[4, 3].Value = asNumbers ? 11 : (object)"h1_r1";
			worksheet.Cells[4, 4].Value = asNumbers ? 12 : (object)"h2_r1";
			worksheet.Cells[4, 5].Value = asNumbers ? 13 : (object)"h3_r1";
			worksheet.Cells[4, 6].Value = asNumbers ? 14 : (object)"h4_r1";
			worksheet.Cells[5, 3].Value = asNumbers ? 21 : (object)"h1_r2";
			worksheet.Cells[5, 4].Value = asNumbers ? 22 : (object)"h2_r2";
			worksheet.Cells[5, 5].Value = asNumbers ? 23 : (object)"h3_r2";
			worksheet.Cells[5, 6].Value = asNumbers ? 24 : (object)"h4_r2";
			worksheet.Cells[6, 3].Value = asNumbers ? 31 : (object)"h1_r3";
			worksheet.Cells[6, 4].Value = asNumbers ? 32 : (object)"h2_r3";
			worksheet.Cells[6, 5].Value = asNumbers ? 33 : (object)"h3_r3";
			worksheet.Cells[6, 6].Value = asNumbers ? 34 : (object)"h4_r3";
			worksheet.Cells[7, 3].Value = asNumbers ? 41 : (object)"h1_r4";
			worksheet.Cells[7, 4].Value = asNumbers ? 42 : (object)"h2_r4";
			worksheet.Cells[7, 5].Value = asNumbers ? 43 : (object)"h3_r4";
			worksheet.Cells[7, 6].Value = asNumbers ? 44 : (object)"h4_r4";
			worksheet.Cells[8, 3].Value = asNumbers ? 51 : (object)"h1_r5";
			worksheet.Cells[8, 4].Value = asNumbers ? 52 : (object)"h2_r5";
			worksheet.Cells[8, 5].Value = asNumbers ? 53 : (object)"h3_r5";
			worksheet.Cells[8, 6].Value = asNumbers ? 54 : (object)"h4_r5";
			worksheet.Cells[9, 3].Value = asNumbers ? 61 : (object)"h1_r6";
			worksheet.Cells[9, 4].Value = asNumbers ? 62 : (object)"h2_r6";
			worksheet.Cells[9, 5].Value = asNumbers ? 63 : (object)"h3_r6";
			worksheet.Cells[9, 6].Value = asNumbers ? 64 : (object)"h4_r6";
		}
		#endregion
	}
}
