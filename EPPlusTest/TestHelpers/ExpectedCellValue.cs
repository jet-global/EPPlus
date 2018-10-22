namespace EPPlusTest.TestHelpers
{
	internal struct ExpectedCellValue
	{
		#region Properties
		public string Sheet { get; set; }
		public int Row { get; set; }
		public int Column { get; set; }
		public object Value { get; set; }
		public string Formula { get; set; }
		#endregion

		#region Constructors
		internal ExpectedCellValue(string sheet, int row, int column, object value, string formula = null)
		{
			this.Sheet = sheet;
			this.Row = row;
			this.Column = column;
			this.Value = value;
			this.Formula = formula;
		}
		#endregion
	}
}
