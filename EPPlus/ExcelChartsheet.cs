using System;
using System.Xml;
using OfficeOpenXml.Drawing.Chart;

namespace OfficeOpenXml
{
	/// <summary>
	/// Represents an Excel Chartsheet and provides access to its properties and methods
	/// </summary>
	public class ExcelChartsheet : ExcelWorksheet
	{
		#region Properties
		/// <summary>
		/// Gets the Chart that this ChartSheet is dedicated to.
		/// </summary>
		public ExcelChart Chart
		{
			get
			{
				return (ExcelChart)Drawings[0];
			}
		}
		#endregion

		#region Constructors
		/// <summary>
		/// Initialize a new Chart Sheet.
		/// </summary>
		/// <param name="ns">The namespace manager.</param>
		/// <param name="pck">The package that this chartsheet is contained in.</param>
		/// <param name="relID">The relID of the new sheet.</param>
		/// <param name="uriWorksheet">The URI to save the new sheet to.</param>
		/// <param name="sheetName">The name of the new sheet.</param>
		/// <param name="sheetID">The ID of the new sheet.</param>
		/// <param name="positionID">The index of the sheet in the Worksheets collection.</param>
		/// <param name="hidden">The <see cref="eWorkSheetHidden"/> state of the worksheet.</param>
		/// <param name="chartType">The <see cref="eChartType"/> of the new chart.</param>
		public ExcelChartsheet(XmlNamespaceManager ns, ExcelPackage pck, string relID, Uri uriWorksheet, string sheetName, int sheetID, int positionID, eWorkSheetHidden hidden, eChartType chartType) :
			base(ns, pck, relID, uriWorksheet, sheetName, sheetID, positionID, hidden)
		{
			this.Drawings.AddChart("Chart 1", chartType);
		}

		/// <summary>
		/// Initialize a new Chart Sheet.
		/// </summary>
		/// <param name="ns">The namespace manager.</param>
		/// <param name="pck">The package that this chartsheet is contained in.</param>
		/// <param name="relID">The relID of the new sheet.</param>
		/// <param name="uriWorksheet">The URI to save the new sheet to.</param>
		/// <param name="sheetName">The name of the new sheet.</param>
		/// <param name="sheetID">The ID of the new sheet.</param>
		/// <param name="positionID">The index of the sheet in the Worksheets collection.</param>
		/// <param name="hidden">The <see cref="eWorkSheetHidden"/> state of the worksheet.</param>
		public ExcelChartsheet(XmlNamespaceManager ns, ExcelPackage pck, string relID, Uri uriWorksheet, string sheetName, int sheetID, int positionID, eWorkSheetHidden hidden) :
			base(ns, pck, relID, uriWorksheet, sheetName, sheetID, positionID, hidden)
		{ }
		#endregion
	}
}
