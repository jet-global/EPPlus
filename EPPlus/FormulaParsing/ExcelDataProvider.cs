﻿using System;
using System.Collections.Generic;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml.FormulaParsing
{
	/// <summary>
	/// This class should be implemented to be able to deliver excel data
	/// to the formula parser.
	/// </summary>
	public abstract class ExcelDataProvider : IDisposable
	{
		/// <summary>
		/// A range of cells in a worksheet.
		/// </summary>
		public interface IRangeInfo : IEnumerator<ICellInfo>, IEnumerable<ICellInfo>
		{
			bool IsEmpty { get; }
			bool IsMulti { get; }
			int GetTotalCellCount();
			ExcelAddressBase Address { get; }
			object GetValue(int row, int col);
			object GetOffset(int rowOffset, int colOffset);

			ExcelWorksheet Worksheet { get; }
		}

		/// <summary>
		/// Information and help methods about a cell
		/// </summary>
		public interface ICellInfo
		{
			string Address { get; }
			int Row { get; }
			int Column { get; }
			string Formula { get; }
			object Value { get; }
			double ValueDouble { get; }
			double ValueDoubleLogical { get; }
			bool IsHiddenRow { get; }
			bool IsExcelError { get; }
			IList<Token> Tokens { get; }
		}

		public interface INameInfo
		{
			ulong Id { get; set; }
			string Worksheet { get; set; }
			string Name { get; set; }
			string Formula { get; set; }
			IList<Token> Tokens { get; }
			object Value { get; set; }
		}

		/// <summary>
		/// Returns the names of all worksheet names
		/// </summary>
		/// <returns></returns>
		public abstract ExcelNamedRangeCollection GetWorksheetNames(string worksheet);

		/// <summary>
		/// Returns all defined names in a workbook
		/// </summary>
		/// <returns></returns>
		public abstract ExcelNamedRangeCollection GetWorkbookNameValues();

		/// <summary>
		/// Returns values from the required range.
		/// </summary>
		/// <param name="worksheetName">The name of the worksheet</param>
		/// <param name="row">Row</param>
		/// <param name="column">Column</param>
		/// <param name="address">The reference address</param>
		/// <returns></returns>
		public abstract IRangeInfo GetRange(string worksheetName, int row, int column, string address);

		public abstract INameInfo GetName(string worksheet, string name);

		public abstract IEnumerable<object> GetRangeValues(string address);

		public abstract string GetRangeFormula(string worksheetName, int row, int column);

		public abstract List<Token> GetRangeFormulaTokens(string worksheetName, int row, int column);

		public abstract bool IsRowHidden(string worksheetName, int row);

		/// <summary>
		/// Returns a single cell value
		/// </summary>
		/// <param name="sheetName"></param>
		/// <param name="row"></param>
		/// <param name="col"></param>
		/// <returns></returns>
		public abstract object GetCellValue(string sheetName, int row, int col);

		/// <summary>
		/// Returns the address of the lowest rightmost cell on the worksheet.
		/// </summary>
		/// <param name="worksheet"></param>
		/// <returns></returns>
		public abstract ExcelCellAddress GetDimensionEnd(string worksheet);

		/// <summary>
		/// Use this method to free unmanaged resources.
		/// </summary>
		public abstract void Dispose();

		/// <summary>
		/// Max number of columns in a worksheet that the Excel data provider can handle.
		/// </summary>
		public abstract int ExcelMaxColumns { get; }

		/// <summary>
		/// Max number of rows in a worksheet that the Excel data provider can handle
		/// </summary>
		public abstract int ExcelMaxRows { get; }

		public abstract object GetRangeValue(string worksheetName, int row, int column);

		public abstract string GetFormat(object value, string format);

		public abstract void Reset();

		public abstract IRangeInfo GetRange(string worksheet, int fromRow, int fromCol, int toRow, int toCol);
	}
}
