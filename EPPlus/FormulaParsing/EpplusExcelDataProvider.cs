using System;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.FormulaParsing
{
	public class EpplusExcelDataProvider : ExcelDataProvider
	{
		public class RangeInfo : IRangeInfo
		{
			internal ExcelWorksheet _ws;
			ICellStoreEnumerator<ExcelCoreValue> _values = null;
			int _fromRow, _toRow, _fromCol, _toCol;
			int _cellCount = 0;
			ExcelAddress _address;
			ICellInfo _cell;

			public RangeInfo(ExcelWorksheet ws, int fromRow, int fromCol, int toRow, int toCol)
			{
				_ws = ws;
				_fromRow = fromRow;
				_fromCol = fromCol;
				_toRow = toRow;
				_toCol = toCol;
				_address = new ExcelAddress(_fromRow, _fromCol, _toRow, _toCol);
				_address._ws = ws.Name;
				_values = ws._values.GetEnumerator(_fromRow, _fromCol, _toRow, _toCol);
				_cell = new CellInfo(_ws, _values);
			}

			public RangeInfo(ExcelWorksheet ws, ExcelAddress address)
			{
				_ws = ws;
				_fromRow = address._fromRow;
				_fromCol = address._fromCol;
				_toRow = address._toRow;
				_toCol = address._toCol;
				_address = address;
				_address._ws = ws.Name;
				_values = ws._values.GetEnumerator(_fromRow, _fromCol, _toRow, _toCol);
				_cell = new CellInfo(_ws, _values);
			}

			public int GetTotalCellCount()
			{
				return ((_toRow - _fromRow) + 1) * ((_toCol - _fromCol) + 1);
			}

			public bool IsEmpty
			{
				get
				{
					if (_cellCount > 0)
					{
						return false;
					}
					else if (_values.MoveNext())
					{
						_values.Reset();
						return false;
					}
					else
					{
						return true;
					}
				}
			}
			public bool IsMulti
			{
				get
				{
					if (_cellCount == 0)
					{
						if (_values.MoveNext() && _values.MoveNext())
						{
							_values.Reset();
							return true;
						}
						else
						{
							_values.Reset();
							return false;
						}
					}
					else if (_cellCount > 1)
					{
						return true;
					}
					return false;
				}
			}

			public ICellInfo Current
			{
				get { return _cell; }
			}

			public ExcelWorksheet Worksheet
			{
				get { return _ws; }
			}

			public void Dispose()
			{
				//_values = null;
				//_ws = null;
				//_cell = null;
			}

			object System.Collections.IEnumerator.Current
			{
				get
				{
					return this;
				}
			}

			public bool MoveNext()
			{
				_cellCount++;
				return _values.MoveNext();
			}

			public void Reset()
			{
				_values.Reset();
			}


			public bool NextCell()
			{
				_cellCount++;
				return _values.MoveNext();
			}

			public IEnumerator<ICellInfo> GetEnumerator()
			{
				Reset();
				return this;
			}

			System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
			{
				return this;
			}

			public ExcelAddress Address
			{
				get { return _address; }
			}

			public object GetValue(int row, int col)
			{
				return _ws.GetValue(row, col);
			}

			public object GetOffset(int rowOffset, int colOffset)
			{
				if (_values.Row < _fromRow || _values.Column < _fromCol)
				{
					return _ws.GetValue(_fromRow + rowOffset, _fromCol + colOffset);
				}
				else
				{
					return _ws.GetValue(_values.Row + rowOffset, _values.Column + colOffset);
				}
			}

			/// <summary>
			/// Return a list containing the value of each and every cell in this <see cref="RangeInfo"/>.
			/// This function exists because iterating normally over a <see cref="ExcelDataProvider.IRangeInfo"/>
			/// and looking at each <see cref="ExcelDataProvider.ICellInfo"/> will not include cells that have not been set
			/// (ie: cells that have been empty since the workbook's creation). This function works around that issue.
			/// </summary>
			/// <returns>The value of every cell in this <see cref="RangeInfo"/>.</returns>
			public IEnumerable<object> AllValues()
			{
				var values = new List<object>();
				if (this.Address.Addresses?.Any() == true)
				{
					foreach (var subAddress in this.Address.Addresses)
					{
						values.AddRange(this.AllValuesInSubAddress(subAddress));
					}
				}
				else
					values.AddRange(this.AllValuesInSubAddress(this.Address));
				return values;
			}

			private IEnumerable<object> AllValuesInSubAddress(ExcelAddress address)
			{
				var values = new List<object>();
				for (var currentRow = address._fromRow; currentRow <= address._toRow; currentRow++)
				{
					for (var currentColumn = address._fromCol; currentColumn <= address._toCol; currentColumn++)
					{
						var value = _ws._values.GetValue(currentRow, currentColumn)._value;
						values.Add(value);
					}
				}
				return values;
			}
		}

		public class CellInfo : ICellInfo
		{
			ExcelWorksheet _ws;
			ICellStoreEnumerator<ExcelCoreValue> _values;
			internal CellInfo(ExcelWorksheet ws, ICellStoreEnumerator<ExcelCoreValue> values)
			{
				_ws = ws;
				_values = values;
			}
			public string Address
			{
				get { return _values.CellAddress; }
			}

			public int Row
			{
				get { return _values.Row; }
			}

			public int Column
			{
				get { return _values.Column; }
			}

			public string Formula
			{
				get
				{
					return _ws.GetFormula(_values.Row, _values.Column);
				}
			}

			public object Value
			{
				get { return _values.Value._value; }
			}

			public double ValueDouble
			{
				get { return ConvertUtil.GetValueDouble(_values.Value._value, true); }
			}
			public double ValueDoubleLogical
			{
				get { return ConvertUtil.GetValueDouble(_values.Value._value, false); }
			}
			public bool IsHiddenRow
			{
				get
				{
					var row = _ws.GetValueInner(_values.Row, 0) as RowInternal;
					if (row != null)
					{
						return row.Hidden || row.Height == 0;
					}
					else
					{
						return false;
					}
				}
			}

			public bool IsExcelError
			{
				get
				{
					var value = _values.Value._value;
					// Parse error string values that were not set by a formula as error values.
					if (value is string stringValue && ExcelErrorValue.Values.TryGetErrorType(stringValue, out _) && string.IsNullOrEmpty(this.Formula))
						return true;
					return value is ExcelErrorValue;
				}
			}

			public IList<Token> Tokens
			{
				get
				{
					return _ws._formulaTokens.GetValue(_values.Row, _values.Column);
				}
			}

		}
		public class NameInfo : ExcelDataProvider.INameInfo
		{
			public ulong Id { get; set; }
			public string Worksheet { get; set; }
			public string Name { get; set; }
			public string Formula { get; set; }
			public IList<Token> Tokens { get; internal set; }
			public object Value { get; set; }
		}

		private readonly ExcelPackage _package;
		private ExcelWorksheet _currentWorksheet;
		private RangeAddressFactory _rangeAddressFactory;
		private Dictionary<ulong, INameInfo> _names = new Dictionary<ulong, INameInfo>();

		public EpplusExcelDataProvider(ExcelPackage package)
		{
			_package = package;

			_rangeAddressFactory = new RangeAddressFactory(this);
		}

		public override ExcelNamedRangeCollection GetWorksheetNames(string worksheet)
		{
			var ws = _package.Workbook.Worksheets[worksheet];
			if (ws != null)
			{
				return ws.Names;
			}
			else
			{
				return null;
			}
		}

		public override ExcelNamedRangeCollection GetWorkbookNameValues()
		{
			return _package.Workbook.Names;
		}

		public override IRangeInfo GetRange(string worksheet, int fromRow, int fromCol, int toRow, int toCol)
		{
			SetCurrentWorksheet(worksheet);
			var wsName = string.IsNullOrEmpty(worksheet) ? _currentWorksheet.Name : worksheet;
			var ws = _package.Workbook.Worksheets[wsName];
			return new RangeInfo(ws, fromRow, fromCol, toRow, toCol);
		}
		public override IRangeInfo GetRange(string worksheet, int row, int column, string address)
		{
			var excelAddress = new ExcelAddress(worksheet, address);
			// External references should not be resolved.
			if (!string.IsNullOrEmpty(excelAddress?.Workbook))
				return null;
			if (excelAddress.IsTableAddress)
				excelAddress.SetRCFromTable(_package, new ExcelAddress(row, column, row, column));
			var wsName = string.IsNullOrEmpty(excelAddress.WorkSheet) ? _currentWorksheet.Name : excelAddress.WorkSheet;
			var ws = _package.Workbook.Worksheets[wsName];
			return new RangeInfo(ws, excelAddress);
		}

		/// <summary>
		/// Returns values from the range defined by the <paramref name="structuredReference"/>.
		/// </summary>
		/// <param name="structuredReference">The <see cref="StructuredReference"/> to resolve.</param>
		/// <param name="originSheet">The sheet referencing the <paramref name="structuredReference"/>.</param>
		/// <param name="originRow">The row referencing the <paramref name="structuredReference"/>.</param>
		/// <param name="originColumn">The column referencing the <paramref name="structuredReference"/>.</param>
		/// <returns>The <see cref="ExcelDataProvider.IRangeInfo"/> containing the referenced data.</returns>
		public override IRangeInfo ResolveStructuredReference(StructuredReference structuredReference, string originSheet, int originRow, int originColumn)
		{
			if (structuredReference == null || !structuredReference.HasValidItemSpecifiers())
				return null;
			var table = _package.Workbook.GetTable(structuredReference.TableName);
			if (table == null)
				return null;
			int startRowPosition = table.Address.Start.Row;
			int endRowPosition = table.Address.End.Row;
			if (structuredReference.ItemSpecifiers.HasFlag(ItemSpecifiers.Data))
			{
				if (table.ShowHeader && !structuredReference.ItemSpecifiers.HasFlag(ItemSpecifiers.Headers))
					startRowPosition++;
				if (table.ShowTotal && !structuredReference.ItemSpecifiers.HasFlag(ItemSpecifiers.Totals))
					endRowPosition--;
			}
			else if (structuredReference.ItemSpecifiers == ItemSpecifiers.ThisRow)
			{
				if (originRow < startRowPosition || originRow > endRowPosition || (originRow == startRowPosition && table.ShowHeader) || (originRow == endRowPosition && table.ShowTotal)) 
					return null;
				startRowPosition = endRowPosition = originRow;
			}
			else if (structuredReference.ItemSpecifiers == ItemSpecifiers.All)
			{
				// Already set.
			}
			else if (structuredReference.ItemSpecifiers == ItemSpecifiers.Headers)
			{
				if (!table.ShowHeader)
					return new RangeInfo(table.WorkSheet, new ExcelAddress(ExcelErrorValue.Values.Ref));
				endRowPosition = startRowPosition;
			}
			else if (structuredReference.ItemSpecifiers == ItemSpecifiers.Totals)
			{
				if (!table.ShowTotal)
					return new RangeInfo(table.WorkSheet, new ExcelAddress(ExcelErrorValue.Values.Ref));
				startRowPosition = endRowPosition;
			}
			int startColumnPosition = table.Address.Start.Column;
			int endColumnPosition = table.Address.End.Column;
			if (!string.IsNullOrEmpty(structuredReference.StartColumn) && !string.IsNullOrEmpty(structuredReference.EndColumn))
			{
				var startColumn = table.Columns[structuredReference.StartColumn];
				if (startColumn != null)
					startColumnPosition += startColumn.Position;
				var endColumn = table.Columns[structuredReference.EndColumn];
				if (endColumn != null)
					endColumnPosition = table.Address.Start.Column + endColumn.Position;
			}
			return new RangeInfo(table.WorkSheet, new ExcelAddress(startRowPosition, startColumnPosition, endRowPosition, endColumnPosition));
		}

		public override INameInfo GetName(string worksheet, string name)
		{
			ExcelNamedRange nameItem;
			ulong id;
			ExcelWorksheet ws;
			if (string.IsNullOrEmpty(worksheet))
			{
				if (_package.Workbook.Names.ContainsKey(name))
				{
					nameItem = _package.Workbook.Names[name];
				}
				else
				{
					return null;
				}
				ws = null;
			}
			else
			{
				ws = _package.Workbook.Worksheets[worksheet];
				if (ws != null && ws.Names.ContainsKey(name))
				{
					nameItem = ws.Names[name];
				}
				else if (_package.Workbook.Names.ContainsKey(name))
				{
					nameItem = _package.Workbook.Names[name];
				}
				else
				{
					return null;
				}
			}
			id = ExcelAddress.GetCellID(nameItem.LocalSheetID, nameItem.Index, 0);
			if (_names.ContainsKey(id))
			{
				return _names[id];
			}
			else
			{
				var ni = new NameInfo()
				{
					Id = id,
					Name = name,
					Worksheet = nameItem.LocalSheet?.Name,
					Formula = nameItem.NameFormula
				};
				var range = nameItem.GetFormulaAsCellRange();
				if (range == null)
					ni.Value = nameItem.NameFormula;
				else
					ni.Value = new RangeInfo(range.Worksheet ?? ws, range);
				_names.Add(id, ni);
				return ni;
			}
		}

		public override IEnumerable<object> GetRangeValues(string address)
		{
			SetCurrentWorksheet(ExcelAddressInfo.Parse(address));
			var addr = new ExcelAddress(address);
			var wsName = string.IsNullOrEmpty(addr.WorkSheet) ? _currentWorksheet.Name : addr.WorkSheet;
			var ws = _package.Workbook.Worksheets[wsName];
			return (IEnumerable<object>)(ws._values.GetEnumerator(addr._fromRow, addr._fromCol, addr._toRow, addr._toCol));
		}


		public object GetValue(int row, int column)
		{
			return _currentWorksheet.GetValueInner(row, column);
		}

		public bool IsMerged(int row, int column)
		{
			//return _currentWorksheet._flags.GetFlagValue(row, column, CellFlags.Merged);
			return _currentWorksheet.MergedCells[row, column] != null;
		}

		public bool IsHidden(int row, int column)
		{
			return _currentWorksheet.Column(column).Hidden || _currentWorksheet.Column(column).Width == 0 ||
					 _currentWorksheet.Row(row).Hidden || _currentWorksheet.Row(column).Height == 0;
		}

		public override object GetCellValue(string sheetName, int row, int col)
		{
			SetCurrentWorksheet(sheetName);
			return _currentWorksheet.GetValueInner(row, col);
		}

		public override ExcelCellAddress GetDimensionEnd(string worksheet)
		{
			ExcelCellAddress address = null;
			try
			{
				address = _package.Workbook.Worksheets[worksheet].Dimension.End;
			}
			catch { }

			return address;
		}

		/// <summary>
		/// Retrieves the <see cref="ExcelPivotTable"/> (if any) at the specified <paramref name="address"/>.
		/// If multiple pivot tables exist within the range, the one that starts closest to 
		/// cell A1 is returned.
		/// </summary>
		/// <param name="address">The <see cref="ExcelAddress"/> to look for a pivot table.</param>
		/// <returns>The pivot table found at the specified address.</returns>
		public override ExcelPivotTable GetPivotTable(ExcelAddress address)
		{
			var worksheet = _package.Workbook.Worksheets[address.WorkSheet];
			var collidingTables = worksheet.PivotTables.Where(pt => pt.Address.Collide(address) != ExcelAddress.eAddressCollition.No);
			var count = collidingTables.Count();
			if (count == 0)
				return null;
			else if (count == 1)
				return collidingTables.First();
			// More than one pivot table matching the specified range was found. 
			// Determine which is first in the sheet (closest to cell A1).
			return collidingTables
				.OrderBy(pt => Math.Sqrt(((pt.Address._fromRow - 1) ^ 2) + ((pt.Address._fromCol - 1) ^ 2)))
				.First();
		}

		private void SetCurrentWorksheet(ExcelAddressInfo addressInfo)
		{
			if (addressInfo.WorksheetIsSpecified)
			{
				_currentWorksheet = _package.Workbook.Worksheets[addressInfo.Worksheet];
			}
			else if (_currentWorksheet == null)
			{
				_currentWorksheet = _package.Workbook.Worksheets.First();
			}
		}

		private void SetCurrentWorksheet(string worksheetName)
		{
			if (!string.IsNullOrEmpty(worksheetName))
			{
				_currentWorksheet = _package.Workbook.Worksheets[worksheetName];
			}
			else
			{
				_currentWorksheet = _package.Workbook.Worksheets.First();
			}

		}

		public override void Dispose()
		{
			_package.Dispose();
		}

		public override int ExcelMaxColumns
		{
			get { return ExcelPackage.MaxColumns; }
		}

		public override int ExcelMaxRows
		{
			get { return ExcelPackage.MaxRows; }
		}

		public override string GetRangeFormula(string worksheetName, int row, int column)
		{
			SetCurrentWorksheet(worksheetName);
			return _currentWorksheet.GetFormula(row, column);
		}

		public override object GetRangeValue(string worksheetName, int row, int column)
		{
			SetCurrentWorksheet(worksheetName);
			return _currentWorksheet.GetValue(row, column);
		}
		public override string GetFormat(object value, string format)
		{
			var styles = _package.Workbook.Styles;
			ExcelNumberFormatXml.ExcelFormatTranslator ft = null;
			foreach (var f in styles.NumberFormats)
			{
				if (f.Format == format)
				{
					ft = f.FormatTranslator;
					break;
				}
			}
			if (ft == null)
			{
				ft = new ExcelNumberFormatXml.ExcelFormatTranslator(format, -1);
			}
			return ExcelRangeBase.FormatValue(value, ft, format, ft.NetFormat);
		}
		public override List<LexicalAnalysis.Token> GetRangeFormulaTokens(string worksheetName, int row, int column)
		{
			return _package.Workbook.Worksheets[worksheetName]._formulaTokens.GetValue(row, column);
		}

		public override bool IsRowHidden(string worksheetName, int row)
		{
			var b = _package.Workbook.Worksheets[worksheetName].Row(row).Height == 0 ||
					  _package.Workbook.Worksheets[worksheetName].Row(row).Hidden;

			return b;
		}

		public override void Reset()
		{
			_names = new Dictionary<ulong, INameInfo>(); //Reset name cache.
		}
	}
}
