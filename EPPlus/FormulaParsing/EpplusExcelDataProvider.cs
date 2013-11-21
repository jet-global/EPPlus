﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace OfficeOpenXml.FormulaParsing
{
    public class EpplusExcelDataProvider : ExcelDataProvider
    {
        public class CellInfo : ICellInfo
        {
            internal ExcelWorksheet _ws;
            CellsStoreEnumerator<object> _values=null;
            int _fromRow, _toRow, _fromCol, _toCol;
            int _cellCount = 0;
            public CellInfo(ExcelWorksheet ws,int fromRow, int fromCol, int toRow, int toCol)                
            {
                _ws = ws;
                _fromRow=fromRow;
                _fromCol=fromCol;
                _toRow=toRow;
                _toCol=toCol;
                _values = new CellsStoreEnumerator<object>(ws._values, _fromRow, _fromCol, _toRow, _toCol);
            }

            public int GetNCells()
            {
                return ((_toRow - _fromRow) + 1) * ((_toCol - _fromCol) + 1);
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

            public new object Value
            {
                get { return _values.Value; }
            }
            
            public double ValueDouble
            {
                get { return ExcelWorksheet.GetValueDouble(_values.Value, true); }
            }
            public double ValueDoubleLogical
            {
                get { return ExcelWorksheet.GetValueDouble(_values.Value, false); }
            }
            public bool IsHiddenRow
            {
                get 
                { 
                    var row=_ws._values.GetValue(_values.Row, 0) as ExcelRow;
                    if(row != null)
                    {
                        return row.Hidden || row.Height==0;
                    }
                    else
                    {
                        return false;
                    }
                }
            }
            public bool IsEmpty
            {
                get 
                {
                    if (_cellCount > 0)
                    {
                        return true;
                    }
                    else if (_values.Next())
                    {
                        _values.Previous();
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
                        if (_values.Next() && _values.Next())
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
                    else if (_cellCount>1)
                    {
                        return true;
                    }
                    return false;
                }
            }
        
            public ICellInfo Current
            {
                get { return this; }
            }

            public void Dispose()
            {
 	            _values=null;
                _ws=null;
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
 	            _values.Init();
            }


            public bool NextCell()
            {
                _cellCount++;
                return _values.MoveNext();
            }

            public IEnumerator<ICellInfo> GetEnumerator()
            {
                return this;
            }

            System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
            {
                return this;
            }


            public IList<Token> Tokens
            {
                get 
                {
                    return _ws._formulaTokens.GetValue(_values.Row, _values.Column);
                }
            }

        }
        public class NameInfo : INameInfo
        {
            public ulong Id { get; set; }
            public string Name  { get; set; }
            public string Formula { get; set; }
            public IList<Token> Tokens { get; internal set; }
            public object Value { get; set; }
        }
        private readonly ExcelPackage _package;
        private ExcelWorksheet _currentWorksheet;
        private RangeAddressFactory _rangeAddressFactory;
        private Dictionary<ulong, INameInfo> _names=new Dictionary<ulong,INameInfo>();
        public EpplusExcelDataProvider(ExcelPackage package)
        {
            _package = package;

            _rangeAddressFactory = new RangeAddressFactory(this);
        }

        public override ExcelNamedRangeCollection GetWorksheetNames()
        {
            return _package.Workbook.Worksheets.First().Names;
        }

        public override ExcelNamedRangeCollection GetWorkbookNameValues()
        {
            return _package.Workbook.Names;
        }
        public override ICellInfo GetRange(string worksheet, int row, int column, string address)
        {
            var addr = new ExcelAddress(worksheet, address);
            if (addr.Table != null)
            {
                addr.SetRCFromTable(_package, new ExcelAddressBase(row, column, row, column));
            }
            SetCurrentWorksheet(addr.WorkSheet); 
            var wsName = string.IsNullOrEmpty(addr.WorkSheet) ? _currentWorksheet.Name : addr.WorkSheet;
            var ws = _package.Workbook.Worksheets[wsName];
            //return new CellsStoreEnumerator<object>(ws._values, addr._fromRow, addr._fromCol, addr._toRow, addr._toCol);
            return new CellInfo(ws, addr._fromRow, addr._fromCol, addr._toRow, addr._toCol);
        }
        public override INameInfo GetName(string worksheet, string name)
        {
            ExcelNamedRange nameItem;
            ulong id;            
            ExcelWorksheet ws;
            if (string.IsNullOrEmpty(worksheet))
            {
                nameItem = _package._workbook.Names[name];
                ws = null;
            }
            else
            {
                ws = _package._workbook.Worksheets[worksheet];
                nameItem = _package._workbook.Names[name];
            }

            id = ExcelAddressBase.GetCellID(nameItem.LocalSheetId, nameItem.Index, 0);

            if (_names.ContainsKey(id))
            {
                return _names[id];
            }
            else
            {
                var ni=new NameInfo(){
                    Id=id,
                    Name=name,
                    Formula=nameItem.Formula
                };
                if (nameItem._fromRow > 0)
                {
                    ni.Value = new CellInfo(nameItem.Worksheet ?? ws, nameItem._fromRow, nameItem._fromCol, nameItem._toRow, nameItem._toCol);
                }
                else
                {
                    ni.Value= nameItem.Value;
                }
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
            return (new CellsStoreEnumerator<object>(ws._values, addr._fromRow, addr._fromCol, addr._toRow, addr._toCol));
        }


        public object GetValue(int row, int column)
        {
            return _currentWorksheet._values.GetValue(row, column);
        }

        public bool IsMerged(int row, int column)
        {
            return _currentWorksheet._flags.GetFlagValue(row, column, CellFlags.Merged);
        }

        public bool IsHidden(int row, int column)
        {
            return _currentWorksheet.Column(column).Hidden || _currentWorksheet.Column(column).Width == 0 ||
                   _currentWorksheet.Row(row).Hidden || _currentWorksheet.Row(column).Height == 0;
        }

        public override object GetCellValue(string sheetName, int row, int col)
        {
            return _package.Workbook.Worksheets[sheetName]._values.GetValue(row, col);
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

        //public override void SetCellValue(string address, object value)
        //{
        //    var addressInfo = ExcelAddressInfo.Parse(address);
        //    var ra = _rangeAddressFactory.Create(address);
        //    SetCurrentWorksheet(addressInfo);
        //    //var valueInfo = (ICalcEngineValueInfo)_currentWorksheet;
        //    //valueInfo.SetFormulaValue(ra.FromRow + 1, ra.FromCol + 1, value);
        //    _currentWorksheet.Cells[ra.FromRow + 1, ra.FromCol + 1].Value = value;
        //}

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
    }
}
    