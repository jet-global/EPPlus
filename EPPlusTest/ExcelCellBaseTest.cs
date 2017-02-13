using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest
{
    [TestClass]
    public class ExcelCellBaseTest
    {
        #region GetAddress(int FromRow, int FromColumn, int ToRow, int ToColumn) Tests
        [TestMethod]
        public void GetAddressSingleCell()
        {
            var address = ExcelCellBase.GetAddress(3, 3, 3, 3);
            Assert.AreEqual("C3", address);
        }

        [TestMethod]
        public void GetAddressMultiCell()
        {
            var address = ExcelCellBase.GetAddress(3, 3, 4, 4);
            Assert.AreEqual("C3:D4", address);
        }

        [TestMethod]
        public void GetAddressMultiCellToRowIsMaxRows()
        {
            var address = ExcelCellBase.GetAddress(1, 1, ExcelPackage.MaxRows, 4);
            Assert.AreEqual("A:D", address);
        }

        [TestMethod]
        public void GetAddressMultiCellToRowIsGreaterThanMaxRows()
        {
            var address = ExcelCellBase.GetAddress(1, 1, ExcelPackage.MaxRows + 1, 4);
            Assert.AreEqual("A:D", address);
        }

        [TestMethod]
        public void GetAddressMultiCellToColumnIsMaxColumns()
        {
            var address = ExcelCellBase.GetAddress(1, 1, 4, ExcelPackage.MaxColumns);
            Assert.AreEqual("1:4", address);
        }

        [TestMethod]
        public void GetAddressMultiCellToColumnIsGreaterThanMaxColumns()
        {
            var address = ExcelCellBase.GetAddress(1, 1, 4, ExcelPackage.MaxColumns + 1);
            Assert.AreEqual("1:4", address);
        }
        #endregion

        #region GetAddress(int FromRow, int FromColumn, int ToRow, int ToColumn, bool FixedFromRow, bool FixedFromColumn, bool FixedToRow, bool FixedToColumn) Tests
        [TestMethod]
        public void GetAddressSingleCellAbsoluteVariation()
        {
            var address = ExcelCellBase.GetAddress(3, 3, 3, 3, true, true, true, true);
            Assert.AreEqual("$C$3", address);
        }

        [TestMethod]
        public void GetAddressMultiCellAbsoluteVariation()
        {
            var address = ExcelCellBase.GetAddress(3, 3, 4, 4, true, true, true, true);
            Assert.AreEqual("$C$3:$D$4", address);
        }

        [TestMethod]
        public void GetAddressMultiCellToRowIsMaxRowsAbsoluteVariation()
        {
            var address = ExcelCellBase.GetAddress(1, 1, ExcelPackage.MaxRows, 4, true, true, true, true);
            Assert.AreEqual("$A:$D", address);
        }

        [TestMethod]
        public void GetAddressMultiCellToRowIsGreaterThanMaxRowsAbsoluteVariation()
        {
            var address = ExcelCellBase.GetAddress(1, 1, ExcelPackage.MaxRows + 1, 4, true, true, true, true);
            Assert.AreEqual("$A:$D", address);
        }

        [TestMethod]
        public void GetAddressMultiCellToColumnIsMaxColumnsAbsoluteVariation()
        {
            var address = ExcelCellBase.GetAddress(1, 1, 4, ExcelPackage.MaxColumns, true, true, true, true);
            Assert.AreEqual("$1:$4", address);
        }

        [TestMethod]
        public void GetAddressMultiCellToColumnIsGreaterThanMaxColumnsAbsoluteVariation()
        {
            var address = ExcelCellBase.GetAddress(1, 1, 4, ExcelPackage.MaxColumns + 1, true, true, true, true);
            Assert.AreEqual("$1:$4", address);
        }
        #endregion

        #region GetRowCol Tests
        [TestMethod]
        public void GetRowColFullAddress()
        {
            int row, col;
            bool fixedRow, fixedCol;
            ExcelCellBase.GetRowCol("C3", out row, out col, true, out fixedRow, out fixedCol);
            Assert.AreEqual(3, row);
            Assert.AreEqual(3, col);
            Assert.IsFalse(fixedRow);
            Assert.IsFalse(fixedCol);
        }

        [TestMethod]
        public void GetRowColFullAddressColumnAbsolute()
        {
            int row, col;
            bool fixedRow, fixedCol;
            ExcelCellBase.GetRowCol("$C3", out row, out col, true, out fixedRow, out fixedCol);
            Assert.AreEqual(3, row);
            Assert.AreEqual(3, col);
            Assert.IsFalse(fixedRow);
            Assert.IsTrue(fixedCol);
        }

        [TestMethod]
        public void GetRowColFullAddressRowAbsolute()
        {
            int row, col;
            bool fixedRow, fixedCol;
            ExcelCellBase.GetRowCol("C$3", out row, out col, true, out fixedRow, out fixedCol);
            Assert.AreEqual(3, row);
            Assert.AreEqual(3, col);
            Assert.IsTrue(fixedRow);
            Assert.IsFalse(fixedCol);
        }

        [TestMethod]
        public void GetRowColFullAddressAbsolute()
        {
            int row, col;
            bool fixedRow, fixedCol;
            ExcelCellBase.GetRowCol("$C$3", out row, out col, true, out fixedRow, out fixedCol);
            Assert.AreEqual(3, row);
            Assert.AreEqual(3, col);
            Assert.IsTrue(fixedRow);
            Assert.IsTrue(fixedCol);
        }

        [TestMethod]
        public void GetRowColJustColumn()
        {
            int row, col;
            bool fixedRow, fixedCol;
            ExcelCellBase.GetRowCol("C", out row, out col, true, out fixedRow, out fixedCol);
            Assert.AreEqual(0, row);
            Assert.AreEqual(3, col);
            Assert.IsFalse(fixedRow);
            Assert.IsFalse(fixedCol);
        }

        [TestMethod]
        public void GetRowColJustColumnAbsolute()
        {
            int row, col;
            bool fixedRow, fixedCol;
            ExcelCellBase.GetRowCol("$C", out row, out col, true, out fixedRow, out fixedCol);
            Assert.AreEqual(0, row);
            Assert.AreEqual(3, col);
            Assert.IsFalse(fixedRow);
            Assert.IsTrue(fixedCol);
        }

        [TestMethod]
        public void GetRowColJustRow()
        {
            int row, col;
            bool fixedRow, fixedCol;
            ExcelCellBase.GetRowCol("3", out row, out col, true, out fixedRow, out fixedCol);
            Assert.AreEqual(3, row);
            Assert.AreEqual(0, col);
            Assert.IsFalse(fixedRow);
            Assert.IsFalse(fixedCol);
        }

        [TestMethod]
        public void GetRowColJustRowAbsolute()
        {
            int row, col;
            bool fixedRow, fixedCol;
            ExcelCellBase.GetRowCol("$3", out row, out col, true, out fixedRow, out fixedCol);
            Assert.AreEqual(3, row);
            Assert.AreEqual(0, col);
            Assert.IsTrue(fixedRow);
            Assert.IsFalse(fixedCol);
        }
        #endregion
    }
}
