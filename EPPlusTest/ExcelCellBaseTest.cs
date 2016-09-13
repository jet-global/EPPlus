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
    }
}
