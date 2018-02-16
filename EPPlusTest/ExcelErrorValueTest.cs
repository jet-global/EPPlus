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
	public class ExcelErrorValueTest
	{
		#region ExcelErrorValue Test Methods
		#region Create Tests
		[TestMethod]
		public void CreateWithValidErrorType()
		{
			Assert.AreEqual(eErrorType.Value, ExcelErrorValue.Create(eErrorType.Value).Type);
			Assert.AreEqual(eErrorType.Name, ExcelErrorValue.Create(eErrorType.Name).Type);
			Assert.AreEqual(eErrorType.Null, ExcelErrorValue.Create(eErrorType.Null).Type);
			Assert.AreEqual(eErrorType.Num, ExcelErrorValue.Create(eErrorType.Num).Type);
			Assert.AreEqual(eErrorType.Ref, ExcelErrorValue.Create(eErrorType.Ref).Type);
			Assert.AreEqual(eErrorType.Div0, ExcelErrorValue.Create(eErrorType.Div0).Type);
			Assert.AreEqual(eErrorType.NA, ExcelErrorValue.Create(eErrorType.NA).Type);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentException))]
		public void CreateWithInvalidErrorTypeThrowsException()
		{
			ExcelErrorValue.Create(0);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentException))]
		public void CreateWithInvalidDefaultErrorTypeThrowsException()
		{
			ExcelErrorValue.Create(default(eErrorType));
		}
		#endregion

		#region Parse Tests
		[TestMethod]
		public void ParseValidErrorStrings()
		{
			Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Value), ExcelErrorValue.Parse("#VALUE!"));
			Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Name), ExcelErrorValue.Parse("#NAME?"));
			Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Div0), ExcelErrorValue.Parse("#DIV/0!"));
			Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Null), ExcelErrorValue.Parse("#NULL!"));
			Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Num), ExcelErrorValue.Parse("#NUM!"));
			Assert.AreEqual(ExcelErrorValue.Create(eErrorType.Ref), ExcelErrorValue.Parse("#REF!"));
			Assert.AreEqual(ExcelErrorValue.Create(eErrorType.NA), ExcelErrorValue.Parse("#N/A"));
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void ParseEmptyStringThrowsException()
		{
			ExcelErrorValue.Parse(string.Empty);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public void ParseNullThrowsException()
		{
			ExcelErrorValue.Parse(string.Empty);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentException))]
		public void ParseNonErrorStringThrowsException()
		{
			ExcelErrorValue.Parse("not an error");
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentException))]
		public void ParseSimilarErrorStringThrowsException()
		{
			ExcelErrorValue.Parse("#VALUE");
		}
		#endregion
		#endregion

		#region Values Test Methods
		[TestMethod]
		public void TryGetErrorTypeTest()
		{
			Assert.IsTrue(ExcelErrorValue.Values.TryGetErrorType("#VALUE!", out eErrorType errorType));
			Assert.AreEqual(eErrorType.Value, errorType);
			Assert.IsTrue(ExcelErrorValue.Values.TryGetErrorType("#NAME?", out errorType));
			Assert.AreEqual(eErrorType.Name, errorType);
			Assert.IsTrue(ExcelErrorValue.Values.TryGetErrorType("#NULL!", out errorType));
			Assert.AreEqual(eErrorType.Null, errorType);
			Assert.IsTrue(ExcelErrorValue.Values.TryGetErrorType("#N/A", out errorType));
			Assert.AreEqual(eErrorType.NA, errorType);
			Assert.IsTrue(ExcelErrorValue.Values.TryGetErrorType("#NUM!", out errorType));
			Assert.AreEqual(eErrorType.Num, errorType);
			Assert.IsTrue(ExcelErrorValue.Values.TryGetErrorType("#REF!", out errorType));
			Assert.AreEqual(eErrorType.Ref, errorType);
			Assert.IsTrue(ExcelErrorValue.Values.TryGetErrorType("#DIV/0!", out errorType));
			Assert.AreEqual(eErrorType.Div0, errorType);
		}

		[TestMethod]
		public void TryGetErrorTypeInvalidString()
		{
			Assert.IsFalse(ExcelErrorValue.Values.TryGetErrorType("#VALUE", out eErrorType errorType));
			Assert.AreEqual(default(eErrorType), errorType);
			Assert.IsFalse(ExcelErrorValue.Values.TryGetErrorType("#VALUE!asdasd", out errorType));
			Assert.AreEqual(default(eErrorType), errorType);
			Assert.IsFalse(ExcelErrorValue.Values.TryGetErrorType("blah", out errorType));
			Assert.AreEqual(default(eErrorType), errorType);
		}
		#endregion
	}
}
