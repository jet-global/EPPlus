using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Extensions;

namespace EPPlusTest.Extensions
{
	[TestClass]
	public class IEnumerableExtensionsTest
	{
		#region ForEach Tests
		[TestMethod]
		public virtual void ForEach()
		{
			IEnumerable<int> enumerable = new int[] { 1, 2, 3 };
			int expected = 1;
			enumerable.ForEach(actual =>
			{
				Assert.AreEqual(expected, actual);
				expected++;
			});
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public virtual void ForEachNullActionThrowsException()
		{
			IEnumerable<int> enumerable = new int[] { 1, 2, 3 };
			enumerable.ForEach((Action<int>)null);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public virtual void ForEachNullListThrowsException()
		{
			((IEnumerable<int>)null).ForEach(actual => { });
		}

		[TestMethod]
		public virtual void ForEachIndex()
		{
			IEnumerable<int> enumerable = new int[] { 1, 2, 3 };
			int expected = 1;
			int expectedIndex = 0;
			enumerable.ForEach((actual, actualIndex) =>
			{
				Assert.AreEqual(expected, actual);
				Assert.AreEqual(expectedIndex, actualIndex);
				expected++;
				expectedIndex++;
			});
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public virtual void ForEachIndexNullActionThrowsException()
		{
			IEnumerable<int> enumerable = new int[] { 1, 2, 3 };
			enumerable.ForEach((Action<int, int>)null);
		}

		[TestMethod]
		[ExpectedException(typeof(ArgumentNullException))]
		public virtual void ForEachIndexNullListThrowsException()
		{
			((IEnumerable<int>)null).ForEach((actual, actualIndex) => { });
		}
		#endregion
	}
}
